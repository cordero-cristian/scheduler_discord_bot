#!/bin/python

from datetime import date
import datetime
import pathlib
from pathlib import Path
import pandas as pd
import numpy as np
import discord
import pprint

#Globals Start
today = date.today()

#check to see what day of the week it is, returns an int. 0 is equal to monday, 1 is equal to tuesday...etc
weekDay = today.weekday()

year = today.year

month = today.month

day = today.day

MondayDate = day - weekDay
#get the file path of the dir the program was exc. out of.
dirPath = pathlib.Path().cwd()

#create a standard file name. Add correct file type.
weeklyFileName = str('Schedule_week_of_' + str(year) + '_' + str(month) + '_' + str(MondayDate) + '.xlsx')

#File path to the soon to be Xcel file.
fullFilePath = str(str(dirPath) + "\\" + weeklyFileName)

#assuming you only want appointments between the hours of 1100 and 2000.
timeSlots = list()
timeSlots = ['1PM','2PM','3PM','4PM','5PM','6PM','7PM','8PM','9PM','10PM','11PM']

serverAdmins = list()

#this is here for testing.
#people = list()
#people = ['personA','personB','personC','personD','personE','personF','personG','personH','personI','personJ']

#end of Globals

#lets create a class to hold all of our functions related to the scheduler
class scheduler():

    #creates a Excel file with dates listed on the left and times listed on the top.

    def createSchedule():

        #returnes a T/F if the file exists or not
        quickCheck = Path(fullFilePath).exists()

        if quickCheck == False:

            Path(fullFilePath).touch()
            scheduler.createSchedule()
            schedule = dict()
            thisWeek = list()
            # get mondays date.
            tempDay = day - weekDay

            # loop thru the week and add the dates to a dict.

            for i in range(7):

                tempDate = datetime.date(year, month, tempDay).strftime("%m/%d/%Y")
                thisWeek.append(tempDate)
                schedule.update({tempDate:dict()})
                # loop thru the time slots and add an entry to the dict.
                for j in timeSlots:
                    schedule[tempDate].update({str(j): np.NaN})

                tempDay = tempDay + 1

            pprint.pp(schedule)
            schedulePd = pd.DataFrame.from_dict(schedule)
            pprint.pp(schedulePd)
            schedulePd.to_excel(fullFilePath)


    def filterDiscordMessage(message):
        #returns a list that contains the date in the format of 'mm/dd/yyyy' and the time

        fromDiscordMessage = message
        #speperate the time and state
        fromDiscordMessage = fromDiscordMessage.split(' ')
        fromDiscodTime = fromDiscordMessage[1]
        fromDiscordDate = fromDiscordMessage[0]
        #get day and Month from 'date'
        fromDiscordDate = fromDiscordDate.split('/')
        fromDiscordDateDay = int(fromDiscordDate[1])
        fromDiscordDateMonth = int(fromDiscordDate[0])
        fromDiscordDate = datetime.date(year, fromDiscordDateMonth, fromDiscordDateDay).strftime("%m/%d/%Y")
        fromDiscordDateAndtime = list()
        fromDiscordDateAndtime = [fromDiscordDate,fromDiscodTime]
        return fromDiscordDateAndtime


    def addAppointment(user,date,time):

        schedule = pd.read_excel(fullFilePath, index_col=0)
        schedule.update(other={date:{time: user}},overwrite=False)
        schedule.to_excel(fullFilePath)

    def removeAppointment(date,time):

        schedule = pd.read_excel(fullFilePath, index_col=0)
        #assign a None value.
        schedule.loc[time,date] = np.NaN
        schedule.to_excel(fullFilePath)

    def findAppointment(date,time):
        #returns the username

        schedule = pd.read_excel(fullFilePath, index_col=0)
        return schedule.loc[time,date]

    def allOpenAppointment():
        #returns a list of all indexs of Open Sessions

        listOfOpenAppointments = list()
        schedule = pd.read_excel(fullFilePath, index_col=0)
        result = schedule.isin([np.NaN])
        allTrues = result.any()
        columnNames = list(allTrues[allTrues == True].index)
        for col in columnNames:
            rows = list(result[col][result[col] == True].index)
            for row in rows:
                listOfOpenAppointments.append((row, col))
        return listOfOpenAppointments


client = discord.Client()

@client.event

async def on_message(message):

    possibleMessageForAdd = tuple()
    possibleMessageForCancel = tuple()
    possibleMessageForRemove = tuple()
    possibleMessageForAll = tuple()

    possibleMessageForAdd = ('!add','!ADD','!Add')
    possibleMessageForCancel = ('!Cancel','!cancel','!CANCEL')
    possibleMessageForRemove = ('!Remove','!remove','!REMOVE')
    possibleMessageForAll =('!All','!ALL','!all')

    if message.author == client.user:
        return

    if message.content.startswith(possibleMessageForAdd):

        scheduler.createSchedule()
        user = message.author
        messageText = message.content[5:]
        messageText = scheduler.filterDiscordMessage(messageText)
        dateFromDiscord = messageText[0]
        timeFromDiscord = messageText[1]
        scheduler.addAppointment(user,dateFromDiscord,timeFromDiscord)

        await message.channel.send('added ' + str(user) + ' to Schedule\n' + 'Appointment set for:' + dateFromDiscord + ' ' + timeFromDiscord)

    if message.content.startswith(possibleMessageForCancel):

        scheduler.createSchedule()
        user = message.author
        user = str(user)
        messageText = message.content[8:]
        messageText = scheduler.filterDiscordMessage(messageText)
        dateFromDiscord = messageText[0]
        timeFromDiscord = messageText[1]

        quickCheck = scheduler.findAppointment(dateFromDiscord,timeFromDiscord)
        #check to make sure that the user is the one assigned to that session
        if quickCheck == user:

            scheduler.removeAppointment(dateFromDiscord,timeFromDiscord)

            await message.channel.send('removed ' + str(user) + ' from Schedule\n' + 'Appointment removed :' + dateFromDiscord + ' ' + timeFromDiscord)

        elif quickCheck != user:

            await message.channel.send('Youre not the User that created the appointment')

    if message.content.startswith(possibleMessageForRemove):

        user = message.author
        role = user.role
        permission = role.permissions
        guild = user.guild
        messageText = message.content[8:]
        messageText = scheduler.filterDiscordMessage(messageText)
        dateFromDiscord = messageText[0]
        timeFromDiscord = messageText[1]
        #keep a list of usernames for server admins
        allUsersInGuild = guild.members
        
        for member in allUsersInGuild:
                role = member.roles
                permission = role.permissions
                
                if permission is True:
                    serverAdmins.append(member)

        if permission is True:
            #verify user is an admin
            value = scheduler.findAppointment(dateFromDiscord,timeFromDiscord)
            scheduler.removeAppointment(dateFromDiscord,timeFromDiscord)

            await message.channel.send('Removed ' + value )

        else:

            await message.channel.send('Only Admins can use this command')

    if message.content.startswith(possibleMessageForAll):

        allAppointments = scheduler.allOpenAppointment()
        await message.channel.send('Please Stand by while I gather a list of open Sessions, I can only send a max of 5 Messages a Second.')

        for appointment in allAppointments:

            await message.channel.send(appointment)

token = input("Enter Bots Token Creds\n")
client.run(token)
