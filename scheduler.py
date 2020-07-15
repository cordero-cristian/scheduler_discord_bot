#!/bin/python

from datetime import date
import datetime
import pathlib
from pathlib import Path
import pandas as pd

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
timeSlots = ['1100','1200','1300','1400','1500','1600','1700','1800','1900','2000']

#this is here for testing.
#people = list()
#people = ['personA','personB','personC','personD','personE','personF','personG','personH','personI','personJ']

#end of Globals 

#lets create a class to hold all of our functions related to the scheduler
class scheduler():

    #creates a Excel file with dates listed on the left and times listed on the top. 
    def newWeek():

        #check to see if a scheudle already exist
        scheduler.createSchedule()
        schedule = dict()
        thisWeek = list()
        #get mondays date.
        tempDay = day - weekDay

        #loop thru the week and add the dates to a dict.
        for i in range(7):

            tempDate = datetime.date(year,month,tempDay).strftime("%m/%d/%Y")
            thisWeek.append(tempDate)
            schedule.update({tempDate : {}})
            #loop thru the time slots and add an entry to the dict.
            for j in timeSlots:
                schedule[tempDate].update({j:{}})

            tempDay = tempDay + 1
        
        schedulePd = pd.DataFrame(schedule)
        #transpose the pandas DataFrame 
        schedulePd = schedulePd.T
        schedulePd.to_excel(fullFilePath)

    def createSchedule():

        #returnes a T/F if the file exists or not 
        quickCheck = Path(fullFilePath).exists()

        if quickCheck == False:
            Path(fullFilePath).touch()

scheduler.newWeek()
