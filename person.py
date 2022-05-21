from fileinput import filename
import pandas as pd
import numpy as np
import pyodbc as pyConnector
import csv
import pyodbc
import log_process

filePath = 'xls\\person.xlsx'
driverName = 'ODBC Driver 17 for SQL Server'
hostName = '192.168.7.150'
userName =  'sa'
passWord = 'mocbip@ssw0rd'
databaseName = 'MOI_Personal'
'''
filePath = 'xls\\person.xlsx'
driverName = 'ODBC Driver 17 for SQL Server'
hostName = 'NOHEREEVIL\SQLEXPRESS'
userName =  'sa'
passWord = 't2907717'
databaseName = 'ops_DB'
'''
'''
PersonalProfile
Position
PositionLevel
Offices
Provinces
'''

checkError = []

@log_process.logTime
def main():

    dataFromExcel = getDataFromXlsx(filePath)
    
    tableName = 'PersonalProfile'
    orderBy = 'CitizenID'
    whereField = '[OrgID] = 9 and [GroupID] = 2'
    personData = getDataFromServer(driverName,hostName,databaseName,tableName,userName,passWord,orderBy,whereField)
    
    tableName = 'Position'
    orderBy = 'PositionThName'
    whereField = ''
    positionData = getDataFromServer(driverName,hostName,databaseName,tableName,userName,passWord,orderBy,whereField)

    tableName = 'PositionLevel'
    orderBy = 'PositionLevelThName'
    whereField = ''
    positionLevelData = getDataFromServer(driverName,hostName,databaseName,tableName,userName,passWord,orderBy,whereField)

    tableName = 'Offices'
    orderBy = 'OfficeName'
    whereField = ''
    officeData = getDataFromServer(driverName,hostName,databaseName,tableName,userName,passWord,orderBy,whereField)

    for indexPerson in range(len(personData)):
        if personData[indexPerson][17] == 2:
            personData[indexPerson][16] = 0
            if not personData[indexPerson][13]:
                personData[indexPerson][13] = 0
            
            if not personData[indexPerson][14]:
                personData[indexPerson][14] = 0

    #checkData = []

    for indexPerson in range(len(personData)):
        for indexExcel in range(len(dataFromExcel)):
            if str(personData[indexPerson][1]) == str(dataFromExcel[indexExcel][1]):
                for indexPosition in range(len(positionData)):
                    if str(dataFromExcel[indexExcel][3]) == str(positionData[indexPosition][2]):
                        personData[indexPerson][13] = positionData[indexPosition][0]
                        break
                
                for indexPositionLevel in range(len(positionLevelData)):
                    if str(dataFromExcel[indexExcel][4]) == str(positionLevelData[indexPositionLevel][2]):
                        personData[indexPerson][14] = positionLevelData[indexPositionLevel][0]
                        break
                
                for indexOffice in range(len(officeData)):
                    if str(dataFromExcel[indexExcel][5]) == str(officeData[indexOffice][2]):
                        personData[indexPerson][8] = officeData[indexOffice][0]
                        break
                
                personData[indexPerson][16] = 1
                #checkData.append(str(dataFromExcel[indexExcel][1]))
   
    
    #exportDataToCSV('export.csv',personData)
    
    for indexData in range(len(personData)):
        updateDataToServer(driverName,hostName,databaseName,userName,passWord,personData[indexData][8],personData[indexData][13],personData[indexData][14],personData[indexData][16],personData[indexData][1])
    
    exportDataToCSV('excel.csv',checkError)

@log_process.logTime
def getDataFromServer(driverName,hostName,databaseName,tableName,userName,passWord,orderBy,whereField):
    try:
        serverDetail = pyConnector.connect('DRIVER={'+str(driverName)+'};SERVER='+str(hostName)+';DATABASE='+str(databaseName)+';UID='+str(userName)+';PWD='+str(passWord))
        connectServer = serverDetail.cursor()
        if whereField:
            sqlScript = "SELECT * FROM "+str(tableName)+" WHERE "+whereField+" ORDER BY "+orderBy+" asc"
        else:
            sqlScript = "SELECT * FROM "+str(tableName)+" ORDER BY "+orderBy+" asc"
        sqlData = serverDetail.execute(sqlScript).fetchall()
        connectServer.commit()
        connectServer.close()
        return sqlData
    except Exception as e:
        print('Get Data Failure >>',e)

@log_process.logTime
def updateDataToServer(driverName,hostName,databaseName,userName,passWord,officeID,positionID,positionLevelID,statusID,CitizenID):
    try:
        serverDetail = pyConnector.connect('DRIVER={'+str(driverName)+'};SERVER='+str(hostName)+';DATABASE='+str(databaseName)+';UID='+str(userName)+';PWD='+str(passWord))
        connectServer = serverDetail.cursor()
        sqlScript = "UPDATE [PersonalProfile] SET [OfficeID] = "+str(officeID)+", [PositionID] = "+str(positionID)+", [PositionLevelID] = "+str(positionLevelID)+", [Status] = "+str(statusID)+" WHERE [CitizenID] = "+str(CitizenID)
        #print(sqlScript)
        serverDetail.execute(sqlScript)
        connectServer.commit()
        connectServer.close()
        #return sqlData
    except Exception as e:
        print('Get Data Failure >>',e)
        checkError.append(CitizenID)

@log_process.logTime
def getDataFromXlsx(filePath):
    try:
        dataFrame = pd.read_excel(filePath)
        getDataExcel = dataFrame.to_numpy(dtype='O',na_value=np.nan)
        return getDataExcel
    except Exception as e:
        print('Get Data Failure >>',e)

@log_process.logTime
def exportDataToCSV(fileName,exportData):
    try:
        with open('xls\\'+fileName,mode='w',encoding='UTF8',newline='') as fileNameCSV:
            exportToCSV = csv.writer(fileNameCSV)
            exportToCSV.writerows(exportData)
    except Exception as e:
        print('Export Error >>',e)

if __name__ == '__main__':
    print('<< Start >>')
    main()