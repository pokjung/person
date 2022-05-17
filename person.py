import pandas as pd
import numpy as np
import pyodbc as pyConnector
import csv
import pyodbc
import log_process
'''
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
PersonalProfile
Position
PositionLevel
Offices
Provinces
'''

@log_process.logTime
def main():

    dataFromExcel = getDataFromXlsx(filePath)
    
    tableName = 'PersonalProfile'
    orderBy = 'CitizenID'
    personData = getDataFromServer(driverName,hostName,databaseName,tableName,userName,passWord,orderBy)
    
    tableName = 'Position'
    orderBy = 'PositionThName'
    positionData = getDataFromServer(driverName,hostName,databaseName,tableName,userName,passWord,orderBy)

    tableName = 'PositionLevel'
    orderBy = 'PositionLevelThName'
    positionLevelData = getDataFromServer(driverName,hostName,databaseName,tableName,userName,passWord,orderBy)

    tableName = 'Offices'
    orderBy = 'OfficeName'
    officeData = getDataFromServer(driverName,hostName,databaseName,tableName,userName,passWord,orderBy)

    for indexPerson in range(len(personData)):
        if personData[indexPerson][17] == 2 or personData[indexPerson][17] == 4:
            personData[indexPerson][16] = False

    for indexPerson in range(len(personData)):
        for indexExcel in range(len(dataFromExcel)):
            #print(type(personData[indexPerson][1]),type(dataFromExcel[indexExcel][1]))
            if str(personData[indexPerson][1]) == str(dataFromExcel[indexExcel][1]):
                #print(personData[indexPerson][1],dataFromExcel[indexExcel][1])
                #print(personData[indexPerson][3],dataFromExcel[indexExcel][3])
                for indexPosition in range(len(positionData)):
                    if str(dataFromExcel[indexExcel][5]) == str(positionData[indexPosition][2]):
                        personData[indexPerson][13] = positionData[indexPosition][0]
                        #print(str(dataFromExcel[indexExcel][5]),str(positionData[indexPosition][2]))
                        break
                
                for indexPositionLevel in range(len(positionLevelData)):
                    if str(dataFromExcel[indexExcel][6]) == str(positionLevelData[indexPositionLevel][2]):
                        personData[indexPerson][14] = positionLevelData[indexPositionLevel][0]
                        #print(str(dataFromExcel[indexExcel][6]),str(positionLevelData[indexPositionLevel][2]))
                        break
                
                for indexOffice in range(len(officeData)):
                    if str(dataFromExcel[indexExcel][7]) == str(officeData[indexOffice][2]):
                        personData[indexPerson][8] = officeData[indexOffice][0]
                        #print(str(dataFromExcel[indexExcel][7]),str(officeData[indexOffice][2]))
                        break
                
                personData[indexPerson][16] = True
                

    exportDataToCSV(personData)
    
    #for indexData in personData:
        #print(indexData[0],indexData[1],indexData[3],indexData[4],indexData[8],indexData[13],indexData[14],indexData[16])
    

    #print(len(dataFromExcel))
    #print(dataFromExcel)


@log_process.logTime
def getDataFromServer(driverName,hostName,databaseName,tableName,userName,passWord,orderBy):
    try:
        serverDetail = pyConnector.connect('DRIVER={'+str(driverName)+'};SERVER='+str(hostName)+';DATABASE='+str(databaseName)+';UID='+str(userName)+';PWD='+str(passWord))
        connectServer = serverDetail.cursor()
        sqlScript = "SELECT * FROM "+str(tableName)+" ORDER BY "+orderBy+" asc"
        sqlData = serverDetail.execute(sqlScript).fetchall()
        connectServer.commit()
        connectServer.close()
        return sqlData
    except Exception as e:
        print('Get Data Failure >>',e)

@log_process.logTime
def getDataFromXlsx(filePath):
    try:
        dataFrame = pd.read_excel(filePath)
        getDataExcel = dataFrame.to_numpy(dtype='O',na_value=np.nan)
        return getDataExcel
    except Exception as e:
        print('Get Data Failure >>',e)

@log_process.logTime
def exportDataToCSV(exportData):
    try:
        with open('xls\\export.csv',mode='w',encoding='UTF8',newline='') as fileNameCSV:
            exportToCSV = csv.writer(fileNameCSV)
            exportToCSV.writerows(exportData)
    except Exception as e:
        print('Export Error >>',e)

@log_process.logTime
def updateDataToServer():
    pass

if __name__ == '__main__':
    print('<< Start >>')
    main()