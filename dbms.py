"""
  Title
    DBMS bridge
  Version
    1.0
  Date
    7 Oct,2021
  Author
    Riviere Liao
  Descript
    This Lib is designed to communicate between python and SQLite server.
  
    - connectDatabase()
    - disconnectDatabase()
    - initTable()
    - createData()
    - readData()
    - updateData()
    - deleteData()
    - exportData()
    - importData()
"""
import sqlite3
import csv

from matplotlib.pyplot import connect

##############################
# Begin of sub functions
##############################

def connectDatabase(database):
  """
  connectDatabase Connect to SQLite3 dbms

  Args:
      database (str): db file path
  """
  connection=None
  try:
    connection=sqlite3.connect(database)
  except sqlite3.Error:
    pass
  return(connection)

def disconnectDatabase(connection):
  connection.close()

def initTable(connection):
  connection.cursor().execute(
    '''CREATE TABLE IF NOT EXISTS result(
      id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,
      datetime DATATIME DEFAULT (DATETIME('NOW','LOCALTIME')),
      cpid TEXT NOT NULL,
      cpbin INTEGER NOT NULL,
      opcu NUMBER NOT NULL,
      freq INTEGER NOT NULL,
      freqduty INTEGER NOT NULL,
      aps NUMBER NOT NULL,
      dps NUMBER NOT NULL,
      apa INTEGER NOT NULL,
      dpa INTEGER NOT NULL,
      opc INTEGER NOT NULL,
      spc INTEGER NOT NULL,
      avr INTEGER NOT NULL,
      dvr INTEGER NOT NULL,
      cpa INTEGER NOT NULL,
      cda INTEGER NOT NULL,
      apnu NUMBER NOT NULL,
      dpnu NUMBER NOT NULL
    );'''
  )
  connection.commit()

def createData(connection,data,desc=None):
  cursor=connection.cursor()
  if desc==None:
    cursor.execute(
      "INSERT INTO result (cpid,cpbin,opcu,freq,freqduty,aps,dps,apa,dpa,opc,spc,avr,dvr,cpa,cda,apnu,dpnu) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
      (
        data['cpid'],
        data['cpbin'],
        data['opcu'],
        data['freq'],
        data['freqduty'],
        data['aps'],
        data['dps'],
        data['apa'],
        data['dpa'],
        data['opc'],
        data['spc'],
        data['avr'],
        data['dvr'],
        data['cpa'],
        data['cda'],
        data['apnu'],
        data['dpnu']
      )
    )
  else:
    cursor.execute(
      "INSERT INTO result (cpid,cpbin,opcu,freq,freqduty,aps,dps,apa,dpa,opc,spc,avr,dvr,cpa,cda,apnu,dpnu,desc) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
      (
        data['cpid'],
        data['cpbin'],
        data['opcu'],
        data['freq'],
        data['freqduty'],
        data['aps'],
        data['dps'],
        data['apa'],
        data['dpa'],
        data['opc'],
        data['spc'],
        data['avr'],
        data['dvr'],
        data['cpa'],
        data['cda'],
        data['apnu'],
        data['dpnu'],
        desc
      )
    )
  connection.commit()
  return(cursor.lastrowid)

def readData(connection,key='*',condition='TRUE',order='id',group='id'):
  return(
    connection.cursor().execute(
      "SELECT "+key+" FROM result WHERE "+condition+" GROUP BY "+group+" ORDER BY "+order
    )
  )

def updateData(connection,data,condition):
  connection.cursor().execute(
    "UPDATE result SET "+data+" WHERE "+condition+";"
  )
  connection.commit()
  return(connection.total_changes)

def deleteData(connection,condition):
  connection.cursor().execute(
    "DELETE FROM result WHERE "+condition+";"
  )
  connection.commit()
  return(connection.total_changes)

def exportData(connection,file_name):
  with open(file_name,'w',newline='') as csvfile:
    csvWriter=csv.writer(csvfile,delimiter=',',quotechar='"',quoting=csv.QUOTE_MINIMAL)
    csvWriter.writerow([
      'ID',
      'DateTime',
      'Chip ID',
      'BIN',
      'Op_Current',
      'INT_Frequency',
      'INT_Frq_Duty_percentage',
      'Active_Pixel_Stdev',
      'Dark_Pixel_Stdev',
      'A_Pixel_Avg',
      'D_Pixel_Avg',
      'Open_Pixel_Count',
      'Short_Pixel_Count',
      'Active_Vox_R',
      'Dark_Vox_R',
      'C_Pixel_Avg',
      'Calibration_Dark_Address',
      'Active_Pixel_Non_Uniformity',
      'Dark_Pixel_Non_Uniformity'
    ])
    for row in readData(connection):
      csvWriter.writerow(row)
    csvfile.close()

def importData(connection,filename):
  with open(filename,newline='') as csvfile:
    csvReader=csv.DictReader(csvfile)
    for row in csvReader:
      createData(connection,{
        'cpid':row['Chip ID']+'_import',
        'cpbin':row['BIN'],
        'opcu':row['Op_Current'],
        'freq':row['INT_Frequency'],
        'freqduty':row['INT_Frq_Duty_percentage'],
        'aps':row['Active_Pixel_Stdev'],
        'dps':row['Dark_Pixel_Stdev'],
        'apa':row['A_Pixel_Avg'],
        'dpa':row['D_Pixel_Avg'],
        'opc':row['Open_Pixel_Count'],
        'spc':row['Short_Pixel_Count'],
        'avr':row['Active_Vox_R'],
        'dvr':row['Dark_Vox_R'],
        'cpa':row['C_Pixel_Avg'],
        'cda':row['Calibration_Dark_Address'],
        'apnu':row['Active_Pixel_Non_Uniformity'],
        'dpnu':row['Dark_Pixel_Non_Uniformity'],
        })
    csvfile.close()

##############################
# End of sub functions
##############################

# cur.executemany(sql,data)

if __name__=='__main__':
  con=connectDatabase('database.db')
  initTable(con)
  data={
    'cpid':'0001',
  }
  createData(con,data)
  createData(con,data)
  # print(updateData(con,'cpid=2468',condition='id=2'))
  # print(deleteData(con,condition='cpid=2468'))
  for row in readData(con):
    print(row)

  exportData(con,'test.csv')
  disconnectDatabase(con)
