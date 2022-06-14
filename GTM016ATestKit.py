#!/usr/bin/env python3

"""
  Name:    GTM016ATestKit.py
  Version: v1.0
  License: GPL
  Author:  Riviere
  Date:    6 May, 2022
  Description:
    This script is built to test GTM016A ic.
    HotKeys:
      ESC Escape:27 QUIT
      F1  F1:112    HELP
      F5  F5:116    RESET
      F6  F6:117    START/STOP
      F7  F7:118    READ
      F8  F8:119    ADD
      F9  F9:120    Start Test
      Del Delete:46 DELETE
                    WRITE
                    EXPORT
                    IMPORT
"""
import os.path
import struct
import PySimpleGUI as sg
from serialPort import *
from dbms import *
from time import sleep,localtime,strftime
# import pyinstaller_versionfile

########################################
# Var
########################################
# pyinstaller_versionfile.create_versionfile(
#     output_file="versionfile.txt",
#     version="1.0",
#     company_name="GoalTop",
#     file_description="This app is built to test GTM016A ic.",
#     internal_name="GTM016A Test Kit",
#     legal_copyright="©GoalTop",
#     original_filename="GTM016ATestKit.exe",
#     product_name="GTM016A Test Kit"
# )

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

serial_refresh_timer=0
datalist=[]
db_connected=False

data_buffer={
  'cpid':0,
  'cpbin':0,
  'opcu':0,
  'freq':0,
  'freqduty':0,
  'aps':0,
  'dps':0,
  'apa':0,
  'dpa':0,
  'opc':0,
  'spc':0,
  'avr':0,
  'dvr':0,
  'cpa':0,
  'cda':0,
  'apnu':0,
  'dpnu':0
}

param={
  # App information
  'APP_NAME':"GTM016ATestKit",
  'Author':"Riviere",
  'Version':1.0,
  'ReleaseDate':"4 May 2022",
  # Gui configuration
  'WINDOW_TITLE':"GTM016A Test Kit",
  'ICON_WINDOW':resource_path("window.ico"),
  'ICON_WARNING':resource_path("warning.ico"),
  'ICON_QUESTION':resource_path("question.ico"),
  'WINDOW_SIZE':(1280,960),
  'FRAME_SIZE_PATH':(300,300),
  'FRAME_SIZE_DATA':(300,600),
  'HK_QUIT':(sg.WIN_CLOSED,"QUIT","Escape:27"),
  'HK_HELP':('__HELP__',"F1:112"),
  'HK_RESET':('__RST__',"F5:116"),
  'HK_START':('__RUN__',"F6:117"),
  'HK_ID_RENAME':('__RNID__'),
  'HK_ID_ADD':('__ADCP__',"F8:119"),
  'HK_ID_READ':('__RD__',"F7:118"),
  'HK_ID_WRITE':('__WR__',),
  'HK_ID_DELETE':('__DECP__','Delete:46'),
  'HK_DB_EXPORT':('__EXPORT__',),
  'HK_DB_IMPORT':('__IMPORT__',),
  'HK_ARROW_UP':('Up:38',),
  'HK_ARROW_DOWN':('Down:40',),
  'HK_DATALIST_CLICK':('__DATA__',),
  'HK_GO':('__GO__',"F9:120"),
  'MSG_HELP':"GTM016A Test Kit\n\nBy Riviere @GoalTop",
  'MSG_DISCONNECT':"Fail to connect GTM016A Test Kit hardware.\nPlease reconnect device, then hit START again.",
  'MSG_CREATE_DB':"New database created!\nAll previous result are saved in old database.",
  'MSG_DOUBLE_ID':'CPID "%s" is already exist',
  'MSG_DELETE_CHECK':'Are you sure to delete "%s"?',
  'MSG_DELETE_WARN':'"%s" is not in database.\n Delete in "Data log" only!',
  'MSG_DELETE_SELECT':'please select one from "Data log" to delete!',
  'MSG_OPEN_CSV':'"%s" saved in "%s", open it?',
  'MSG_OPEN_FAIL':'"%s" can not update,\nit might opened by another application.',
  'MSG_IMPORT_OK':'Import completed!',
  'MSG_IMPORT_ERROR':'Fail to import "%s" or content error!',
  'MSG_OVERWRITE':"Selected Chip ID already in database.\n Replace previous result?",
  'MSG_EOL':'End of list!',
  'MSG_PREVIEW':'In preview mode\nClick OK to trun off IC power.',
  'LAYOUT_RESULT':[
    ["GTM016A_BIN",'__CPBIN__'],
    ["Operating_Current",'__OPCU__'],
    ["INT_Frequency",'__FREQ__'],
    ["INT_Frq_Duty_percentage",'__FREQ_DUTY__'],
    ["Active_Pixel_Stdev",'__APS__'],
    ["Dark_Pixel_Stdev",'__DPS__'],
    ["A_Pixel_Avg",'__APA__'],
    ["D_Pixel_Avg",'__DPA__'],
    ["Open_Pixel_Count",'__OPC__'],
    ["Short_Pixel_Count",'__SPC__'],
    ["Active_Vox_R",'__AVR__'],
    ["Dark_Vox_R",'__DVR__'],
    ["C_Pixel_Avg",'__CPA__'],
    ["Calibration_Dark_Address",'__CDA__'],
    ["Active_Pixel_Non_Uniformity",'__APNU__'],
    ["Dark_Pixel_Non_Uniformity",'__DPNU__'],
  ],
  # Api connection
  'serial_selected':None,
  'SERIAL_REFRESH_TIME':50, # times APP_SPEED
  'database_file':".log.db",
  'DB_KEY_LIST':'cpbin,opcu,freq,freqduty,aps,dps,apa,dpa,opc,spc,avr,dvr,cpa,cda,apnu,dpnu',
  'CMD_OPEN_CSV':'cmd /c "start excel testlog.csv"',
  'CSV_NAME':"testlog.csv",
  'CSV_NAME_IMPORT':'testlog_import.csv',
  'VID':22417,
  'PID':22, # GTM016A Test Kit PID
  'UID':0,
  'ENDIAN_SYSTEM':"little",
  # App status
  'AUTO_CONNECTION':True,
  'APP_SPEED':100,
  'app_status':'standby',
  'STANDBY_LOCK_FUNCTIONS':[
    '__RNID__',
    '__RST__',
    '__CPID__',
    '__ADCP__',
    '__DECP__',
    '__RD__',
    '__EXPORT__',
    '__IMPORT__',
    '__GO__'
  ],
  'getAllResults':False,
  # CMD codes below should correspond to DAQ firmware
  'CMD_RST':[17,82],
  'CMD_HALT':[17,72],
  'CMD_VID_PID_UID':[18,85],
  'CMD_FREQ':[19,0],
  'CMD_FREQDUTY':[19,1],
  'CMD_APS':[19,2],
  'CMD_DPS':[19,3],
  'CMD_APA':[19,4],
  'CMD_DPA':[19,5],
  'CMD_OPC':[19,6],
  'CMD_SPC':[19,7],
  'CMD_AVR':[19,8],
  'CMD_DVR':[19,9],
  'CMD_CPA':[19,10],
  'CMD_CDA':[19,11],
  'CMD_APNU':[19,12],
  'CMD_DPNU':[19,13],
  'CMD_OPCU':[19,15],
  'CMD_CPBIN':[19,14],
  'CMD_POWERON':[17,100],
  'CMD_POWEROFF':[17,101],
  'CMD_DATAREADY':[17,200],
}

########################################
# Sub func
########################################
# ---------- UI function ------------------------
def layout_path():
  visible=False if param['AUTO_CONNECTION'] else True
  return(
    [
      [
        sg.Text(
          text='',
          justification='center',
          key='__TIME__',
          expand_x=True,
        )
      ],
      [
        sg.Text(
          text='Select Port:',
          justification='center',
          size=(8,1),
          visible=visible
        ),
        sg.Combo(
          values=['COM'+str(x) for x in range(128)],
          key='__PORT__',
          expand_x=True,
          visible=visible
        )
      ],
      [
        sg.Text(
          text='Test Kit',
          justification='center',
          size=(8,1)
        ),
        sg.Text(
          text='Not connect',
          justification='center',
          key='__DAQ__',
          expand_x=True,
        )
      ],
      [
        sg.Text(
          text='DB Path',
          justification='center',
          size=(8,1)
        ),
        sg.Input(
          key='__PATH__',
          expand_x=True
        )
      ],
      [
        sg.Listbox(
          values=[],
          key='__LIST__',
          expand_x=True,
          expand_y=True
        )
      ],
      [
        sg.Button(
          button_text='RST',
          size=(4,2),
          button_color='red',
          key='__RST__'
        ),
        sg.Button(
          button_text='START',
          size=(15,2),
          key='__RUN__',
          expand_x=True
        ),
        sg.Button(
          button_text='Help',
          size=(4,2),
          key='__HELP__'
        )
      ]
    ]
  )

def layout_data():
  return(
    [
      [
        sg.Listbox(
          values=[],
          enable_events=True,
          key='__DATA__',
          expand_x=True,
          expand_y=True
        )
      ],
      [
        sg.Text(text='CPID:',size=(4,1)),
        sg.Input(default_text='',key='__CPID__',expand_x=True),
      ],
      [
        sg.Button(button_text='Rename ID',key='__RNID__',expand_x=True)
      ],
      [
        sg.Button(button_text='ADD',key='__ADCP__',expand_x=True),
        sg.Button(button_text='READ',key='__RD__',expand_x=True),
        sg.Button(button_text='DEL',key='__DECP__',expand_x=True),
        # sg.Button(button_text='UPDATE',key='__WR__',expand_x=True)
      ],
      [
        sg.Button(button_text='Export All',key='__EXPORT__',expand_x=True),
        sg.Button(button_text='Import',key='__IMPORT__',expand_x=True)
      ]
    ]
  )

def layout_result():
  layout=[]
  for li in param['LAYOUT_RESULT']:
    layout.append([
      sg.Text(text=li[0],size=(25,1)),
      sg.Text(text="0",key=li[1],expand_x=True)
    ])
  return(layout)

def layout_chipinfo():
  return(
    [
      [
        # sg.Multiline(default_text='',expand_x=True,expand_y=True),
        sg.Button(button_text="READY",font=('Helvetica','42'),button_color=('white','#f0f0f0'),key='__BIN__',expand_x=True,expand_y=True),
        sg.Button(button_text="GO",size=(10,10),key='__GO__')
      ]
    ]
  )

def layout_display():
  return(
    [
      # [
      #   sg.Frame(title="Control",layout=[[sg.Text('',key='__CPIDKEY__')]],border_width=1,expand_x=True)
      # ],
      [
        sg.Frame(title="Result",layout=layout_result(),border_width=1,expand_x=True,expand_y=True)
      ],
      [
        sg.Checkbox("AUTO Chip ID increment",key='__AINC__',expand_x=True),
        sg.Checkbox("Preview mode",key='__PVMD__',expand_x=True)
      ],
      [
        sg.Frame(title='Chip description',size=(1,200),layout=layout_chipinfo(),border_width=1,expand_x=True)
      ]
    ]
  )

def layout_ui():
  return(
    [
      [
        sg.Frame(
          title='',
          layout=[
            [
              sg.Frame(
                title='Connection',
                layout=layout_path(),
                size=param['FRAME_SIZE_PATH'],
                border_width=1
              )
            ],
            [
              sg.Frame(
                title='Data log',
                layout=layout_data(),
                size=param['FRAME_SIZE_DATA'],
                border_width=1
              )
            ],
          ],
          border_width=0
        ),
        sg.Frame(
          title='',
          layout=[
            [
              sg.Frame(
                title='Display Panel',
                layout=layout_display(),
                size=(600,1),
                border_width=1,
                expand_y=True
              ),
            ],
            [
              sg.Button(
                button_text="QUIT",
                size=(10,2),
                expand_x=True
              )
            ]
          ],
          border_width=0,
          expand_y=True
        ),
      ]
    ]
  )

# ---------- UI control function ------------------------

def LockFunctions(key_list):
  for key in key_list:
    window[key].update(disabled=True)

def UnLockFunctions(key_list):
  for key in key_list:
    window[key].update(disabled=False)

def ToggleLockFunctions(key_list):
  for key in key_list:
    if param['app_status']=='standby':
      window[key].update(disabled=True)
    elif param['app_status']=='running':
      window[key].update(disabled=False)

def serial_read_write(port,cmd,rlength):
  data=[]
  try:
    p=serial_open(port)
    p.flushInput()
    p.write(cmd)
    sleep(0.001)
    if p.in_waiting>=rlength:
      for i in range(rlength):
        data.append(int.from_bytes(p.read(),byteorder='big'))
      p.close()
    else:
      p.close()
  except:
    pass
  return(data)

def serial_handler(serial_list):
  param['serial_selected']=None
  # Automatic DAQ finder
  if serial_list:
    for port in serial_list:
      # get VID/PID to check device connection
      data=serial_read_write(port,param['CMD_VID_PID_UID'],8)
      if len(data)>=8:
        vid=int(data[1])*256+int(data[0])
        pid=int(data[3])*256+int(data[2])
        if vid==param['VID'] and pid==param['PID']:
          param['serial_selected']=port
          param['UID']=int(data[7])*16777216+int(data[6])*65536+int(data[5])*256+int(data[4])

  # update GUI
  if param['serial_selected']!=None:
    # connected
    window['__DAQ__'].Update(value='Connected (ID:%s)'%(hex(param['UID'])),text_color='#00FF00')
  else:
    # connect fail
    if param['app_status']=='running':
      # disconnect when running
      param['app_status']='standby'
      window['__RUN__'].update('START')
      ToggleLockFunctions(param['STANDBY_LOCK_FUNCTIONS'])
      sg.popup('DAQ connection is not stable, please check the cable!',title='Unstable connection',auto_close=True,icon=param['ICON_WARNING'])
    param['UID']=0
    window['__DAQ__'].Update(value='Not connect',text_color='#FF0000')

def refreshPathList(path):
  if path=='INITIAL':
    pset_list=[a for a in os.listdir('./') if not a.startswith('.')]
    pset_list.insert(0,'..')
    window['__LIST__'].update(pset_list)
    window['__PATH__'].update(os.getcwd())
  elif os.getcwd()!=path:
    try:
      os.chdir(path)
      pset_list=[a for a in os.listdir('./') if not a.startswith('.')]
      pset_list.insert(0,'..')
      window['__LIST__'].update(pset_list)
    except:
      pass

# ---------- running only events --------------

def setPASSFAIL(result):
  if result==-1:
    window['__BIN__'].update('READY')
    window['__BIN__'].update(button_color=('white','#f0f0f0'))#sg.COLOR_SYSTEM_DEFAULT))
  elif result==0:
    window['__BIN__'].update('PASS')
    window['__BIN__'].update(button_color=('black','#00ff00'))
  else:
    window['__BIN__'].update('FAIL: BIN'+str(result))
    window['__BIN__'].update(button_color=('black','#ff0000'))

def getData(cmd,rlength,decodetype):
  data=0
  received=serial_read_write(param['serial_selected'],cmd,rlength)
  if received:
    if decodetype=='int':
      data=int.from_bytes(received,byteorder=param['ENDIAN_SYSTEM'],signed=False)
    elif decodetype=='float':
      data=struct.unpack('f',bytes(received))[0]
  return(data)

def updateResult(data):
  window['__CPBIN__'].update(data['cpbin'])
  window['__OPCU__'].update(data['opcu'])
  window['__FREQ__'].update(data['freq'])
  window['__FREQ_DUTY__'].update(data['freqduty'])
  window['__APS__'].update(data['aps'])
  window['__DPS__'].update(data['dps'])
  window['__APA__'].update(data['apa'])
  window['__DPA__'].update(data['dpa'])
  window['__OPC__'].update(data['opc'])
  window['__SPC__'].update(data['spc'])
  window['__AVR__'].update(data['avr'])
  window['__DVR__'].update(data['dvr'])
  window['__CPA__'].update(data['cpa'])
  window['__CDA__'].update(data['cda'])
  window['__APNU__'].update(data['apnu'])
  window['__DPNU__'].update(data['dpnu'])
  setPASSFAIL(data['cpbin'])

def getAllResults():
  # Get datalist
  data_buffer['cpbin']=getData(param['CMD_CPBIN'],1,'int')
  data_buffer['opcu']=getData(param['CMD_OPCU'],4,'float')
  data_buffer['freq']=getData(param['CMD_FREQ'],4,'int')
  data_buffer['freqduty']=getData(param['CMD_FREQDUTY'],4,'int')
  data_buffer['aps']=getData(param['CMD_APS'],4,'float')
  data_buffer['dps']=getData(param['CMD_DPS'],4,'float')
  data_buffer['apa']=getData(param['CMD_APA'],4,'int')
  data_buffer['dpa']=getData(param['CMD_DPA'],4,'int')
  data_buffer['opc']=getData(param['CMD_OPC'],4,'int')
  data_buffer['spc']=getData(param['CMD_SPC'],4,'int')
  data_buffer['avr']=getData(param['CMD_AVR'],4,'int')
  data_buffer['dvr']=getData(param['CMD_DVR'],4,'int')
  data_buffer['cpa']=getData(param['CMD_CPA'],4,'int')
  data_buffer['cda']=getData(param['CMD_CDA'],1,'int')
  data_buffer['apnu']=getData(param['CMD_APNU'],4,'float')
  data_buffer['dpnu']=getData(param['CMD_DPNU'],4,'float')

  #  Update UI
  updateResult(data_buffer)

  # STORAGE INTO DATABASE
  createData(connection=con_db,data={
    'cpid':values['__DATA__'][0],
    'cpbin':data_buffer['cpbin'],
    'opcu':data_buffer['opcu'],
    'freq':data_buffer['freq'],
    'freqduty':data_buffer['freqduty'],
    'aps':data_buffer['aps'],
    'dps':data_buffer['dps'],
    'apa':data_buffer['apa'],
    'dpa':data_buffer['dpa'],
    'opc':data_buffer['opc'],
    'spc':data_buffer['spc'],
    'avr':data_buffer['avr'],
    'dvr':data_buffer['dvr'],
    'cpa':data_buffer['cpa'],
    'cda':data_buffer['cda'],
    'apnu':data_buffer['apnu'],
    'dpnu':data_buffer['dpnu'],
  })

  # Auto increment
  if values['__AINC__']:
    next_index=datalist.index(values['__DATA__'][0])+1
    if next_index<len(datalist):
      window['__DATA__'].update(set_to_index=[next_index])
    else:
      window['__AINC__'].update(False)
      sg.popup(param['MSG_EOL'],icon=param['ICON_WARNING'])

  if values['__PVMD__']:
    sg.popup_ok(param['MSG_PREVIEW'],icon=param['ICON_WARNING'])
  tmp=serial_read_write(param['serial_selected'],param['CMD_POWEROFF'],1)


########################################
# Init
########################################
window=sg.Window(
  title=param["WINDOW_TITLE"],
  layout=layout_ui(),
  return_keyboard_events=True,
  icon=param['ICON_WINDOW'],
  finalize=True
)
window['__DAQ__'].bind('<Button-1>','')

LockFunctions(['__BIN__'])
LockFunctions(param['STANDBY_LOCK_FUNCTIONS'])

refreshPathList('INITIAL')
########################################
# Main loop
########################################
while True:
  event,values=window.read(timeout=param['APP_SPEED'])

  # 'QUIT' event
  if event in param['HK_QUIT']:
    break

  # Timestamp update
  window['__TIME__'].update(strftime('%Y-%m-%d %H:%M:%S',localtime()))

  # Check connection
  if param['AUTO_CONNECTION']:
    if param['serial_selected']==None or serial_refresh_timer>=param['SERIAL_REFRESH_TIME']:
      serial_refresh_timer=0
      serial_handler(list_ports())
    else:
      serial_refresh_timer+=1
  else:
    # manual connect
    if values['__PORT__']:
      param['serial_selected']=values['__PORT__']
    if event=='__DAQ__':
      serial_handler([param['serial_selected'],])

  # 'database path' event
  if values['__LIST__']:
    path=os.getcwd()+'\\'+values['__LIST__'][0]
    window['__PATH__'].update(path)
    refreshPathList(path)

  # 'HELP' event
  if event in param['HK_HELP']:
    sg.popup(param['MSG_HELP'],icon=param['ICON_QUESTION'])

  # 'RST' event
  if event in param['HK_RESET']:
    serial_read_write(param['serial_selected'],param['CMD_RST'],1)

  # 'START/STOP' event
  if event in param['HK_START']:
    if param['serial_selected']==None:
      param['app_status']='standby'
      window['__RUN__'].update('START')
      sg.popup(param['MSG_DISCONNECT'],title='Error message',auto_close=True,icon=param['ICON_WARNING'])
    else:
      if param['app_status']=='standby':
        param['app_status']='running'
        UnLockFunctions(param['STANDBY_LOCK_FUNCTIONS'])
        window['__RUN__'].update('STOP')
        # Sync datalist from database
        if not param['database_file'] in os.listdir():
          db_connected=True
          con_db=sqlite3.connect(param['database_file'])
          initTable(con_db)
          sg.popup(param['MSG_CREATE_DB'],title='Notice',auto_close=True,icon=param['ICON_WARNING'])
        else:
          con_db=sqlite3.connect(param['database_file'])
        datalist=[row[0] for row in readData(con_db,key='cpid',order='cpid',group='cpid')]
        window['__DATA__'].update(values=datalist)
      else:
        param['app_status']='standby'
        LockFunctions(param['STANDBY_LOCK_FUNCTIONS'])
        window['__RUN__'].update('START')
        # Stop DAQ all action
        serial_read_write(param['serial_selected'],param['CMD_HALT'],1)

# ---------- dbms function below ------------
  # Rename chip ID
  if event in param['HK_ID_RENAME'] and values['__CPID__'] and values['__DATA__'] and param['app_status']=='running':
    if values['__CPID__'] in datalist:
      sg.popup(param['MSG_DOUBLE_ID']%values['__CPID__'])
    else:
      datalist[datalist.index(values['__DATA__'][0])]=values['__CPID__']
      window['__DATA__'].update(values=datalist)
      updateData(connection=con_db,data='cpid="%s"'%(values['__CPID__']),condition='cpid="%s"'%(values['__DATA__'][0]))

  # Get chip IDs from dbms
  if event in param['HK_ID_READ'] and param['app_status']=='running':
    datalist=[row[0] for row in readData(con_db,key='cpid',order='cpid',group='cpid')]
    window['__DATA__'].update(values=datalist)

  # if event in param['HK_ID_WRITE'] and param['app_status']=='running':
  #   print(datalist)

  # Add chip ID
  if event in param['HK_ID_ADD'] and values['__CPID__']!='' and param['app_status']=='running':
    if not values['__CPID__'] in datalist:
      datalist.append(values['__CPID__'])
      window['__DATA__'].update(values=datalist)
    else:
      sg.popup(param['MSG_DOUBLE_ID']%values['__CPID__'],title='Error',auto_close=True,icon=param['ICON_WARNING'])

  # Delete chip ID
  if event in param['HK_ID_DELETE'] and param['app_status']=='running':
    if values['__DATA__']:
      if values['__DATA__'][0] in datalist and [row for row in readData(connection=con_db,key='cpid',condition='cpid="'+values['__DATA__'][0]+'"',group='cpid')]:
        del_ans=sg.popup_yes_no(param['MSG_DELETE_CHECK']%values['__DATA__'][0],icon=param['ICON_QUESTION'])
        if del_ans=="Yes":
          deleteData(con_db,'cpid="%s"'%(values['__DATA__'][0]))
          datalist.remove(values['__DATA__'][0])
          window['__DATA__'].update(values=datalist)
      else:
        sg.popup(param['MSG_DELETE_WARN']%values['__DATA__'][0],title='Error',auto_close=True,icon=param['ICON_WARNING'])
        datalist.remove(values['__DATA__'][0])
        window['__DATA__'].update(values=datalist)
    else:
      sg.popup(param['MSG_DELETE_SELECT'],title='Error',auto_close=True,icon=param['ICON_WARNING'])

  # if event=='__DESC__' and param['app_status']=='running':
  #   if values['__VALUE_CPID__'] and param['app_status']=='running':
  #     if values['__VALUE_CPID__'] in [x[0] for x in readData(con_db,key='cpid',condition='onWidth=0',group='cpid')]:
  #       deleteData(con_db,'cpid="%s" and onWidth=0'%(values['__VALUE_CPID__']))
  #     createData(
  #       con_db,
  #       {
  #         'cpid':values['__VALUE_CPID__'],
  #       },
  #       values['__CPDE__']
  #     )

  # Export results from dbms
  if event in param['HK_DB_EXPORT']:
    try:
      exportData(con_db,param['CSV_NAME'])
      if 'OK'==sg.popup_ok_cancel(param['MSG_OPEN_CSV']%(param['CSV_NAME'],os.getcwd()),title='Notice',icon=param['ICON_WARNING']):
        os.system(param['CMD_OPEN_CSV'])
    except:
      sg.popup(param['MSG_OPEN_FAIL']%(param['CSV_NAME']),title='Notice',icon=param['ICON_WARNING'])

  # Import results into dbms
  if event in param['HK_DB_IMPORT']:
    try:
      importData(con_db,param['CSV_NAME_IMPORT'])
      sg.popup(param['MSG_IMPORT_OK'],icon=param['ICON_WARNING'])
      datalist=[row[0] for row in readData(con_db,key='cpid',order='cpid',group='cpid')]
      window['__DATA__'].update(values=datalist)
    except:
      sg.popup(param['MSG_IMPORT_ERROR']%param['CSV_NAME_IMPORT'],icon=param['ICON_WARNING'])

  # Update UI from previous result
  if event in param['HK_DATALIST_CLICK'] and param['app_status']=='running':
    if values['__DATA__']!=[]:
      r_list=[]
      for row in readData(connection=con_db,key=param['DB_KEY_LIST'],condition='cpid="%s"'%(values['__DATA__'][0])):
        r_list=row
      if r_list:
        data_buffer['cpbin']=r_list[0]
        data_buffer['opcu']=r_list[1]
        data_buffer['freq']=r_list[2]
        data_buffer['freqduty']=r_list[3]
        data_buffer['aps']=r_list[4]
        data_buffer['dps']=r_list[5]
        data_buffer['apa']=r_list[6]
        data_buffer['dpa']=r_list[7]
        data_buffer['opc']=r_list[8]
        data_buffer['spc']=r_list[9]
        data_buffer['avr']=r_list[10]
        data_buffer['dvr']=r_list[11]
        data_buffer['cpa']=r_list[12]
        data_buffer['cda']=r_list[13]
        data_buffer['apnu']=r_list[14]
        data_buffer['dpnu']=r_list[15]
      else:
        data_buffer['cpbin']=-1
        data_buffer['opcu']=0
        data_buffer['freq']=0
        data_buffer['freqduty']=0
        data_buffer['aps']=0
        data_buffer['dps']=0
        data_buffer['apa']=0
        data_buffer['dpa']=0
        data_buffer['opc']=0
        data_buffer['spc']=0
        data_buffer['avr']=0
        data_buffer['dvr']=0
        data_buffer['cpa']=0
        data_buffer['cda']=0
        data_buffer['apnu']=0
        data_buffer['dpnu']=0
      updateResult(data_buffer)

  # Up/Down movement events
  if event in param['HK_ARROW_UP'] and param['app_status']=='running':
    if values['__DATA__']:
      current_index=datalist.index(values['__DATA__'][0])
      if current_index>0:
        window['__DATA__'].update(set_to_index=[current_index-1])
        r_list=[]
        for row in readData(connection=con_db,key=param['DB_KEY_LIST'],condition='cpid="%s"'%(datalist[current_index-1])):
          r_list=row
        data_buffer['cpbin']=r_list[0]
        data_buffer['opcu']=r_list[1]
        data_buffer['freq']=r_list[2]
        data_buffer['freqduty']=r_list[3]
        data_buffer['aps']=r_list[4]
        data_buffer['dps']=r_list[5]
        data_buffer['apa']=r_list[6]
        data_buffer['dpa']=r_list[7]
        data_buffer['opc']=r_list[8]
        data_buffer['spc']=r_list[9]
        data_buffer['avr']=r_list[10]
        data_buffer['dvr']=r_list[11]
        data_buffer['cpa']=r_list[12]
        data_buffer['cda']=r_list[13]
        data_buffer['apnu']=r_list[14]
        data_buffer['dpnu']=r_list[15]
        updateResult(data_buffer)
    else:
      window['__DATA__'].update(set_to_index=[0])
      r_list=[]
      for row in readData(connection=con_db,key=param['DB_KEY_LIST'],condition='cpid="%s"'%(datalist[0])):
        r_list=row
      data_buffer['cpbin']=r_list[0]
      data_buffer['opcu']=r_list[1]
      data_buffer['freq']=r_list[2]
      data_buffer['freqduty']=r_list[3]
      data_buffer['aps']=r_list[4]
      data_buffer['dps']=r_list[5]
      data_buffer['apa']=r_list[6]
      data_buffer['dpa']=r_list[7]
      data_buffer['opc']=r_list[8]
      data_buffer['spc']=r_list[9]
      data_buffer['avr']=r_list[10]
      data_buffer['dvr']=r_list[11]
      data_buffer['cpa']=r_list[12]
      data_buffer['cda']=r_list[13]
      data_buffer['apnu']=r_list[14]
      data_buffer['dpnu']=r_list[15]
      updateResult(data_buffer)

  if event in param['HK_ARROW_DOWN'] and param['app_status']=='running':
    if values['__DATA__']:
      current_index=datalist.index(values['__DATA__'][0])
      if current_index<(len(datalist)-1):
        window['__DATA__'].update(set_to_index=[current_index+1])
        r_list=[]
        for row in readData(connection=con_db,key=param['DB_KEY_LIST'],condition='cpid="%s"'%(datalist[current_index+1])):
          r_list=row
        data_buffer['cpbin']=r_list[0]
        data_buffer['opcu']=r_list[1]
        data_buffer['freq']=r_list[2]
        data_buffer['freqduty']=r_list[3]
        data_buffer['aps']=r_list[4]
        data_buffer['dps']=r_list[5]
        data_buffer['apa']=r_list[6]
        data_buffer['dpa']=r_list[7]
        data_buffer['opc']=r_list[8]
        data_buffer['spc']=r_list[9]
        data_buffer['avr']=r_list[10]
        data_buffer['dvr']=r_list[11]
        data_buffer['cpa']=r_list[12]
        data_buffer['cda']=r_list[13]
        data_buffer['apnu']=r_list[14]
        data_buffer['dpnu']=r_list[15]
        updateResult(data_buffer)
    else:
      window['__DATA__'].update(set_to_index=[0])
      r_list=[]
      for row in readData(connection=con_db,key=param['DB_KEY_LIST'],condition='cpid="%s"'%(datalist[0])):
        r_list=row
      data_buffer['cpbin']=r_list[0]
      data_buffer['opcu']=r_list[1]
      data_buffer['freq']=r_list[2]
      data_buffer['freqduty']=r_list[3]
      data_buffer['aps']=r_list[4]
      data_buffer['dps']=r_list[5]
      data_buffer['apa']=r_list[6]
      data_buffer['dpa']=r_list[7]
      data_buffer['opc']=r_list[8]
      data_buffer['spc']=r_list[9]
      data_buffer['avr']=r_list[10]
      data_buffer['dvr']=r_list[11]
      data_buffer['cpa']=r_list[12]
      data_buffer['cda']=r_list[13]
      data_buffer['apnu']=r_list[14]
      data_buffer['dpnu']=r_list[15]
      
      updateResult(data_buffer)

  # Start a new test
  if event in param['HK_GO']:
    database=[]

    # Get datalist from dbms
    for row in readData(connection=con_db,key='cpid',order='cpid',group='cpid'):
      database.append(row[0])

    # Check cpid selection
    if param['app_status']=='running' and values['__DATA__']:
      database_push=False
      if values['__DATA__'][0] in database:
        if values['__AINC__']==False:
          answer=sg.popup_yes_no(param['MSG_OVERWRITE'],icon=param['ICON_QUESTION'])
          if answer=='Yes':
            deleteData(connection=con_db,condition='cpid="%s"'%(values['__DATA__'][0]))
            database_push=True
          else:
            database_push=False
        else:
          database_push=True
      else:
        database_push=True
      
      # Start transfer
      if database_push:
        tmp=serial_read_write(param['serial_selected'],param['CMD_POWERON'],1)
        param['getAllResults']=True

  if param['getAllResults']:
    if serial_refresh_timer<=20:
      serial_refresh_timer+=1
    else:
      tmp=serial_read_write(param['serial_selected'],param['CMD_DATAREADY'],1)
      if len(tmp)>0:
        if tmp[0]==1:
          getAllResults()
          param['getAllResults']=False

      serial_refresh_timer=0

########################################
# Memery recycle
########################################
if db_connected:
  con_db.close()
window.close()
