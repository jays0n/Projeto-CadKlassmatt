import os
import time
import PySimpleGUI as sg
import tempfile

#Defines
STRNULL=''
NULL=None
STRZERO='0'
ZERO=0
EOF=-1
LOST=-1
LISTNULL=[]
LSTNULL=[]
DICTNULL={}
DCTNULL={}
TUPLNULL=()
TUPNULL=()
MATNULL=[[]]
TRUE=True
FALSE=False


#Standard Paths
PATH_STANDARD=os.path.expanduser('~')+"\Desktop"
PATH_INCONE_MAINWINDOW="Module\__xbase__\__incones__\inconeKmt2.png"
PATH_INCONE_BUTTON="__xbase__\__incones__\iconeBotao.png"
PATH_INCONE_CHROME="Module\__xbase__\__incones__\chrIcon2.png"
PATH_INCONE_MOZILA="Module\__xbase__\__incones__\mzIcon3.png"
PATH_INCONE_EXPLORER="Module\__xbase__\__incones__\ieIcon.png"
PATH_INCONE_EDGE="Module\__xbase__\__incones__\edIcon.png"
PATH_INCONE_ADD=r"Module\__xbase__\__incones__\plusIcon.png"
PATH_INCONE_SETTINGS=r"Module\__xbase__\__incones__\settingsIcon2.png"
PATH_INCONE_TIP=r"Module\__xbase__\__incones__\tipIcon.png"
PATH_INCONE_REPORT=r"Module\__xbase__\__incones__\reportIcon.png"
PATH_INCONE_EXIT=r"Module\__xbase__\__incones__\exitIcon.png"
PATH_INCONE_UPDATE=r"Module\__xbase__\__incones__\updateIcon.png"
PATH_INCONE_EXPORT=r"Module\__xbase__\__incones__\exportIcon.png"
PATH_INCONE_CLEAR=r"Module\__xbase__\__incones__\clearIcon.png"
PATH_INCONE_REFRESH=r"Module\__xbase__\__incones__\refreshIcon.png"
PATH_INCONE_SENDEMAIL=r"Module\__xbase__\__incones__\sendIcon.png"
PATH_INCONE_GITHUB=r"Module\__xbase__\__incones__\gitIcon.png"
PATH_INCONE_ROW_ADD=r"Module\__xbase__\__incones__\increaseIcon.png"
PATH_INCONE_ROW_DEL=r"Module\__xbase__\__incones__\decreaseIcon.png"
PATH_INCONE_UPLOAD_FILE=r"Module\__xbase__\__incones__\uploadIcon.png"
PATH_INCONE_ADD_RECORD=r"Module\__xbase__\__incones__\recordIcon.png"


PATH_SETTINGS_INITIALMENU='Module\__xbase__\__files__\__paths__\__users__.json'
PATH_BASE='Module\__xbase__\__files__\__paths__\__db__.json'
PATH_TEMP='Module\__xbase__\__files__\__paths__\__temp__.json'
PATH_TABLE='Module\__xbase__\__files__\__paths__\__tables__.json'
PATH_STANDARD_TEMP_DIRECTORY=tempfile.gettempdir()

PATH_CHROME_EXE="Module\__xbase__\__files__\__drives__\chromedriver.exe"
PATH_MOZILA_EXE="Module\__xbase__\__files__\__drives__\geckodriver.exe"
PATH_EDGE_EXE="Module\__xbase__\__files__\__drives__\msedgedriver.exe"
PATH_EXPLORER_EXE="Module\__xbase__\__files__\__drives__\IEDriverServer.exe"
PATH_SETTINGS='Module\__xbase__\__files__\__paths__\__users__.json'

#Consts
CONST_TABLE_FOLLOW='FOLLOWSINS'
CONST_TABLE_SEND='SENDSINS'

TABLE_FOLLOW="FOLLOW"
TABLE_SEND="SEND"
TABLE_QUEUE="QUEUE"
TABLE_RECORDS="RECORDS"
TABLE_MATSERV="MATSERV"
TABLE_UNIDADE="UNIDADE"
TABLE_GRUPO="GRUPO"
TABLE_SUBGRUPO="SUBGRUPO"
TABLE_MODALIDADE="MODALIDADE"

STATUS_OPENNED=1
STATUS_CLOSED=2

#DATA TYPES ON DATABASE:
TP_DATA_INT=0
TP_DATA_DOUBLE=0.1
TP_DATA_STRING=STRNULL
TP_DATA_DATE_SHORT='01/01/1111'
TP_DATA_DATE='01/01/1111 00:00:00'
TP_DATA_HOUR='00:00:00'
TP_DATA_BOOLEAN=FALSE
TP_DATA_NULL=NULL

#Fields:
FIELDS_FOLLOW=['CODIGO','SIN','DESCRICAO','STATUS','KLASSMATT','OBS','CONTROLE']
FIELDS_SEND=['CODIGO','SIN','DESCRICAO','OBS']
FIELDS_QUEUE=['QUEUESIZE']
FIELDS_RECORDS=['NOME','DESCRICAO','MATSERV','UNIDADE','GRUPO','SUBGRUPO','URGENTE','USO','CAMINHO','CODIGO']
FIELDS_RECORDS_TABLE_FORM=['CODIGO','NOME','DESCRICAO','MATSERV','UNIDADE','GRUPO','SUBGRUPO','URGENTE','CAMINHO','USO']
FIELD_TYPE_FOLLOW=[TP_DATA_STRING,TP_DATA_INT,TP_DATA_STRING,TP_DATA_STRING,TP_DATA_INT,TP_DATA_STRING,TP_DATA_INT]
FIELD_TYPE_SEND=[TP_DATA_STRING,TP_DATA_INT,TP_DATA_STRING,TP_DATA_STRING]
FIELD_TYPE_QUEUE=[TP_DATA_INT]
FIELD_TYPE_RECORDS=[TP_DATA_STRING,TP_DATA_STRING,TP_DATA_STRING,TP_DATA_STRING,TP_DATA_STRING,TP_DATA_STRING,TP_DATA_BOOLEAN,TP_DATA_STRING,TP_DATA_STRING,TP_DATA_STRING]
            
#File TYPES:
FILE_XLSX='xlsx'
FILE_XLS='xls'
FILE_XLSM='xlsm'
FILE_CSV='csv'
FILE_TXT='txt'
FILE_JSON='json'

#Returns
RET_NO_ERRO=0   
RET_ERRO_USER_BLANK=2001
RET_ERRO_PASSWORD_BLANK=2002

#Erros:
ERR_NO_ERRO=0
ERR_FILE_TYPE_INVALID=71
ERR_XLS_NOT_DENIED=72
ERR_FILE_EMPTY=73
ERR_WORKBOOK_NOT_CREATED=74

ERR_LEN_TABLES=20
ERR_FILE_NOT_EXIST=21
ERR_INCORRECT_DATA=22
ERR_TABLE_ALREADY_EXIST=23
ERR_INCORRECT_FIELDS=24
ERR_FILE_UNKNOWN_ERROR=25
ERR_DATABASE_NOT_DELETED=26
ERR_TABLE_NOT_DELETED=27
ERR_TABLE_NOT_EXIST=28
ERR_RECORD_NOT_EXIST=29
ERR_CRITERION_INVALID=30
ERR_EMPTY_TABLE=31
ERR_TABLE_NOT_CREATED=32

ERR_EMAIL_NOT_FOUND=80
ERR_EMAIL_NOT_SEND=81
ERR_WITHOUT_ANY_ATTACHMENT=82
ERR_NO_PERMISSION_ACCESS_OUTLOOK=83

ERR_ELEMENT_NOT_EXIST=100
ERR_INVALID_SCRIPT=101
ERR_WEB_PAGE_ERRO=102
ERR_USER_PASSAWORD_ERRO=103
ERR_TAB_URL_INVALID=104
ERR_PATH_NOT_FOUND=105
ERR_UNKNOWN_ERROR=106
ERR_NOT_UPDATED=107

#Directives:
BY_ID='getElementById'
BY_CLASS='getElementsByClassName'
BY_NAME='getElementsByName'
BY_TAG_NAME='getElementsByTagName'
BY_TAG_NAME_NS='getElementsByTagNameNS'

#Element Action:
ELEMENT_CLICK='click()'
ELEMENT_SET_VALUE="value='{valor}'"
ELEMENT_SET_TEXT="text='{valor}'"
ELEMENT_SET_INNER_TEXT="innertext='{valor}'"
ELEMENT_GET_VALUE='value'
ELEMENT_GET_TEXT='text'
ELEMENT_GET_INNER_TEXT='innerText'

#Scripts:
SCRIPT_GET_VALUE_BY_ID="return document.getElementById('{id}').value"
SCRIPT_GET_TEXT_BY_ID="return document.getElementById('{id}').text"
SCRIPT_GET_INNERTEXT_BY_ID="return document.getElementById('{id}').innerText"

#Standard Text:
EMAIL_SIGNATURE = (r"""<P><font color='blue'><b>Atenciosamente,</b></font>
                              <br><font color='green'>Cadastro Klassmatt</font>
                              <br><font color='blue'><b>Gerência Geral de Frota</b>
                              <br>Av. General Justo, 375 - Edifício Bay View - 6º andar
                              <br>Centro - Rio de Janeiro - RJ - 20021-130</font>
                              <br> <a href='www.loginlogistica.com.br'>www.loginlogistica.com.br</a>
                              </p>""")
EMAIL_MESSAGE_SINS_SEND=(r"""<p><font color='blue'>Prezados,<br> Item(s) cadastrado(s), aguardando a geração do(s) código(s).<br>   <br> </font>""")
EMAIL_MESSAGE_SINS_FOLLOW=(r"""<p><font color='blue'>Prezados,<br> Seguem os códigos dos itens cadastrados:<br>   <br> </font>""")

#Browser Type:
BROWSER_CHROME=1
BROWSER_MOZILA=2
BROWSER_EDGE=3
BROWSER_EXPLORER=4
BROWSER_SHOW_HIDED=True
BROWSER_SHOW_RISED=False

class progressBar:
    def __init__(self,extension=100):
        self.create(extension)


    def create(self,extension=100):
        sg.theme('Dark2')
        self.layout = [[sg.Text('Processando....'),sg.Text('0%',key='progCaption',size=(5,1))],      
              [sg.ProgressBar(extension, orientation='h', size=(50, 20), key='progressbar')],      
              ]

        self.window = sg.Window('Progresso').Layout(self.layout)      
        self.progress_bar = self.window.FindElement('progressbar')    
        self.progress_caption=self.window.FindElement('progCaption')

        event, values = self.window.Read(timeout=0)
        time.sleep(1)
        #return window

    def update(self,percent):     
        time.sleep(0.01)
        self.progress_caption.Update(str(percent)+'%')
        self.progress_bar.UpdateBar(percent)
        time.sleep(0.01)

    def close(self):
        time.sleep(1)
        self.window.close()

    
    

