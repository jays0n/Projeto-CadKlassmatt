import sys
from PyQt5 import QtWidgets as gui
from PyQt5 import QtGui, QtCore
import webNetPy as wb
import PyFile as pf
from PyDefines import *
import MolPy as mp


#Message Icons
MSG_ICONE_CRITICAL=gui.QMessageBox.Critical
MSG_ICONE_QUESTION=gui.QMessageBox.Question
MSG_ICONE_WARNING=gui.QMessageBox.Warning
MSG_ICONE_INFORMATION=gui.QMessageBox.Information




class __pqClsLayoutMenu__:
    def __init__(self,title=STRNULL):

        self.wndApplication=gui.QApplication(LISTNULL)
        self.wndWidget=gui.QMainWindow()    
        self.multiPages=gui.QMdiArea()
        self.__status__=STATUS_OPENNED
        self.setTitle(title)
        self.__stdFont__=QtGui.QFont('Arial',12)
        self.__gerarLayout__()
        self.__stdMessageBoxStyle__=(
                                       'QMessageBox{background-color:rgb(20,20,20)}' 
                                       'QLabel{color:white;background-color:rgb(20,20,20)}' 
                                       'QPushButton{color:white;background-color:blue;height:20px;width:30px;border:1px solid white;border-radius:8px}'
                                       'QPushButton:hover{border:2px solid red}'
                                       'QPushButton:pressed{background-color:red}'

                                     )

    def __setPreferenceSettings__(self):
        pass    
            
    def setTitle(self,title=None):
        if title!=None:
            self.wndTitle=title
            self.wndWidget.setWindowTitle(title)
    
    def messageBox(self,prompt=STRNULL,title=STRNULL,incone=NULL,buttons=gui.QMessageBox.Ok,msgStyleSheet=NULL,parent=NULL):
        if parent==NULL:
            msg=gui.QMessageBox()
        else:
            msg=gui.QMessageBox(parent)

        msg.setText(prompt)
        msg.setWindowTitle(title)
        if msgStyleSheet!=NULL:
             msg.setStyleSheet(msgStyleSheet)

        if incone!=NULL:
            msg.setIcon(incone)
        if buttons!=LISTNULL:
            msg.setStandardButtons(buttons)
        
        return msg.exec()

    def inputBox(self,labelText='Informe:'):
        input=gui.QInputDialog()
        input.setLabelText(labelText)
        input.exec()      
        return input.textValue()

    def quickMessageBox(self,prompt=STRNULL,title=STRNULL):
        msg=gui.QMessageBox()
        msg.about(None,title,prompt)

    def __createPersonalWndBar__(self,title=NULL,stdSize=20, barColor=STRNULL,barWidth=390,barHeight=30):
        self.wndWidget.setWindowFlags(QtCore.Qt.FramelessWindowHint)
        
        if barColor==STRNULL:
            barColor='none'

        BARWIDTH=barWidth
        BARHEIGHT=barHeight
        STDSIZE=stdSize
        STDLOCATBUT_X=BARWIDTH-3*STDSIZE-12
        STDLOCATBUT_Y=int((BARHEIGHT-STDSIZE)/2)


        self.ptbLytBar=gui.QGroupBox(self.wndWidget)
        self.ptbLytBar.setGeometry(QtCore.QRect(5,5,BARWIDTH,BARHEIGHT))
        self.ptbLytBar.setStyleSheet('background-color:'+ barColor + ';border:2px solid rgb(20,20,20);border-radius: 6px')

        self.ptbTitleBar=gui.QLabel(self.ptbLytBar)
        self.ptbTitleBar.setText(title)
        self.ptbTitleBar.setGeometry(QtCore.QRect(int(0.5*(BARWIDTH-200)),1,200,30))
        self.ptbTitleBar.setStyleSheet('background-color:none;color:white;font:bold 12px;border:none;border-radius:3px')

        self.ptbBtClose=gui.QPushButton(self.ptbLytBar)
        self.ptbBtClose.setText(STRNULL)
        self.ptbBtClose.setStyleSheet(
                                      'QPushButton{background-color:rgb(100,0,0);color:white;font:bold 12px;border:1px solid white;border-radius:10px}'
                                      'QPushButton::pressed{background-color:grey}'  
                                      'QPushButton::hover{border:2px solid lightblue}'
                                      )
        self.ptbBtClose.setGeometry(QtCore.QRect(STDLOCATBUT_X+2*STDSIZE+2,STDLOCATBUT_Y,STDSIZE,STDSIZE))
        self.ptbBtClose.clicked.connect(self.button_close)

        self.ptbBtMaximize=gui.QPushButton(self.ptbLytBar)
        self.ptbBtMaximize.setText(STRNULL)
        self.ptbBtMaximize.setStyleSheet(
                                         'QPushButton{background-color:rgb(0,0,100);color:white;font:bold 22px;border:1px solid white;border-radius:10px}'
                                         'QPushButton::pressed{background-color:grey}'  
                                         'QPushButton::hover{border:2px solid lightblue}'
                                         )
        self.ptbBtMaximize.setGeometry(QtCore.QRect(STDLOCATBUT_X+STDSIZE+2,STDLOCATBUT_Y,STDSIZE,STDSIZE))
        self.ptbBtMaximize.clicked.connect(self.button_maximize)

        self.ptbBtMinimize=gui.QPushButton(self.ptbLytBar)
        self.ptbBtMinimize.setText(STRNULL)
        self.ptbBtMinimize.setStyleSheet(
                                         'QPushButton{background-color:rgb(0,100,0);color:white;font:bold 12px;border:1px solid white;border-radius:10px}'
                                         'QPushButton::pressed{background-color:grey}'  
                                         'QPushButton::hover{border:2px solid lightblue}'
                                         )
        self.ptbBtMinimize.setGeometry(QtCore.QRect(STDLOCATBUT_X,STDLOCATBUT_Y,STDSIZE,STDSIZE))
        self.ptbBtMinimize.clicked.connect(self.button_minimize)

    def show(self,maximized=TRUE):
        self.__setPreferenceSettings__()
        if maximized==TRUE:
            self.wndWidget.showMaximized()
        self.wndWidget.show()
        self.wndApplication.exec()

    def button_close(self):
        self.wndWidget.close()
    
    def button_minimize(self):
        self.wndWidget.showMinimized()

    def button_maximize(self):
        self.wndWidget.showMaximized()

    def close(self):
        self.wndApplication.closeAllWindows()

class __manageData__:
    def __init__(self):
        return 

    def __createTable__(self,path,tableName,fields,fieldsType):
        clsDb=pf.pj.manageJsonBd(path)
        response=clsDb.isTableEmpty(tableName)
        if response:
            clsDb.deleteTable(tableName)
        else:
            return
        return clsDb.createTable(tableName,fields,fieldsType)

    def __isDataBaseEmpty__(self,path):
        clsManageDb=pf.pj.manageJsonBd(path)
        return clsManageDb.isJsonEmpty()

    def __isTableEmpty__(self,path,tableName):
        clsManageDb=pf.pj.manageJsonBd(path)
        return clsManageDb.isTableEmpty(tableName)

    def __updateTable__(self,path,tableName,data,version=1):
        clsManageDb=pf.pj.manageJsonBd(path)
        if version!=1:
            return clsManageDb.updateTable2(tableName,data)
        return clsManageDb.updateTable(tableName,data)

    def __deleteTable__(self,path,tableName):
        clsDb=pf.pj.manageJsonBd(path)
        return clsDb.deleteTable(tableName)

    def __getTable__(self,path,tableName):
         clsDb=pf.pj.manageJsonBd(path)
         return clsDb.getTable(tableName)

    def __getRecord__(self,path,tableName,tableCriterionField,tabelCriterionValue):
        clsDb=pf.pj.manageJsonBd(path)
        return clsDb.getRecord(tableName,tableCriterionField,tabelCriterionValue)

    def __getLastRecord__(self,path,tableName):
        clsDb=pf.pj.manageJsonBd(path)
        return clsDb.getLastRecord(tableName)

    def __emptyTable__(self,path,tableName,fields,fieldsType):
        clsDb=pf.pj.manageJsonBd(path)
        response=self.__deleteTable__(path,tableName)
        if not response[0]:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=response[1]),'Erro encontrado: \n {msg}'.format(msg=response[2]))
            return
        self.__createTable__(path,tableName,fields,fieldsType)
        return

    def createAllDataBases(self):
        self.__createTable__(PATH_TEMP,TABLE_RECORDS,FIELDS_RECORDS,FIELD_TYPE_RECORDS)
        self.__createTable__(PATH_TEMP,TABLE_QUEUE,FIELDS_QUEUE,FIELD_TYPE_QUEUE)
        self.__createTable__(PATH_BASE,TABLE_SEND,FIELDS_SEND,FIELD_TYPE_SEND)
        self.__createTable__(PATH_BASE,TABLE_FOLLOW,FIELDS_FOLLOW,FIELD_TYPE_FOLLOW)
        return 
 
    def updateTempRecord(self,data):       
        response=self.__updateTable__(PATH_TEMP,TABLE_RECORDS,data)
        if not response[0]:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=response[1]),'Erro encontrado: \n {msg}'.format(msg=response[2]))
            return
        response=LSTNULL
        queueSize=[{"ID":0,"QUEUESIZE":len(data)}]
        response=self.__updateTable__(PATH_TEMP,TABLE_QUEUE,queueSize,2)
        if not response[0]:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=response[1]),'Erro encontrado: \n {msg}'.format(msg=response[2]))
            return        
        return

    def getTempSize(self):
        response=self.__getLastRecord__(PATH_TEMP,TABLE_QUEUE)
        if not response[0]:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=response[1]),'Erro encontrado: \n {msg}'.format(msg=response[2]))
            return NULL
        return response[3]["QUEUESIZE"]

    def emptySender(self):
        self.__emptyTable__(PATH_BASE,TABLE_SEND,FIELDS_SEND,FIELD_TYPE_SEND)
        return

    def emptyFollower(self):
        self.__emptyTable__(PATH_BASE,TABLE_FOLLOW,FIELDS_FOLLOW,FIELD_TYPE_FOLLOW)
        return

    def emptyTempDB(self):
        self.__emptyTable__(PATH_TEMP,TABLE_RECORDS,FIELDS_RECORDS,FIELD_TYPE_RECORDS)
        self.__emptyTable__(PATH_TEMP,TABLE_QUEUE,FIELDS_QUEUE,FIELD_TYPE_QUEUE)
        return
        
    def isDbEmpty(self):
        clsManageDb=pf.pj.manageJsonBd(PATH_BASE)
        if clsManageDb.isJsonEmpty()==FALSE:
            return FALSE
        response=clsManageDb.createTable(TABLE_FOLLOW,FIELDS_FOLLOW,FIELDS_TYPE_FOLLOW)
        if response[0]==FALSE:
            return NULL

        response=clsManageDb.createTable(TABLE_SEND,FIELDS_SEND,FIELDS_TYPE_SEND)
        if response[0]==FALSE:
            return NULL
        return TRUE    

    def orderTableBase(self,tableName,lstData):
        
        if lstData in (LSTNULL,NULL):
            return NULL

        if tableName==TABLE_SEND:
            filterField1='CODIGO'
            filterField2='SIN'
        else:
            filterField1='CONTROLE'
            filterField2='CODIGO'

        sortedData=sorted(lstData,key=lambda row:(row[filterField1],row[filterField2]))
        if tableName!=TABLE_SEND:
            sortedData=list(filter(lambda row:row['CONTROLE'] in [0,1],sortedData))          

        return sortedData

    def insertBaseRecord(self,tblName,lstData):
        response=self.isDbEmpty()
        if response==NULL:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=ERR_TABLE_NOT_CREATED),'Não foi possivel criar as tabelas no banco de dados')
            return ZERO

        clsManageDb=pf.pj.manageJsonBd(PATH_BASE)
        response=clsManageDb.insertNewsRecords(tblName,lstData)
        if response[0]==FALSE:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=response[1]),'Erro encontrado: \n {msg}'.format(msg=response[2]))
        return response[3]

    def updateBaseRecord(self,tblName,lstData):
        response=self.isDbEmpty()
        if response==NULL:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=ERR_TABLE_NOT_CREATED),'Não foi possivel criar as tabelas no banco de dados')
            return ZERO
        clsManageDb=pf.pj.manageJsonBd(PATH_BASE)
        response=clsManageDb.updateTable2(tblName,lstData)
        if response[0]==FALSE:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=response[1]),'Erro encontrado: \n {msg}'.format(msg=response[2]))
        return

    def tableBaseDifference(self,tblName,lstData):
        clsFile=pf.pj.manageJsonBd(PATH_BASE)
        response=clsFile.getTable(tblName)
        if response[0]==FALSE:
            return response
        table=response[3]

        if len(table)==len(lstData):
            return LSTNULL
        
        difTable=[]
        for eachElement in table:
            if not eachElement in lstData:
                difTable.append(eachElement)
        
        return difTable  

    def supplyDataFollowingSins(self,data):
        lstDataToReturn=[]
        for eachElement in data:
            dctElement=eachElement
            dctElement['STATUS']=STRNULL
            dctElement['KLASSMATT']=ZERO
            dctElement['OBS']=STRNULL
            dctElement['CONTROLE']=ZERO
            lstDataToReturn.append(dctElement)
        return lstDataToReturn

    def getTableData(self,tableName):
        clsManageDb=pf.pj.manageJsonBd(PATH_BASE)
        response=clsManageDb.getTable(tableName)
        if response[0]==FALSE:
            return LSTNULL
        return response[3]

    def getTableTemp(self,tableName=TABLE_RECORDS):
        response=self.__getTable__(PATH_TEMP,tableName)
        if response[0]==FALSE:
            return LSTNULL
        return response[3]

    def getDataExternFile(self,path=STRNULL):
        clsFile=pf.manageDataSettings(path)
        return clsFile.getDataToRecord()

    def getSettings(self):
        clsFile=pf.manageDataSettings(PATH_SETTINGS_INITIALMENU)
        url=clsFile.getUrl()     
        login=clsFile.__getLogin__()
        return url,login[0],login[1]

    def getDataJson(self,path,tableName):
        clsDb=pf.pj.manageJsonBd(path)
        response=clsDb.getRecords(tableName)
        if not response[0]:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=response[1]),'Erro encontrado: \n {msg}'.format(msg=response[2]))
            return
        return response[3]

class __toolbox__:
    def __init__(self):
        return

    def getStdButton(self,parent=NULL,text=STRNULL,rect=LSTNULL):
        stdButton=gui.QPushButton(parent)
        stdButton.setStyleSheet(
                                 'QPushButton{background-color:rgb(30,70,130);color:lightgray;font:12px;border:1px solid gray;border-radius:8px}'
                                 'QPushButton:Pressed{background-color:rgb(255,30,30)}'
                                 'QPushButton:hover{border:1px solid red}'
                                )
        if rect==STRNULL:
            rect=[0,0,30,20]
        stdButton.setGeometry(QtCore.QRect(rect[0],rect[1],rect[2],rect[3]))
        stdButton.setText(text)
        return stdButton

    def getStdLabel(self,parent=NULL,text=STRNULL,rect=[0,0,10,10],font=12,styleSheet=STRNULL):
        stdLabel=gui.QLabel(parent)
        if styleSheet==STRNULL:
            stdLabel.setStyleSheet('background-color:none;border:none;color:lightgray;font:bold ' +str(font)+'px')
        else:
            stdLabel.setStyleSheet(styleSheet)
        stdLabel.setGeometry(QtCore.QRect(rect[0],rect[1],rect[2],rect[3]))
        stdLabel.setText(text)
        return stdLabel

    def getStdCheckBox(self,parent=NULL,text=STRNULL,rect=[0,0,10,10],checked=FALSE,toolTip=STRNULL):
        stdCheckBox=gui.QCheckBox(parent)
        stdCheckBox.setGeometry(QtCore.QRect(rect[0],rect[1],rect[2],rect[3]))
        stdCheckBox.setText(text)
        stdCheckBox.setChecked(checked)
        stdCheckBox.setStyleSheet(
                                   "QCheckBox{color:white;font:11px}" 
                                   "QCheckBox::indicator{width:15px;height:15px;border:1px solid white;border-radius:6px}"
                                   "QCheckBox::indicator:checked{background-color:red;border:1px solid gray}"
                                  )
        stdCheckBox.setToolTip(toolTip)
        return stdCheckBox

    def getStdFrame(self,parent=NULL,rect=[0,0,10,10]):
        stdFrame=gui.QFrame(parent)
        stdFrame.setStyleSheet('background-color:rgb(30,30,30);border-radius:18px;color:white')
        stdFrame.setGeometry(QtCore.QRect(rect[0],rect[1],rect[2],rect[3]))
        return stdFrame

    def getStdTable(self,parent=NULL,nRows=1000,nCols=100,columnHeaders=NULL,rect=[0,0,10,10]):
        table=gui.QTableWidget(parent)
        table.setRowCount(nRows)
        table.setColumnCount(nCols)
        table.setStyleSheet('QTableView{background-color:rgb(60,60,60);selection-background-color:rgb(40,120,70)}')
        if columnHeaders!=NULL:
            table.setHorizontalHeaderLabels(columnHeaders)
        table.setGeometry(QtCore.QRect(rect[0],rect[1],rect[2],rect[3]))
        return table

    def getStdLineEdit(self,parent=NULL,rect=[0,0,10,10],enable=TRUE,tooltip=STRNULL):
        stdLineEdit=gui.QLineEdit(parent)
        stdLineEdit.setStyleSheet(
                                        'QLineEdit{color:lightgray;font:bold 14px;background-color:rgb(20,20,20);border:1px solid gray;border-radius:10px}'
                                       )
        stdLineEdit.setGeometry(QtCore.QRect(rect[0],rect[1],rect[2],rect[3]))
        stdLineEdit.setEnabled(enable)
        stdLineEdit.setToolTip(tooltip)
        return stdLineEdit

    def getStdComboBox(self,parent=NULL,rect=[0,0,10,10]):
        stdComboBox=gui.QComboBox(parent)
        stdComboBox.setStyleSheet(
                                   'QComboBox{color:lightgray;font:14px;background-color:rgb(20,20,20);border:1px solid gray;border-radius:6px}'
                                   'QComboBox::down-arrow{background-color:none}'
                                   'QComboBox::drop-down{background-color:rgb(210,200,200);border:1px solid black;width:8px;height:26px}'
                                   'QComboBox::drop-down:hover{border:1px solid red}'
                                   'QComboBox::drop-down:pressed{background-color:red}'
                                   'QComboBox QAbstractItemView{background-color:rgb(30,30,30);color:lightgray}'   
                                   )
        stdComboBox.setGeometry(QtCore.QRect(rect[0],rect[1],rect[2],rect[3]))
        return stdComboBox

    def getStdMenuButton(self,parent=NULL,iconPath=STRNULL,rect=[0,0,10,10],toolTip=STRNULL):
        stdButton=gui.QPushButton(parent)
        stdButton.setStyleSheet(
                                 'QPushButton{background-color:None;color:white;font:12px;border:none;border-radius:20px}'
                                 'QPushButton:Pressed{background-color:rgb(255,30,30)}'

                                )
        if rect==STRNULL:
            rect=[0,0,30,20]
        stdButton.setGeometry(QtCore.QRect(rect[0],rect[1],rect[2],rect[3]))
        stdButton.setIcon(QtGui.QIcon(iconPath))
        stdButton.setIconSize(QtCore.QSize(rect[2],rect[3]))
        stdButton.setToolTip(toolTip)
        return stdButton  

class __manageGui__:
    def __init__(self):
        pass
    def __getTableElement__(self,lst):
        for eachChild in lst:            
            className=eachChild.__class__.__name__.upper()
            if className=="QTABLEWIDGET":
                return eachChild

    def emptyTable(self,lstElements):
        table=self.__getTableElement__(lstElements)
        table.clearContents() 

    def clearTable(self,lstElements):
        msgbox=gui.QMessageBox()
        answer=msgbox.question(NULL,"Atenção!","Deseja proceder com a exclusão?")
        if answer!=gui.QMessageBox.Yes:
            return        
        self.emptyTable(lstElements)

    def updateTableContent(self,data,lstElements):    
        tableElement=self.__getTableElement__(lstElements)
        row=0
        for eachElement in data:
            column=0
            for eachValue in eachElement.values():
                cell=gui.QTableWidgetItem(NULL)
                cell.setText(str(eachValue))
                tableElement.setItem(row,column,cell)
                column+=1
            row+=1
        return 

    def getFirstRow(self,table):
        row=0
        while table.item(row,0) not in (NULL,STRNULL):
            row+=1
        return row

    def updateTableContent2(self,data,lstElements,firstRow=TRUE):    
        tableElement=self.__getTableElement__(lstElements)
        if firstRow==TRUE:
            row=0
        else:
            row=self.getFirstRow(tableElement)
        for eachElement in data:
            column=0
            for eachValue in eachElement:
                cell=gui.QTableWidgetItem(NULL)
                cell.setText(str(eachValue))
                tableElement.setItem(row,column,cell)
                column+=1
            row+=1
        return

    def getTableData(self,lstElements,fields):
        columnsCount=len(fields)
        table=self.__getTableElement__(lstElements)
        matrix=[]
        row=0
        while table.item(row,0) != NULL:
            rows={}
            for columns in range(columnsCount):
                if table.item(row,columns)==NULL:
                    cell=STRNULL
                else:
                    cell=table.item(row,columns).text()
                rows[fields[columns]]=cell
            matrix.append(rows)
            row+=1
        return matrix

    def dowloadTable(self,lstElements,fields,name):
        try:
            askFilePath=gui.QFileDialog()
            path=askFilePath.getExistingDirectory(NULL,"Selecione o diretorio")
            if path in (STRNULL,NULL):
                return FALSE
            table=self.getTableData(lstElements,fields)
            clsDb=pf.pj.manageXls(path=path,name=name)
            clsDb.recordXls(table)
            return TRUE
        except:
            return FALSE

class __recorderWindow__(gui.QDialog):
    def __init__(self, parent=NULL,title=STRNULL,text=STRNULL,size=(1365,800)):        
        self.textContent=text
        self.pSize=size
        flags=QtCore.Qt.WindowCloseButtonHint
        super().__init__(parent,flags=flags)
        self.toolbox=__toolbox__()
        self.clsManageGui=__manageGui__()
        self.setWindowTitle(title)
        self.createLayout()        
        self.__data__=[]

    def createLayout(self):  
        self.background=self.toolbox.getStdFrame(self,[-10,-10,self.pSize[0]+10,self.pSize[1]+10])
        self.setFixedSize(self.pSize[0],self.pSize[1])
        self.tblRecord=self.toolbox.getStdTable(
                                            self.background,2000,12,
                                            ('Codigo','Nome port','PN','Nome Ingles','Equip Ingles','Modelo','Fabricante','Embarcacao','N. Serie','Equip Port','Material','Obs')
                                            ,rect=[15,15,1360,400]
                                           )
        self.tblRecord.setColumnWidth(11,230)
        self.tblRecord.setColumnWidth(0,50)
        self.frBackgroundEmail=self.toolbox.getStdFrame(self.background,[950,420,420,300])
        self.frBackgroundEmail.setStyleSheet("background-color:rgb(37,39,38)")
        self.lblEmailCode=self.toolbox.getStdLabel(self.frBackgroundEmail,'Codigo do Email',[5,5,100,30])
        self.txtEmailCode=self.toolbox.getStdLineEdit(self.frBackgroundEmail,[5,40,230,30])
        self.btGetFromEmail=self.toolbox.getStdButton(self.frBackgroundEmail,'Obter do E-mail',[250,40,150,30])
        self.lblRecordToDescription=self.toolbox.getStdLabel(self.frBackgroundEmail,'Gerar Descricoes',[5,80,150,30])
        self.btRecordToDescription=self.toolbox.getStdButton(self.frBackgroundEmail,'Gerar',[5,110,150,30])
        self.lblNextStep=self.toolbox.getStdLabel(self.frBackgroundEmail,'Proxima etapa',[5,160,150,30])
        self.btNextStep=self.toolbox.getStdButton(self.frBackgroundEmail,'Continuar',[5,190,150,30])
        self.frBackgroundDescription=self.toolbox.getStdFrame(self.background,[15,420,930,300])
        self.frBackgroundDescription.setStyleSheet("background-color:rgb(37,39,38)")
        self.tblDescription=self.toolbox.getStdTable(self.frBackgroundDescription,2000,4,('Tamanho','Descricao Curta','Descricao Longa','Codigo'),[1,1,920,270])
        self.tblDescription.setColumnWidth(1,320)
        self.tblDescription.setColumnWidth(2,320)
        self.myEvents()

    def myEvents(self):
        self.btNextStep.clicked.connect(self.goNextStep)
        self.btRecordToDescription.clicked.connect(self.generateDescription)
        self.btGetFromEmail.clicked.connect(self.getDataFromEmail)
        
    def goNextStep(self):
        recContinue=__recorderWindowContinue__(NULL,'Gerar o Cadastro',[1365,800],self.__data__)
        recContinue.exec()
        self.close()

    def getDataToRecord(self):
        tabela=self.clsManageGui.__getTableElement__(self.background.children())
        cell=tabela.item(0,0)
        if cell in (NULL,STRNULL):
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro na Tabela','As celulas da coluna codigo estao vazias.')

        row=0
        data=[]
        headers=[tabela.horizontalHeaderItem(item) for item in range(12)]
        headersLabel=[item.text() for item in headers]

        while tabela.item(row,0) not in (NULL,STRNULL):                                                                                                                                                                                                     
            rows=[]
            shortDescription=''
            longDescription=''
            for column in range(1,12):
                if not tabela.item(row,column) in (STRNULL,NULL):
                    cellValue=tabela.item(row,column).text()                
                    if column in (1,2,6,7,9):
                        h=str(headersLabel[column])[0:5].strip()+': '
                        if h.upper().strip().find("NOME")>EOF:
                            h=STRNULL
                        shortDescription=shortDescription+h+cellValue +'-'
                    longDescription=longDescription+headersLabel[column]+': '+cellValue +'-'
            
            rows.append(len(shortDescription))
            rows.append(shortDescription)
            rows.append(longDescription)
            rows.append(tabela.item(row,0).text())
            data.append(rows)
            row+=1
        self.__data__=data
        return data

    def generateDescription(self):
        data=self.getDataToRecord()
        self.clsManageGui.updateTableContent2(data,self.frBackgroundDescription.children())

    def getDataFromEmail(self):
        codSol=self.txtEmailCode.text()
        msgbox=gui.QMessageBox()
        if codSol in (STRNULL,NULL):            
            msgbox.critical(NULL,'Erro de preenchimento','Favor informar o código do e-mail.')
            return
        clsEmail=mp.pyEmail(codSol)
        response=clsEmail.saveAttachement()
        if response[0]==FALSE:
            msgbox.critical(NULL,'Erro {num}'.format(num=response[1]),'Erro encontrado:\n{msg}'.format(msg=response[2]))
            return
        path=response[3]
        clsManageData=pf.dataFromRecordsXls(path)
        data=clsManageData.__getXlsData__()
        column=0
        table=[]
        for eachElement in data:
            rows=[]
            rows.append(codSol)
            for column in range(11):
                if column!=2:                   
                    cellValue=eachElement[column]
                    if cellValue in (NULL,STRNULL):
                        rows.append(STRNULL)
                    else:
                        rows.append(cellValue)
            rows.append(str(eachElement[2])+'\n'+str(eachElement[11]))
            table.append(rows)
        
        self.clsManageGui.updateTableContent2(table,self.background.children(),FALSE)
        msgbox.information(NULL,'Informe','Concluido')

class __recorderWindowContinue__(gui.QDialog):

    def __init__(self, parent=NULL,title=STRNULL,size=(410,350),data=LSTNULL):
        self.pSize=size
        self.dataTable=data
        flags=QtCore.Qt.WindowCloseButtonHint        
        super().__init__(parent=parent, flags=flags)
        self.toolbox=__toolbox__()
        self.clsManageGui=__manageGui__()
        self.clsManageData=__manageData__()
        self.setWindowTitle(title)
        self.createLayout()
        self.fillTable()        

    def createLayout(self):
        self.background=self.toolbox.getStdFrame(self,[-10,-10,self.pSize[0]+2000,self.pSize[1]+2000])
        self.setFixedSize(self.pSize[0],self.pSize[1])   
        
        self.tblContent=self.toolbox.getStdTable(
            self.background,2500,10,
            ('Codigo','Nome','Descricao','Material/Servico','Unidade Medida','Grupo','Subgrupo','Urgente','Caminho Midia','Modalidade'),
            [15,15,1050,700]
            )
        self.frMenuRecord=self.toolbox.getStdFrame(self.background,[1070,15,290,700])
        self.frMenuRecord.setStyleSheet("background-color:rgb(37,39,38)")
        self.lblRow=self.toolbox.getStdLabel(self.frMenuRecord,'Linha',[15,15,50,30])
        self.txtRow=self.toolbox.getStdLineEdit(self.frMenuRecord,[15,50,50,20])
        self.txtRow.setEnabled(FALSE)
        self.txtRow.setAlignment(QtCore.Qt.AlignCenter)
        self.btPlus=self.toolbox.getStdMenuButton(self.frMenuRecord,PATH_INCONE_ROW_ADD,[70,40,20,20],'Adicionar linha')
        self.btLess=self.toolbox.getStdMenuButton(self.frMenuRecord,PATH_INCONE_ROW_DEL,[70,60,20,20],'Retirar linha')
        self.lblUse=self.toolbox.getStdLabel(self.frMenuRecord,'Modalidade',[130,15,100,30])
        self.cbUse=self.toolbox.getStdComboBox(self.frMenuRecord,[130,50,150,25])
        self.lblMaterial=self.toolbox.getStdLabel(self.frMenuRecord,'Material/Servico',[15,90,100,30])
        self.cbMaterial=self.toolbox.getStdComboBox(self.frMenuRecord,[15,125,100,25])
        self.lblUnit=self.toolbox.getStdLabel(self.frMenuRecord,'Unidade Medida',[140,90,100,30])
        self.cbUnit=self.toolbox.getStdComboBox(self.frMenuRecord,[140,125,143,25])
        self.lblGroup=self.toolbox.getStdLabel(self.frMenuRecord,'Grupo',[15,160,100,30])
        self.cbGroup=self.toolbox.getStdComboBox(self.frMenuRecord,[15,195,250,25])
        self.lblSubgroup=self.toolbox.getStdLabel(self.frMenuRecord,'Subgrupo',[15,230,100,30])
        self.cbSubgroup=self.toolbox.getStdComboBox(self.frMenuRecord,[15,265,250,25])
        self.ckUrgent=self.toolbox.getStdCheckBox(self.frMenuRecord,'Urgente',[200,325,80,25],FALSE,'Solicitacao urgente')
        self.lblPath=self.toolbox.getStdLabel(self.frMenuRecord,'Caminho',[15,300,80,30])
        self.txtPath=self.toolbox.getStdLineEdit(self.frMenuRecord,[15,335,120,25],FALSE)
        self.btPath=self.toolbox.getStdMenuButton(self.frMenuRecord,PATH_INCONE_UPLOAD_FILE,[140,335,25,25])
        self.btClean=self.toolbox.getStdMenuButton(self.frMenuRecord,PATH_INCONE_CLEAR,[30,400,25,25],'Limpar')
        self.btRecord=self.toolbox.getStdMenuButton(self.frMenuRecord,PATH_INCONE_ADD_RECORD,[120,400,25,25],'Gravar')
        self.ckRepeat=self.toolbox.getStdCheckBox(self.frMenuRecord,'Repetir',[180,400,100,25],FALSE,'Repetir informações para todos')
        self.btFinish=self.toolbox.getStdButton(self.frMenuRecord,'Finalizar',[40,500,200,25])           

        self.myEvents()

    def myEvents(self):
        self.btPlus.clicked.connect(self.increaseRange)
        self.btLess.clicked.connect(self.decreaseRange)     
        self.fillComboBox()
        self.cbGroup.currentIndexChanged.connect(self.fillSubGroupBox)
        self.btPath.clicked.connect(self.getPathMidia)
        self.btRecord.clicked.connect(self.registerData)
        self.btClean.clicked.connect(self.deleteRow)
        self.btFinish.clicked.connect(self.finishClick)

    def fillTable(self):
        data=[]

        for eachElement in self.dataTable:
            rows=[]
            rows.append(eachElement[3])
            description=eachElement[1] + '\n'+eachElement[2]
            rows.append(description[0:50:1])
            rows.append(description)
            data.append(rows)

        self.clsManageGui.updateTableContent2(data,self.background.children(),FALSE)
        self.txtRow.setText('1')

    def getTableFromJson(self,tableName):
        clsManageData=pf.pj.manageJsonBd(PATH_TABLE)
        response=clsManageData.getTable(tableName)
        if response[0]==FALSE:
            return NULL
        return response[3]

    def getElementsFromTable(self,tableName,filter=STRNULL):
        tblData=self.getTableFromJson(tableName)        
        data=[]
        for eachItem in tblData:
            if filter==STRNULL:
                data.append(eachItem['ITEM'])
            else:
                if eachItem['CODIGO']==filter:
                    data.append(eachItem['ITEM'])
        return data

    def fillComboBox(self):   
        self.cbUse.addItems(self.getElementsFromTable(TABLE_MODALIDADE))
        self.cbMaterial.addItems(self.getElementsFromTable(TABLE_MATSERV))
        self.cbUnit.addItems(self.getElementsFromTable(TABLE_UNIDADE))
        self.cbGroup.addItems(self.getElementsFromTable(TABLE_GRUPO))
        self.fillSubGroupBox()

    def increaseRange(self):
        try:
            if self.txtRow.text() in (STRNULL,NULL):
                rowValue=1
            else:
                rowValue=int(self.txtRow.text())
            rowValue+=1
            self.txtRow.setText(str(rowValue))
            return
        except Exception as msgErr:
            gui.QMessageBox.critical(NULL,'Erro inesperado','Erro encontrado \n{msg}'.format(msg=msgErr))

    def decreaseRange(self):

        try:
            if self.txtRow.text() in (STRNULL,NULL,'1'):
                rowValue=1
            else:
                rowValue=int(self.txtRow.text())
                rowValue-=1
            self.txtRow.setText(str(rowValue))
            return
        except Exception as msgErr:
            gui.QMessageBox.critical(NULL,'Erro inesperado','Erro encontrado \n{msg}'.format(msg=msgErr))    

    def fillSubGroupBox(self):
        code=str(self.cbGroup.currentText())[0:2:1]
        data=self.getElementsFromTable(TABLE_SUBGRUPO,code)
        self.cbSubgroup.clear()
        self.cbSubgroup.addItems(data)

    def getPathMidia(self):
        try:
            askFilePath=gui.QFileDialog()
            path=askFilePath.getOpenFileName()
            self.txtPath.setText(path[0])
        except Exception as msgErr:
            gui.QMessageBox.critical(NULL,'Erro file',"Erro encontrado\n{msg}".format(msg=msgErr))

    def deleteRow(self):
        row=int(self.txtRow.text())
        table=self.clsManageGui.__getTableElement__(self.background.children())
        table.removeRow(row-1)
        if row==1:
            row=1
        else:
            row-=1
        self.txtRow.setText(str(row))

    def finishClick(self):        
        answer=gui.QMessageBox.question(NULL,'Atenção','Deseja realmente finalizar?')
        if answer!=gui.QMessageBox.Yes:
            return
        table=self.clsManageGui.getTableData(self.background.children(),FIELDS_RECORDS_TABLE_FORM)
        data=[]
        for eachElements in table:
            rows={}
            for eachField in FIELDS_RECORDS:
                rows[eachField]=eachElements[eachField]
            data.append(rows)
        self.clsManageData.updateTempRecord(data)
        self.close()        
       
    def generateRecord(self,row):
        
        table=self.clsManageGui.__getTableElement__(self.background.children())
        if self.ckUrgent.isChecked():
            urgent="SIM"
        else:
            urgent="NAO"

        values=[self.cbMaterial.currentText(),self.cbUnit.currentText(),self.cbGroup.currentText(),self.cbSubgroup.currentText(),urgent,self.txtPath.text(),self.cbUse.currentText()]
  
        column=3
        for eachItem in values:
            cell=gui.QTableWidgetItem()
            cell.setText(eachItem)
            table.setItem(row-1,column,cell)
            column+=1
        row+=1
        self.txtRow.setText(str(row))

    def registerData(self):
        row=int(self.txtRow.text())
        table=self.clsManageGui.__getTableElement__(self.background.children())
        if self.ckRepeat.isChecked():
            while table.item(row-1,0) not in (NULL,STRNULL):
                self.generateRecord(row)
                row+=1
        else:
            self.generateRecord(row)


class __tipWindow__(gui.QDialog):
    def __init__(self, parent=NULL,text=STRNULL,title=STRNULL,size=(410,350)):
        self.text=text
        self.pSize=size
        flags=QtCore.Qt.WindowCloseButtonHint
        
        super().__init__(parent=parent, flags=flags)
        self.toolbox=__toolbox__()
        self.setWindowTitle(title)
        self.createLayout()

    def createLayout(self):
        self.background=self.toolbox.getStdFrame(self,[-10,-10,self.pSize[0]+2000,self.pSize[1]+2000])
        self.setFixedSize(self.pSize[0],self.pSize[1])        
        self.lblTextContent=self.toolbox.getStdLabel(self.background,self.text,[15,-80,490,490])
        self.btGitHub=self.toolbox.getStdMenuButton(self.background,PATH_INCONE_GITHUB,[200,320,30,30],'Acesso ao Git Project')    
        self.myEvents()

    def myEvents(self):
        self.btGitHub.clicked.connect(self.goGitHub)
        return
    
    def goGitHub(self):
        browsing=wb.goGitHub()
        browsing.go()

class pqClsMenuLoginKmt(__pqClsLayoutMenu__):

    def __init__(self, title=STRNULL):
        super().__init__(title=title)
    
    def __setPreferenceSettings__(self):
        clsSettings=pf.manageDataSettings(PATH_SETTINGS_INITIALMENU)
        settings=clsSettings.getInitialMenuSettings()
        self.ckRememberLogin.setChecked(settings['RECORDLOGIN'])
        if settings['INITIALIZE']==1:
            self.wndWidget.showMaximized()
        userInfo=clsSettings.getLoginSaved()
        self.__login__=[]
        self.__login__.append(str(userInfo['USER']).upper())
        self.txtUser.setText(str(userInfo['USER']).upper())
        self.__login__.append(pf.sinHash(2,userInfo['PASSWORD']))
        self.txtPassword.setText(pf.sinHash(2,userInfo['PASSWORD']))        

    def show(self):   
        self.__setPreferenceSettings__()
        self.wndWidget.show()
        self.wndApplication.exec()
        
    def __validacaoLoginEmBranco__(self,user,senha):
        if user in (STRNULL,NULL):
            return RET_ERRO_USER_BLANK
        elif senha in (STRNULL,NULL):
            return RET_ERRO_PASSWORD_BLANK
        else:
            return RET_NO_ERRO
  
    def __gerarLayout__(self):
        self.setTitle('Login on CadKMT')

        self.wndApplication.setStyle('Fusion')
        self.wndApplication.setWindowIcon(QtGui.QIcon(PATH_INCONE_MAINWINDOW))

        self.wndWidget.setStyleSheet('background-color:rgb(12,12,12);border-radius:12px')
                
        self.gruAgrupamentoLogin=gui.QGroupBox(STRNULL,self.wndWidget)
        #self.gruAgrupamentoLogin.setGeometry(QtCore.QRect(10,10,380,400))
        self.gruAgrupamentoLogin.setStyleSheet('background-color:rgb(20,20,20);border:2px solid rgb(20,20,20);border-radius: 12px')

        self.lblUser=gui.QLabel('Usuario do Klassmatt',self.gruAgrupamentoLogin)
        self.lblUser.setGeometry(QtCore.QRect(40,20,180,30))
        self.lblUser.setFont(self.__stdFont__)
        self.lblUser.setStyleSheet('color:lightgray;border:solid rgb(20,20,20)')

        self.txtUser=gui.QLineEdit(self.gruAgrupamentoLogin)
        self.txtUser.setGeometry(QtCore.QRect(40,60,280,30))
        self.txtUser.setStyleSheet('background-color:rgb(12,12,12);border: 1px solid gray;border-radius: 8px;color:gray')
        self.txtUser.setToolTip('<p style="font-family:Arial" style="font-size:10px" style="color:green">Usuario Klassmatt</p>')
        self.txtUser.textEdited.connect(self.qlineEdit_on_changed)

        self.lblUser=gui.QLabel('Senha do Klassmatt',self.gruAgrupamentoLogin)
        self.lblUser.setGeometry(QtCore.QRect(40,100,180,30))
        self.lblUser.setFont(self.__stdFont__)
        self.lblUser.setStyleSheet('color:white;solid rgb(20,20,20)')

        self.txtPassword=gui.QLineEdit(self.gruAgrupamentoLogin)
        self.txtPassword.setGeometry(QtCore.QRect(40,140,280,30))
        self.txtPassword.setStyleSheet('background-color:rgb(12,12,12);border: 1px solid gray;border-radius: 8px;color:gray')
        self.txtPassword.setToolTip('<p style="font-family:Arial" style="font-size:10px" style="color:green">Senha Klassmatt</p>')
        self.txtPassword.setEchoMode(gui.QLineEdit.Password)
        self.txtPassword.textEdited.connect(self.qlineEdit_on_changed)

        self.btLogin=gui.QPushButton(parent=self.gruAgrupamentoLogin)
        self.btLogin.setGeometry(QtCore.QRect(40,210,250,30))
        self.btLogin.setStyleSheet(
                                   'QPushButton{color:white;background-color:rgb(0,80,220);border: 1px solid gray;border-radius: 12px}'
                                   'QPushButton::pressed {background-color:lightblue}'
                                   'QPushButton::hover {border:1px solid red}'
                                   )
        self.btLogin.setText('Sing-In')
        self.btLogin.clicked.connect(self.button_on_clicked)

        self.ckRememberLogin=gui.QCheckBox(parent=self.gruAgrupamentoLogin)
        self.ckRememberLogin.setGeometry(QtCore.QRect(300,210,100,30))
        self.ckRememberLogin.setStyleSheet(
                                            "QCheckBox{color:gray}" 
                                            "QCheckBox::indicator{width:15px;height:15px;border:2px solid white;border-radius:6px}"
                                            "QCheckBox::indicator:checked{background-color:red;border: 3px solid gray}"
                                           )
        self.ckRememberLogin.setToolTip('<p style="font-family:Arial" style="font-size:10px" style="color:green">Armazenar informacoes de Login</p>')
        self.ckRememberLogin.setText('Lembrar')

        self.wndWidget.setFixedHeight(320)
        self.wndWidget.setFixedWidth(400)
        self.gruAgrupamentoLogin.setGeometry(QtCore.QRect(10,40,380,270))
        self.__createPersonalWndBar__(self.wndTitle,barColor='rgb(15,15,15)')
        self.ptbBtMaximize.setEnabled(False)

    def qlineEdit_on_changed(self):
        self.btLogin.setStyleSheet(
                                   'QPushButton{color:red;background-color:rgb(0,80,220);border: 1px solid gray;border-radius: 12px}'
                                   'QPushButton::pressed {background-color:lightblue}'
                                   'QPushButton::hover {border:1px solid red}'
                                   )
        self.btLogin.setText('Register')
        
    def button_on_clicked(self):
        if self.__validacaoLoginEmBranco__(self.txtUser.text(),self.txtPassword.text())!=RET_NO_ERRO:            
            self.messageBox('Login ou senha em Branco!','Erro de Login',MSG_ICONE_CRITICAL,msgStyleSheet=self.__stdMessageBoxStyle__)
            return     
        
        user=str(self.txtUser.text()).upper()
        password=str(self.txtPassword.text())

        if self.btLogin.text()=='Register':
            browsing=wb.kmtBrowsing()
            response=browsing.goToKmtHomePage(user,password)
            browsing.closeBrowser()
        else:
            response=(TRUE,STRNULL)

        if response[0]==True:
            self.messageBox('Login efetuado com sucesso!','Informação',MSG_ICONE_INFORMATION,msgStyleSheet=self.__stdMessageBoxStyle__)
            if self.ckRememberLogin.isChecked()==TRUE:
                saverSettings=pf.manageDataSettings(PATH_SETTINGS_INITIALMENU)
                saverSettings.saveUser(user,password)    
            self.__status__=STATUS_CLOSED        
            self.close()

        else:
            self.__status__=STATUS_OPENNED
            self.messageBox('Erro no login ou senha!','Informação',MSG_ICONE_CRITICAL,msgStyleSheet=self.__stdMessageBoxStyle__)        
       
class __pqClsHomePage__(gui.QWidget):

    def __init__(self, parent=None, flags=QtCore.Qt.WindowFlags()):
        self.clsManageData=__manageData__()
        self.clsManageData.createAllDataBases()
        self.toolbox=__toolbox__()
        self.clsManageGui=__manageGui__()
        super().__init__(parent=parent, flags=flags)
        self.initUserInterface()       
        self.updateTable_follow()
        self.updateTable_send()

    def initUserInterface(self):
        self.gbLayoutGeral=gui.QGroupBox(self)
        self.gbLayoutGeral.setGeometry(QtCore.QRect(0,0,20000,40000))
        self.gbLayoutGeral.setStyleSheet('color:gray;background-color:rgb(15,15,15);border:none')
        self.createLayoutRegister()
        self.createLayoutSendRecord()
        self.createLayoutFollow()
        self.createLayoutGetItens()
        self.createLayouBottomBar()

    def setInitialWindowRect(self,left=0,top=0):
        self.frResgister.setGeometry(QtCore.QRect(10+left,10+top,430,195))
        self.frSenderRecord.setGeometry(QtCore.QRect(10+left,210+top,430,450))
        self.frFollowSins.setGeometry(QtCore.QRect(445+left,10+top,800,650))   
        
    def createLayoutFollow(self):
        pWidth=800
        self.frFollowSins=self.toolbox.getStdFrame(self,[445,10,pWidth,650])
        self.lblFollowSins=self.toolbox.getStdLabel(self.frFollowSins,'Acompanhamento Sins',[int(0.5*(pWidth-130)),5,145,25])
        self.tblFollowSins=self.toolbox.getStdTable(self.frFollowSins,2000,7,('Codigo','Numero do sin','Descricao','Status','Codigo Klassmatt','Obs','Controle'),[10,40,750,550]) 
        self.ckFollowSins=self.toolbox.getStdCheckBox(self.frFollowSins,'Automatico',[500,600,110,30],FALSE,'Enviar automaticamente os E-mails')
        self.btFollowSinsVerify=self.toolbox.getStdMenuButton(self.frFollowSins,PATH_INCONE_REFRESH,[350,600,40,40],'Atualizar sins')
        self.btFollowSinsSend=self.toolbox.getStdButton(self.frFollowSins,'Enviar',[600,605,90,25])
        self.btFollowUpdate=self.toolbox.getStdMenuButton(self.frFollowSins,PATH_INCONE_UPDATE,[700,5,20,20],'Atualizar Base')
        self.btFollowExport=self.toolbox.getStdMenuButton(self.frFollowSins,PATH_INCONE_EXPORT,[730,5,20,20],'Exportar')
        self.btFollowClear=self.toolbox.getStdMenuButton(self.frFollowSins,PATH_INCONE_CLEAR,[760,5,20,20],'Esvaziar todos os registros')        

        self.btFollowClear.clicked.connect(self.clearTableContent_follow)
        self.btFollowUpdate.clicked.connect(self.updateTable_follow)
        self.btFollowExport.clicked.connect(self.export_follow)
        self.btFollowSinsSend.clicked.connect(self.send_email_tblfollow)
        self.btFollowSinsVerify.clicked.connect(self.updateSinsFollow)

    def createLayoutGetItens(self):
        self.frGetItens=self.toolbox.getStdFrame(self,[8,50,80,630])
        self.btMenuGetItens=self.toolbox.getStdMenuButton(self.frGetItens,PATH_INCONE_ADD,[15,25,45,45],'Cadastrar um novo item')
        self.btMenuSettings=self.toolbox.getStdMenuButton(self.frGetItens,PATH_INCONE_SETTINGS,[15,115,45,45],'Atualizar')
        self.btMenuTips=self.toolbox.getStdMenuButton(self.frGetItens,PATH_INCONE_TIP,[15,205,45,45],'Sobre')
        self.btMenuReport=self.toolbox.getStdMenuButton(self.frGetItens,PATH_INCONE_REPORT,[15,295,45,45],'Relatórios')
        self.btMenuExit=self.toolbox.getStdMenuButton(self.frGetItens,PATH_INCONE_EXIT,[15,565,45,45],'Sair')

        self.btMenuGetItens.clicked.connect(self.showWindowRecord)
        self.btMenuTips.clicked.connect(self.showTipWindow)        
        self.btMenuSettings.clicked.connect(self.updateSoftware)

    def showWindowRecord(self):
        wndRec=__recorderWindow__(title='Cadastrar itens:')
        wndRec.exec()

    def showTipWindow(self):
        informe='Informações:\n\nCADKMT 1.0.0.0\n© 2021 AUTOPYVBA. Todos os direitos Resevados\n\nTermos de Uso:\nFree license\n\nContato:\nAUTOPYVBA\n\nTEL: (21) 96965-6759\nEMAIL: francisco.sousa1993@outlook.com'
        informe=informe+'\n\n\n\n\n\n\nPara saber como você pode contribuir para CADKMT,\nfaça check-out no projeto em GitHub.'
        wndTip=__tipWindow__(NULL,text=informe,title='Sobre')
        wndTip.exec()
 
    def createLayouBottomBar(self):
        self.frGetItens=self.toolbox.getStdFrame(self,[8,700,1345,20])
        self.lblInforms=self.toolbox.getStdLabel(self.frGetItens,'Kmt Klassmatt:',[20,0,800,20])          
        
    def createLayoutSendRecord(self):
        self.frSenderRecord=self.toolbox.getStdFrame(self,[10,185,430,475])
        self.lblSenderInfo=self.toolbox.getStdLabel(self.frSenderRecord,'Sins Criados - Enviar',[160,10,130,20])
        self.tblSinsRecorded=self.toolbox.getStdTable(self.frSenderRecord,1500,4,('Codigo','Numero do sin','Descricao','Obs'),[10,40,410,350])
        self.ckSinsSendDirect=self.toolbox.getStdCheckBox(self.frSenderRecord,'Automatico',[205,410,110,30],FALSE,'Enviar automaticamente os E-mails')       
        self.btSendEmailSinsRec=self.toolbox.getStdButton(self.frSenderRecord,'Enviar',[320,410,90,25])
        self.btSinsUpdate=self.toolbox.getStdMenuButton(self.frSenderRecord,PATH_INCONE_UPDATE,[340,5,20,20],'Atualizar')
        self.btSinsExport=self.toolbox.getStdMenuButton(self.frSenderRecord,PATH_INCONE_EXPORT,[370,5,20,20],'Exportar')
        self.btSinsClear=self.toolbox.getStdMenuButton(self.frSenderRecord,PATH_INCONE_CLEAR,[400,5,20,20],'Esvaziar todos os registros')
        
        self.btSinsClear.clicked.connect(self.clearTableContent_Send)
        self.btSinsUpdate.clicked.connect(self.updateTable_send)
        self.btSinsExport.clicked.connect(self.export_send)
        self.btSendEmailSinsRec.clicked.connect(self.send_email_tblsend)

    def createLayoutRegister(self):      
        
        self.frResgister=self.toolbox.getStdFrame(self,[10,10,430,170])  
        self.lblChooseFile=self.toolbox.getStdLabel(self.frResgister,'Carregar arquivo',[10,5,170,30])
        self.txtFilePath=self.toolbox.getStdLineEdit(self.frResgister,[10,40,300,20],FALSE,'Caminho do arquivo (not enabled)')
        self.btChooseFilePath=self.toolbox.getStdButton(self.frResgister,'Importar',[320,40,90,25])
        self.btChooseFilePath.clicked.connect(self.home_button_import)

        self.lblQueue=self.toolbox.getStdLabel(self.frResgister,'Fila',[10,70,40,30])
        self.txtQueue=self.toolbox.getStdLineEdit(self.frResgister,[40,75,60,20],FALSE,'Fila de arquivos')
        self.start_queue()
        self.btQueue=self.toolbox.getStdMenuButton(self.frResgister,PATH_INCONE_CLEAR,[110,75,20,20],'Esvaziar fila') 
        self.btQueue.clicked.connect(self.empty_queue)

        self.lblChooseBrowser=self.toolbox.getStdLabel(self.frResgister,'Escolher navegador',[10,100,170,40])
        self.cbBrowserType=self.toolbox.getStdComboBox(self.frResgister,[10,140,170,30])
        self.comboBoxAddItens()
        self.ckModoAnonimo=self.toolbox.getStdCheckBox(self.frResgister,'Modo oculto',[205,140,110,30],TRUE,'Nao exibir o navegador')
        self.btStartRegister=self.toolbox.getStdButton(self.frResgister,'Cadastrar',[320,140,90,25])
        self.btStartRegister.clicked.connect(self.home_button_register)

    def home_button_register(self,event):
        self.set_text_bottom_bar('carregando...')
        if self.txtFilePath.text() in (NULL,STRNULL) and self.txtQueue.text() in (NULL,STRNULL,ZERO,STRZERO):
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro de Importacao','Arquivo nao selecionado ou inexistente')
            return

        url,user,password=self.clsManageData.getSettings()
        data=self.clsManageData.getDataJson(PATH_TEMP,TABLE_RECORDS)

        browserTypeChoosen=1+self.cbBrowserType.currentIndex()
        hiddenMode=self.ckModoAnonimo.isChecked()
        browsing=wb.kmtBrowsing(data,browserTypeChoosen,hiddenMode,url)       
        response=browsing.recordOnKlassmatt(user,password)
        if response[0]==FALSE:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=response[1]),'Um erro ocorreu: \n {msg}'.format(msg=response[2]))
        self.save_records_not_sent(data,len(browsing.__sins__))
        self.save_records_sent(browsing.__sins__)
        self.set_text_bottom_bar('kmt Cadastro:')
        return

    def home_button_import(self):

        self.set_text_bottom_bar('carregando...')

        if self.txtQueue.text() not in (STRNULL,STRZERO):
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro de Prioridade','Esvazie a fila para gerar um novo cadastro')
            return      
        
        queueSize=self.clsManageData.getTempSize()
        if queueSize>0:
            self.txtQueue.setText(str(queueSize))
            self.txtFilePath.setText(PATH_TEMP)
            return        

        askFilePath=gui.QFileDialog()
        path=askFilePath.getOpenFileName()
        self.txtFilePath.setText(path[0])
        data=self.clsManageData.getDataExternFile(path[0])

        self.clsManageData.updateTempRecord(data)
        self.txtQueue.setText(str(len(data)))

        self.set_text_bottom_bar('kmt Cadastro:')

    def comboBoxAddItens(self):
        self.cbBrowserType.addItems(['Chrome','Mozila','Edge','Explorer'])
        self.cbBrowserType.setEditable(TRUE)
        self.cbBrowserType.setItemIcon(0,QtGui.QIcon(PATH_INCONE_CHROME))
        self.cbBrowserType.setItemIcon(1,QtGui.QIcon(PATH_INCONE_MOZILA))
        self.cbBrowserType.setItemIcon(2,QtGui.QIcon(PATH_INCONE_EDGE))
        self.cbBrowserType.setItemIcon(3,QtGui.QIcon(PATH_INCONE_EXPLORER))

    def set_text_bottom_bar(self,text):
        self.lblInforms.setText(text)
        return

    def empty_queue(self):
        self.clsManageData.emptyTempDB()
        self.txtQueue.setText(STRZERO)
        self.txtFilePath.setText(STRNULL)

    def start_queue(self):
        settings=pf.manageDataSettings(PATH_TEMP)
        data=settings.getData()
        queueSize=str(data['QUEUE'][0]['QUEUESIZE'])
        self.txtQueue.setText(queueSize)
    
    def save_records_sent(self,data):
        if len(data)==ZERO or data==[]:
            return
        followingData=self.clsManageData.supplyDataFollowingSins(data)
        response=self.clsManageData.insertBaseRecord(TABLE_FOLLOW,followingData)
        if response!=TRUE:
            return NULL
        response=self.clsManageData.insertBaseRecord(TABLE_SEND,data)
        if response!=TRUE:
            return NULL
        return TRUE     
 
    def save_records_not_sent(self,dataToRecord,numDataRecorded):
        if numDataRecorded==ZERO:
            return
        position=numDataRecorded
        data=dataToRecord[position::]
        queue=len(data)
        self.clsManageData.updateTempRecord(data)
        self.txtQueue.setText(queue)

    def clearTableContent_follow(self):   
        self.clsManageGui.clearTable(self.frFollowSins.children())
        self.clsManageData.emptyFollower()         
        
    def clearTableContent_Send(self):
        self.clsManageGui.clearTable(self.frSenderRecord.children())
        self.clsManageData.emptySender()

    def updateTable_send(self):
        data=self.clsManageData.getDataJson(PATH_BASE,TABLE_SEND)
        self.clsManageGui.updateTableContent(data,self.frSenderRecord.children())

    def updateTable_follow(self):
        self.clsManageGui.emptyTable(self.frFollowSins.children())
        data=self.clsManageData.getDataJson(PATH_BASE,TABLE_FOLLOW)
        self.clsManageGui.updateTableContent(data,self.frFollowSins.children())

    def updateSoftware(self):
        self.updateTable_follow()
        self.updateTable_send()
        self.start_queue()


    def export_follow(self):
        self.set_text_bottom_bar("CARREGANDO ...")
        work=self.clsManageGui.dowloadTable(self.frFollowSins.children(),FIELDS_FOLLOW,"KLASSMATT_SINS_ACOMPANHAMENTO.xlsx")
        if work==TRUE:
            messagebox=gui.QMessageBox()
            messagebox.information(NULL,"Informe!","Item exportado com sucesso!")
        self.set_text_bottom_bar("Kmt Klassmatt")
        return

    def export_send(self):
        self.set_text_bottom_bar("CARREGANDO ...")
        work=self.clsManageGui.dowloadTable(self.frSenderRecord.children(),FIELDS_SEND,"KLASSMATT_SINS_ENVIO.xlsx")
        if work==TRUE:
            messagebox=gui.QMessageBox()
            messagebox.information(NULL,"Informe!","Item exportado com sucesso!")
        self.set_text_bottom_bar("Kmt Klassmatt")
        return

    def __sendEmail__(self,data,sendAuto=FALSE,tableName=TABLE_SEND,messageMail=STRNULL):        
        sortedData=self.clsManageData.orderTableBase(tableName,data) 
        dataNotSent=self.clsManageData.tableBaseDifference(tableName,sortedData)
        codes=[]
        for eachItem in sortedData:
            codes.append(eachItem['CODIGO'])
        if codes in (LSTNULL,['']):
            return FALSE,sortedData,LSTNULL
        codes=list(dict(zip(codes,codes)).keys())
        loadBar=progressBar()
        n=len(codes)
        i=0
        for eachSin in codes:
            dataEmail=list(filter(lambda row:row['CODIGO'] in [eachSin],sortedData))
            clsEmail=mp.pyEmail(eachSin,sendAuto,messageMail)
            clsEmail.setText(dataEmail)
            clsEmail.sendEmail()
            percent=int(100*i/n)
            loadBar.update(percent)
            i+=1
        loadBar.update(100)
        loadBar.close()
        
        return TRUE,sortedData,dataNotSent

    def __adjustAndSend__(self,tableName,sendAuto=FALSE,msgEmail=STRNULL):
        table=self.clsManageData.getTableData(tableName)
        messageBox=gui.QMessageBox()
        try:
            work,sData,notSData=self.__sendEmail__(table,sendAuto,tableName,msgEmail)
        except Exception as msgErr:            
            messageBox.critical(NULL,"Erro {num}".format(num=ERR_EMAIL_NOT_SEND),"Message Erro: \n{msg}".format(msg=msgErr))
            return LSTNULL   
        if work==FALSE:
            messageBox.information(NULL,'Atenção','Itens não enviados!')
            return LSTNULL

        self.clsManageData.updateBaseRecord(tableName,notSData)
        messageBox.information(NULL,'Informe','Itens enviados com sucesso!!!')
        return sData

    def send_email_tblsend(self):
        try:
            self.__adjustAndSend__(TABLE_SEND,self.ckSinsSendDirect.isChecked(),EMAIL_MESSAGE_SINS_SEND)
        except Exception as msgErr:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=ERR_EMPTY_TABLE),'Erro encontrado \n {msg}'.format(msg=msgErr))
        return
        
    def send_email_tblfollow(self):
       
        try:
            self.__adjustAndSend__(TABLE_FOLLOW,self.ckFollowSins.isChecked(),EMAIL_MESSAGE_SINS_FOLLOW)
        except Exception as msgErr:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=ERR_EMPTY_TABLE),'Erro encontrado \n {msg}'.format(msg=msgErr))
        return

    def updateSinsFollow(self):
        try:
            data=self.clsManageData.getDataJson(PATH_BASE,TABLE_FOLLOW)
            sinsUpdated=wb.kmtCheckSins(data)
            data=sinsUpdated.updateSins()
            response=self.clsManageData.updateBaseRecord(TABLE_FOLLOW,data)
            msgbox=gui.QMessageBox()
            msgbox.information(NULL,'Informe','Atualizado com sucesso!')
            self.updateTable_follow()
        except Exception as msgErr:
            msgbox=gui.QMessageBox()
            msgbox.critical(NULL,'Erro {num}'.format(num=ERR_EMPTY_TABLE),'Erro encontrado: \n {msg}'.format(msg=msgErr))        
        
class pqClsAplicacaoKmt(__pqClsLayoutMenu__):

    def __init__(self, title=STRNULL):
        super().__init__(title=title)

    def __gerarLayout__(self):
        
        self.wndApplication.setStyle('Fusion')
        self.wndApplication.setWindowIcon(QtGui.QIcon(PATH_INCONE_MAINWINDOW))
        self.wndWidget=__pqClsHomePage__()        
        self.__createPersonalWndBar__('KMT - Cadastro Klassmatt',barColor='rgb(15,15,15)',barWidth=1365)
        self.wndWidget.setInitialWindowRect(100,30)
        self.wndWidget.showMaximized()
        self.wndWidget.show()
        self.wndWidget.btMenuExit.clicked.connect(self.close)

    def close(self):
        retornoMsgBox=self.messageBox('Deseja realmente sair?','Atencao',MSG_ICONE_QUESTION,gui.QMessageBox.Yes|gui.QMessageBox.No)
        if retornoMsgBox==gui.QMessageBox.Yes:
           self.wndApplication.closeAllWindows()

class menu():
    def __init__(self):
        pass

    def show(self):
        menuLogin=pqClsMenuLoginKmt()
        menuLogin.show()
        if menuLogin.__status__==STATUS_CLOSED:
            menuApplication=pqClsAplicacaoKmt()
            menuApplication.__login__=menuLogin.__login__
            menuApplication.show()








