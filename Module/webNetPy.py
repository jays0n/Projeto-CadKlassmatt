import selenium.webdriver as wd
from selenium.common.exceptions import * 
from selenium.webdriver.support.ui import WebDriverWait as esperar
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from PyFile import manageDataSettings as mds
import time,os
from PyDefines import *


class __webBrowsing__: 
    def __init__(self,browserType=BROWSER_CHROME,hiddenMode=BROWSER_SHOW_HIDED,urlWeb=STRNULL): 
        self.__url__=urlWeb
        self.hwnd=ZERO
        self.browser=self.__setBrowserChosen__(browserType,hiddenMode)         

    def setUrl(self,urlWeb):
        self.__url__=urlWeb

    def getCurrentUrl(self):
        return str(self.browser.current_url)

    def __switchWebTab__(self,lstPartialUrl=LSTNULL,newHwnd=NULL,attempts=3):
        if newHwnd!=NULL:
            self.browser.switch_to.window(newHwnd)
            return TRUE,ERR_NO_ERRO,STRNULL

        for eachAttempt in range(attempts):
            for eachHWND in self.browser.window_handles:
                self.browser.switch_to.window(eachHWND)
                currentUrl=str(self.getCurrentUrl()).upper()
                if currentUrl.find(str(lstPartialUrl[0]).upper())!=LOST and currentUrl.find(str(lstPartialUrl[1]).upper())!=LOST:
                    self.browser.switch_to.window(eachHWND)
                    return TRUE,ERR_NO_ERRO,STRNULL    
                
        self.browser.switch_to.window(self.hwnd)
        return FALSE,ERR_TAB_URL_INVALID,STRNULL

    def __setBrowserChosen__(self,BrowserType=BROWSER_CHROME,hiddenMode=BROWSER_SHOW_RISED):
        if BrowserType==BROWSER_EXPLORER:
            browserOption=wd.IeOptions()
            return wd.Ie(ie_options=browserOption,executable_path=PATH_EXPLORER_EXE)
        elif BrowserType==BROWSER_EDGE:
            return wd.Edge(executable_path=PATH_EDGE_EXE)
        elif BrowserType==BROWSER_MOZILA:
            browserOption=wd.FirefoxOptions()
            browserOption.headless=hiddenMode
            return wd.Firefox(options=browserOption,executable_path=PATH_MOZILA_EXE)
        else:
            browserOption=wd.ChromeOptions()
            browserOption.headless=hiddenMode
            return wd.Chrome(options=browserOption,executable_path=PATH_CHROME_EXE)

    def closeBrowser(self):
        time.sleep(0.5)
        self.browser.quit()

    def __getStandardScriptText__(self,elementCode,elementPropertieValue=STRNULL,elementDirective=BY_ID,elementAction=ELEMENT_CLICK):
        stdText="document.{directive}('{code}').{action}"
        eAction=elementAction
        if elementAction!=ELEMENT_CLICK:
            eAction=str(elementAction).format(valor=elementPropertieValue)
        stdText=stdText.format(directive=elementDirective,code=elementCode,action=eAction)
        return stdText

    def waitBrowserAction(self,elementPropertType,elementCode,timeWaiting=5,attempts=3):
        attempt=1
        while TRUE:
            try:
                waitingTime=esperar(self.browser,timeWaiting).until(EC.presence_of_element_located((elementPropertType,elementCode)))
                return TRUE,ERR_NO_ERRO,STRNULL,waitingTime
            except WebDriverException as msgErr:
                if attempt>attempts:
                    self.closeBrowser()
                    return FALSE,ERR_ELEMENT_NOT_EXIST,msgErr.args[0],NULL
                attempt+=1

    def runScript(self,scriptText,attempts=3):
        attempt=1
        while TRUE:
            try:
                webReturn=self.browser.execute_script(scriptText)
                return TRUE,ERR_NO_ERRO,STRNULL,webReturn
            except WebDriverException as msgErr:
                if attempt>attempts:
                    self.closeBrowser()
                    return FALSE,ERR_ELEMENT_NOT_EXIST,msgErr.args[0],NULL
                attempt+=1

    def waitSeveralActionsToClick(self,elementsProperties,elementsCodes,timeWating=5,attempts=3):
        position=0
        while position<len(elementsProperties):
            response=self.waitBrowserAction(elementsProperties[position],elementsCodes[position],timeWaiting=timeWating,attempts=attempts)
            if response[0]==FALSE:
                return response[0:3]
            time.sleep(0.5)
            response[3].click()
            position+=1
        return TRUE,ERR_NO_ERRO,STRNULL

    def runSeveralScripts(self,scriptsList,attempts=3):
        for eachScript in scriptsList:
            response=self.runScript(eachScript,attempts)
            if response[0]==False:
                return response
            time.sleep(0.1)
        return response

    def getWebElement(self,elementPropertType,elementCode,attempts=3):
        attempt=1
        while TRUE:
            try:
                webElement=self.browser.find_element(elementPropertType,elementCode)    
                return TRUE,ERR_NO_ERRO,STRNULL,webElement
            except NoSuchElementException as msgErr:
                if attempt>attempts:
                    self.closeBrowser()
                    return FALSE,ERR_ELEMENT_NOT_EXIST,msgErr.args[0],NULL
                attempt+=1

    def goToWebPage(self,url=STRNULL,attempts=3):
        if url==STRNULL:
            url=self.__url__
        self.setUrl(url)
        attempt=1
        try:
            self.browser.get(url)
            self.hwnd=self.browser.current_window_handle.__str__()     
            return TRUE,ERR_NO_ERRO,STRNULL
        except Exception as msgErr:
            if attempt>attempts:
                return FALSE,ERR_WEB_PAGE_ERRO,msgErr.args[0]
            attempt+=1
        return TRUE,ERR_NO_ERRO,STRNULL

class kmtBrowsing(__webBrowsing__):

    def __init__(self,recordset=LSTNULL,browserType=BROWSER_CHROME, hiddenMode=BROWSER_SHOW_HIDED, urlWeb=STRNULL):
        self.recordset=recordset
        url=mds(PATH_SETTINGS)
        super().__init__(browserType=browserType, hiddenMode=hiddenMode, urlWeb=url.getUrl())
        
    def goToKmtHomePage(self,user,password):
        response=self.goToWebPage(self.__url__)
        if response[0]==FALSE:
            self.closeBrowser()
            return response
        
        response=self.waitBrowserAction(By.ID,'txtUsuario')
        if response[0]==FALSE:
            self.closeBrowser()
            return response
        response[3].send_keys(user)

        scriptsToRun=[
                      self.__getStandardScriptText__('txtSenha',elementPropertieValue=password,elementAction=ELEMENT_SET_VALUE),
                      self.__getStandardScriptText__('cmdEntrar')
                     ]
        response=self.runSeveralScripts(scriptsToRun)
        if response[0]==FALSE:
            self.closeBrowser()
            return response

        try:
            elementVerify=self.browser.find_element_by_id('LabelResult')
            self.closeBrowser()
            return FALSE,ERR_USER_PASSAWORD_ERRO,STRNULL
        except Exception as msgErr:
            return response

    def __setNameDescriptionOnPageKmt__(self,dctData):
        response=self.waitSeveralActionsToClick([By.PARTIAL_LINK_TEXT,By.ID],['Solicitar um Item Novo','butNotFoundItem'])
        if response[0]==FALSE:
            return response

        response=self.waitBrowserAction(By.ID,'txtNome')
        if response[0]==FALSE:
            return response[0:3]
        scripts=[
                    self.__getStandardScriptText__('txtNome',dctData[FIELDS_RECORDS[0]],BY_ID,ELEMENT_SET_VALUE),
                    self.__getStandardScriptText__('txtDescricao',dctData[FIELDS_RECORDS[1]],BY_ID,ELEMENT_SET_VALUE),    
                    self.__getStandardScriptText__('butContinuar')
                ]

        response=self.runSeveralScripts(scripts)
        return response

    def __setUrgentTypeMeasurementOnPageKmt__(self,dctData):

        response=self.waitBrowserAction(By.ID,'txtTipo')
        if response[0]==FALSE:
            return response[0:3]
        
        response[3].send_keys(dctData[FIELDS_RECORDS[2]])
        scripts=self.__getStandardScriptText__('txtUnidadeMedida',dctData[FIELDS_RECORDS[3]],BY_ID,ELEMENT_SET_VALUE)
        response=self.runScript(scripts)
        if response[0]==FALSE:
            return response

        if dctData[FIELDS_RECORDS[6]]==TRUE:
            scripts=["__doPostBack('ctl00$Body$dlPrioridade$ctl02$LinkButton2','')"]
        else:
            scripts=["__doPostBack('ctl00$Body$dlPrioridade$ctl01$LinkButton2','')"]

        scripts.append("__doPostBack('ctl00$Body$dlTab$ctl02$lbutMenu','')")

        response=self.runSeveralScripts(scripts)
  
        return response

    def __setGroupSubgroupUse__(self,dctData):
        response=self.waitBrowserAction(By.ID,'ctl00_Body_tabCategorias_ucCategorias1_gvCategorias_ctl04_txtCat')
        if response[0]==FALSE:
            return response

        scripts=[
                  self.__getStandardScriptText__('ctl00_Body_tabCategorias_ucCategorias1_gvCategorias_ctl04_txtCat',dctData[FIELDS_RECORDS[7]],BY_ID,ELEMENT_SET_VALUE),
                  self.__getStandardScriptText__('ctl00_Body_tabCategorias_ucCategorias1_gvCategorias_ctl02_txtCat',dctData[FIELDS_RECORDS[4]],BY_ID,ELEMENT_SET_VALUE),
                  self.__getStandardScriptText__('ctl00_Body_tabCategorias_ucCategorias1_gvCategorias_ctl03_txtCat',dctData[FIELDS_RECORDS[5]],BY_ID,ELEMENT_SET_VALUE)
                ]

        response=self.runSeveralScripts(scripts)
        if response[0]==FALSE:
            return response

        response=self.waitSeveralActionsToClick([By.PARTIAL_LINK_TEXT],['Mídias'])
        return response

    def __setMidias__(self,dctData):
        response=self.__switchWebTab__(['login.klassmatt.com.br','Midia.aspx?'])
        if response[0]==FALSE:
            return response
        
        for eachPath in dctData[FIELDS_RECORDS[8]].split(';'):
            response=self.runScript("this.mostraAba('aba-caminho')")
            if response[0]==FALSE:
                return response

            response=self.getWebElement(By.ID,'file')
            if response[0]==FALSE:
                return response
            
            try:
                response[3].send_keys(eachPath)
            except Exception as msgErr:
                return FALSE,ERR_PATH_NOT_FOUND,msgErr.args[0]

            response=self.waitSeveralActionsToClick(2*[By.ID],['cmdSalvar','Linkbutton1'])
            if response==FALSE:
                return response

        response=self.runScript('this.window.close()')
        if response[0]==FALSE:
            return response
        response=self.__switchWebTab__(newHwnd=self.hwnd)
        return response

    def __setFinalInforms__(self):
        response=self.waitSeveralActionsToClick(2*[By.ID],['butAcao2','butContinuar'])
        if response[0]==FALSE:
            return response

        response=self.getWebElement(By.ID,'label_numeroSin')
        if response[0]==FALSE:
            return response
        idSin=str(response[3].text)[-5:]

        response=self.waitSeveralActionsToClick([By.ID],['butNao'])
        if response[0]==FALSE:
            return response

        response=self.runScript("this.__doPostBack('ctl00$Body$TopMenu1$dlmenu$ctl00$lbutopcao','')")
        if response[0]==FALSE:
            return response

        return True,ERR_NO_ERRO,STRNULL,idSin
 
    def __loopRecord__(self,lstData):
        self.__sins__=[]
        node=0
        while node<len(lstData):
            dctData=lstData[node]         
            response=self.__setNameDescriptionOnPageKmt__(dctData)
            if response[0]==FALSE:
                return response

            response=self.__setUrgentTypeMeasurementOnPageKmt__(dctData)
            if response[0]==FALSE:
                return response

            response=self.__setGroupSubgroupUse__(dctData)
            if response[0]==FALSE:
                return response

            response=self.__setMidias__(dctData)
            if response[0]==FALSE:
                return response

            response=self.__setFinalInforms__()
            if response[0]==FALSE:
                return response

            self.__sins__.append({'CODIGO':dctData['codigo'],'SIN':response[3],'DESCRICAO':dctData['descricao']})
            node+=1

        return TRUE,ERR_NO_ERRO,STRNULL

    def __recordItensKmt__(self,data,user,password):
        response=self.goToKmtHomePage(user,password)
        if response[0]==FALSE:
            return response       

        response=self.__loopRecord__(data)
        if response[0]==FALSE:
            return response

    def recordOnKlassmatt(self,user,password):
        try:
            response=self.__recordItensKmt__(self.recordset,user,password)
            if response[0]==FALSE:
                self.closeBrowser()
            return response[0:3]
        except Exception as msgErr:
            return False,ERR_UNKNOWN_ERROR,msgErr

class kmtCheckSins(__webBrowsing__):

    def __init__(self,sins=LSTNULL,browserType=BROWSER_CHROME, hiddenMode=BROWSER_SHOW_HIDED):
        self.sins=sins
        manage=mds(PATH_SETTINGS)
        self.__login__=manage.__getLogin__()
        super().__init__(browserType=browserType, hiddenMode=hiddenMode, urlWeb=manage.getUrl())

    def goToKmtHomePage(self):
        user=self.__login__[0]
        password=self.__login__[1]
        response=self.goToWebPage(self.__url__)
        if response[0]==FALSE:
            self.closeBrowser()
            return response
        
        response=self.waitBrowserAction(By.ID,'txtUsuario')
        if response[0]==FALSE:
            self.closeBrowser()
            return response
        response[3].send_keys(user)

        scriptsToRun=[
                      self.__getStandardScriptText__('txtSenha',elementPropertieValue=password,elementAction=ELEMENT_SET_VALUE),
                      self.__getStandardScriptText__('cmdEntrar')
                     ]
        response=self.runSeveralScripts(scriptsToRun)
        if response[0]==FALSE:
            self.closeBrowser()
            return response

        try:
            elementVerify=self.browser.find_element_by_id('LabelResult')
            self.closeBrowser()
            return FALSE,ERR_USER_PASSAWORD_ERRO,STRNULL
        except Exception as msgErr:
            return response

    def goToKmtFollowSins(self):
        self.runScript("__doPostBack('ctl00$Body$rptSolicitacaoItemNovo$ctl01$lkbutMenu','')")
        try:
            response=self.__loop__()
        except Exception as msgErr:
            return FALSE,ERR_NOT_UPDATED,msgErr

        if response[0]==FALSE:
            return response
     
    def getCodeControl(self,text):
        if text.find('APROVAD')>EOF:
            return 0
        elif text.find('CANCEL')>EOF:
            return 1
        else:
            return 2

    def webScraping(self,sinInfo):
        dataDct=sinInfo

        response=self.runScript(SCRIPT_GET_VALUE_BY_ID.format(id="txtStatus"))
        if response[0]==FALSE:
            return response[0:3],(dataDct,)
        resp=response[3]

        dataDct['STATUS']=resp
        dataDct['CONTROLE']=self.getCodeControl(resp)
        if resp.find('APROVAD')>EOF:
            response=self.runScript(SCRIPT_GET_INNERTEXT_BY_ID.format(id="txtD0"))
            if response[0]==FALSE:
                return response[0:3],(dataDct,)
            dataDct['DESCRICAO']=response[3]

            response=self.runScript(SCRIPT_GET_VALUE_BY_ID.format(id="txtCodigo"))
            if response[0]==FALSE:
                return response[0:3],(dataDct,)
            dataDct['KLASSMATT']=response[3]

        elif resp.find('CANCEL')>EOF or resp.find('REVIS')>EOF or resp.find('COMPLEM')>EOF:
            response=self.runScript("return document.getElementsByTagName('strong')[0].textContent")
            if response[0]==FALSE:
                return response[0:3],(dataDct,)
            dataDct['OBS']=response[3]
        return TRUE,ERR_NO_ERRO,STRNULL,dataDct           
           
    def __loop__(self):        
        self.__sins__=[]
        progressbar=progressBar()
        n=len(self.sins)
        i=0
        for eachSin in self.sins:
            self.waitBrowserAction(By.ID,'butFiltrar')    
            scriptText= "abreSIN({sin})".format(sin=eachSin['SIN'])
            response=self.runScript(scriptText)
            if response[0]==FALSE:
                return response
            
            self.waitBrowserAction(By.ID,'txtStatus')
            response=self.webScraping(eachSin)
            if response[0]==TRUE:
                self.__sins__.append(response[3])
            response=self.runScript("__doPostBack('ctl00$Body$TopMenu1$dlmenu$ctl01$lbutopcao','')")
            if response[0]==FALSE:
                return response
            percent=int(100*i/n)
            progressbar.update(percent)
            i=i+1
        progressbar.update(100)
        progressbar.close()
        return TRUE,ERR_NO_ERRO,STRNULL

    def updateSins(self):
         self.goToKmtHomePage()
         self.goToKmtFollowSins()
         self.browser.close()
         return self.__sins__

class goGitHub(__webBrowsing__):
    def __init__(self, browserType=BROWSER_CHROME, hiddenMode=BROWSER_SHOW_RISED, urlWeb=STRNULL):
        super().__init__(browserType=browserType, hiddenMode=hiddenMode, urlWeb=urlWeb)
    def go(self):
        self.goToWebPage(r'https://github.com/jays0n/Projeto-CadKlassmatt/tree/main/Module')
