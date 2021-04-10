import selenium.webdriver as wd
from selenium.common.exceptions import * 
from selenium.webdriver.support.ui import WebDriverWait as esperar
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
import XlsPy as xp
from TxtPy import txtCreateListSinCad
import time,os

conjuntoErros={'ErrCarregarDado':49,'ErrAbrirNav':50,'ErrObjNaoLocal':51,'ErrBaseVazia':52,'ErrExecJavaScript':53,'ErrMidia':54}
dictOpcoesNavegador={'1':'chrome','2':'firefox','3':'edge','4':'explorer'}

class kmtCadastro:
    nomeItem=''
    descricao=''
    matServ=''
    unMed=''
    grupo=''
    subGrupo=''
    urgente=True
    modalidade=''
    MidiaPaths=[]
    codSol=''
 
    def Preenche(self,nome,descricaoItem,MaterialOuServico,unidadeMedida,GrupoItem,SubGrupoItem,ItemEUrgente,modalidadeItem,midias,pcodSol=''):
        self.nomeItem=nome
        self.descricao=descricaoItem
        self.matServ=MaterialOuServico
        self.unMed=unidadeMedida
        self.grupo=GrupoItem
        self.subGrupo=SubGrupoItem        
        self.urgente=ItemEUrgente
        self.modalidade=modalidadeItem
        self.MidiaPaths=midias
        self.codSol=pcodSol

    def Show(self):
        print(self.nomeItem+';'+self.descricao)


class __wpNavegarWeb__: 
    def __init__(self,strTipoNavegador='1',modoDeNavegacaoOculto=False): 
        self.__url__=''
        self.hwnd=0
        self.navegador=None
        self.__modoOculto__=modoDeNavegacaoOculto
        self.__lstOpcaoNavegador__={'chrome':wd.ChromeOptions(),
                                    'mozila':wd.FirefoxOptions(),
                                    'edge':0,
                                    'explorer':wd.IeOptions()}

        self.tipoNavegador=dictOpcoesNavegador[strTipoNavegador]
        self.opcaoNavegador=self.__lstOpcaoNavegador__[self.tipoNavegador]
        self.opcaoNavegador.headless=modoDeNavegacaoOculto     
        
        

    def setUrl(self,urlWeb):
        self.__url__=urlWeb

    def getCurrentUrl(self):
        return str(self.navegador.current_url)

    def __returnNavegateChosen__(self,navegateCod,configOptions):
        if navegateCod=='4':
            return wd.Ie(ie_options=configOptions)
        elif navegateCod=='3':
            return wd.Edge()
        elif navegateCod=='2':
            return wd.Firefox(firefox_options=configOptions)
        else:
            return wd.Chrome(chrome_options=configOptions)

    def __changeNavegatorParams__(self,tipoDeNavegador='1',modoNavegacaoOculto=False):
        self.__modoOculto__=modoNavegacaoOculto
        self.tipoNavegador=dictOpcoesNavegador[tipoDeNavegador]
        self.opcaoNavegador=self.__lstOpcaoNavegador__[self.tipoNavegador]
        self.opcaoNavegador.headless=self.__modoOculto__
        self.navegador=self.__returnNavegateChosen__(tipoDeNavegador,self.opcaoNavegador)      
 

    def closeNavegator(self):
        self.navegador.quit()

    def __aguardarAcaoNavegador__(self,tipoComponente,codigoComponente):
        tentativa=1
        while True:
            try:
                TempoEspera=esperar(self.navegador,5).until(EC.presence_of_element_located((tipoComponente,codigoComponente)))
                return True,'Sem Erro',0
            except WebDriverException as wdErro:
                if tentativa>3:
                    self.closeNavegator()
                    return False,wdErro.args[0],conjuntoErros.pop('ErrObjNaoLocal')
                tentativa+=1
    
    def __executarJavaScript__(self,scriptText):
        tentativa=1
        while True:
            try:
                self.navegador.execute_script(scriptText)
                return True,'Sem Erro',0
            except WebDriverException as erro:
                if tentativa>3:
                    self.closeNavegator()
                    return False,erro.args[0],conjuntoErros.pop('ErrExecJavaScript')
                tentativa+=1


    def __localizarWebElemento__(self,tipoComponente,codigoComponente):
        tentativa=1
        while True:
            try:
                webElemento=self.navegador.find_element(tipoComponente,codigoComponente)              
                return True,'Sem Erro',0,webElemento
            except NoSuchElementException as wdErro:
                if tentativa>3:
                    self.navegador.quit()
                    return False,wdErro.args[0],conjuntoErros.pop('ErrObjNaoLocal'),None
                tentativa+=1



    def navegateOnWeb(self,urlWeb,modoNavegacaoOculto=False,tipoDeNavegador='1'):
        self.setUrl(urlWeb)
        self.__changeNavegatorParams__(tipoDeNavegador,modoNavegacaoOculto)
        self.hwnd=self.navegador.current_window_handle

        try:
            self.navegador.get(urlWeb)
            return True,'Sem Erros', 0

        except Exception as erroNav:
            self.navegador.quit()
            return False, erroNav.args[0], conjuntoErros.pop('ErrAbrirNav') 


    def __TrocarAbaNavegacao__(self,partialUrl=[],newHwnd=None):

        if newHwnd!=None:
            self.navegador.switch_to.window(newHwnd)
            return

        self.__originalHWND__=self.hwnd

        tentativa=1
        while tentativa<=3:
            for eachHwnd in self.navegador.window_handles:
                self.navegador.switch_to.window(eachHwnd)
                if str(str(self.navegador.current_url).upper()).find(str(partialUrl[0]).upper())>-1 and str(str(self.navegador.current_url).upper()).find(str(partialUrl[1]).upper())>-1:
                    self.hwnd=eachHwnd                    
                    tentativa=4
                    break
                else:
                    self.navegador.switch_to.window(self.__originalHWND__)
            tentativa+=1
            

class navegadorKmt(__wpNavegarWeb__):
    def __init__(self, caminhoArquivoCadastro,abaPlanilhaCadastro,strTipoNavegador='1', modoDeNavegacaoOculto=False):
        self.__pathRegisterFile__=caminhoArquivoCadastro
        self.__worksheetTab__=abaPlanilhaCadastro
        super().__init__(strTipoNavegador=strTipoNavegador, modoDeNavegacaoOculto=modoDeNavegacaoOculto)        
    
    def __obterDadosCadastro__(self):
        tentativa=1
        while tentativa<=3:
            try:
                dados2Cadastro,quantidadeRegistros,quantidadeAnexosPorRegistro=xp.xlGeraVetClass(self.__pathRegisterFile__,self.__worksheetTab__)
                return True,'Sem Erro',0,dados2Cadastro,quantidadeRegistros,quantidadeAnexosPorRegistro
            except Exception as erro:
                return False, erro.args[0],conjuntoErros.pop('ErrCarregarDado'),None,0,None
            else:
                if dados2Cadastro[0] in [100,101,102,103]:
                    return False,dados2Cadastro[1],dados2Cadastro[0],None,0,None
            tentativa+=1
        return False, 'Erro Não Localizado',1,None,0,None

    def __singInSiteKmt__(self,loginKmt,senhaKmt):
        funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'txtUsuario')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro

        funcionou,msgErro,codErro=self.__executarJavaScript__("document.getElementById('txtUsuario').value='"+str(loginKmt).strip()+"'")
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro

        funcionou,msgErro,codErro=self.__executarJavaScript__("document.getElementById('txtSenha').value='"+str(senhaKmt).strip()+"'")
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro

        funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(By.ID,'cmdEntrar')
        if funcionou==False:
            self.closeNavegator()
        else:
            webElemento.click()            
        return funcionou,msgErro,codErro

    def __registerMidiaKmt__(self,dadoSin,qtdAnexosSin):

        self.__TrocarAbaNavegacao__(['login.klassmatt.com.br','Midia.aspx?']) 

        funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'cmdSalvar')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro
        time.sleep(0.5)

        funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(By.PARTIAL_LINK_TEXT,'Arquivo')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro
        webElemento.click()
        time.sleep(0.5)

        if qtdAnexosSin>1:
            midiasDoSin=dadoSin.MidiaPaths
        else:
            midiasDoSin=[dadoSin.MidiaPaths]          

        for eachElement in range(qtdAnexosSin):
            funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'file')
            if funcionou==False:
                self.closeNavegator()
                return funcionou,msgErro,codErro

            funcionou,msgErro,codErro, webElemento=self.__localizarWebElemento__(By.ID,'file')
            if funcionou==False:
                self.closeNavegator()
                return funcionou,'Problemas com a midia indicada.',conjuntoErros.pop('ErrMidia')
            webElemento.send_keys(os.path.abspath(str(midiasDoSin[eachElement]).strip()))

            funcionou,msgErro,codErro, webElemento=self.__localizarWebElemento__(By.ID,'cmdSalvar')
            if funcionou==False:
                self.closeNavegator()
                return funcionou,msgErro,codErro
            time.sleep(1)
            webElemento.click()

            if qtdAnexosSin>1:
                funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'Linkbutton1')
                if funcionou==False:
                    self.closeNavegator()
                    return funcionou,msgErro,codErro

                funcionou,msgErro,codErro, webElemento=self.__localizarWebElemento__(By.ID,'Linkbutton1')
                if funcionou==False:
                    self.closeNavegator()
                    return funcionou,msgErro,codErro
                webElemento.click()
                time.sleep(1)

        

        funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'Label1')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro

        funcionou,msgErro,codErro, webElemento=self.__localizarWebElemento__(By.ID,'cmdFechar')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro
        webElemento.click()
        time.sleep(0.8)

        self.__TrocarAbaNavegacao__(self.__originalHWND__)

        return True,'Sem Erro',0

    def __registerSinsInKmt__(self,dadoSin,qtdAnexosSin):

        conjByTipo=[By.CLASS_NAME,By.PARTIAL_LINK_TEXT,By.ID,By.ID]
        conjTextElement=['ks-box-item-title','Solicitar um Item Novo','butNotFoundItem','butNotFoundItem']

        posicao=0
        while posicao<4:
            funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(conjByTipo[posicao],conjTextElement[posicao])
            if funcionou==False:
                self.closeNavegator()
                return funcionou,msgErro,codErro,'__registerSinsInKmt__1.1',[]

            funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(conjByTipo[posicao+1],conjTextElement[posicao+1])
            if funcionou==False:
                self.closeNavegator()
                return funcionou,msgErro,codErro,'__registerSinsInKmt__1.2',[]
            webElemento.click()
            posicao+=2
            time.sleep(0.5)

        funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'txtNome')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__1.3',[]

        funcionou,msgErro,codErro=self.__executarJavaScript__("document.getElementById('txtNome').value='"+str(dadoSin.nomeItem).strip()+"'")
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__1.4',[]
        time.sleep(0.5)


        funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(By.ID,'txtDescricao')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__1.5',[]
        webElemento.send_keys(str(dadoSin.descricao).strip())
        time.sleep(0.5)

        funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(By.ID,'butContinuar')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__1.6',[]
        webElemento.click()
        self.__aguardarAcaoNavegador__(By.ID,'butAcao2')
        time.sleep(0.5)

        if dadoSin.urgente==True:
            funcionou,msgErro,codErro=self.__executarJavaScript__("document.all('LinkButton2').item(1).click()")
            if funcionou==False:
                self.closeNavegator()
                return funcionou,msgErro,codErro,'__registerSinsInKmt__1.7',[]
            time.sleep(0.5)

        conjScriptText=["document.getElementById('txtTipo').value='"+str(dadoSin.matServ).strip()+"'",
                        "document.getElementById('txtUnidadeMedida').value='"+str(dadoSin.unMed).strip()+"'"]

        posicao=0
        while posicao<2:
            funcionou,msgErro,codErro=self.__executarJavaScript__(conjScriptText[posicao])
            time.sleep(0.5)
            if funcionou==False:
                self.closeNavegator()
                return funcionou,msgErro,codErro,'__registerSinsInKmt__1.8',[]
            posicao+=1

        funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(By.PARTIAL_LINK_TEXT,'Classificações')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__1.9',[]
        webElemento.click()

        funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'ctl00_Body_tabCategorias_ucCategorias1_gvCategorias_ctl04_txtCat')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__2.0',[]
        time.sleep(0.5)

        funcionou,msgErro,codErro=self.__executarJavaScript__("document.getElementById('ctl00_Body_tabCategorias_ucCategorias1_gvCategorias_ctl04_txtCat').value='"+dadoSin.modalidade+"'")
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__2.1',[]
        time.sleep(0.5)

        funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(By.PARTIAL_LINK_TEXT,'Mídia')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__2.2',[]
        webElemento.click()
        time.sleep(0.5)

        funcionou,msgErro,codErro=self.__registerMidiaKmt__(dadoSin,qtdAnexosSin)
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__2.3',[]

        funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'ctl00_Body_tabCategorias_ucCategorias1_gvCategorias_ctl04_txtCat')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__2.4',[]
        time.sleep(0.5)

        grupoDoSin=str(dadoSin.grupo).strip()
        if grupoDoSin!='' and dadoSin.grupo!=None:
            funcionou,msgErro,codErro=self.__executarJavaScript__("document.getElementById('ctl00_Body_tabCategorias_ucCategorias1_gvCategorias_ctl02_txtCat').value='"+grupoDoSin+"'")
            if funcionou==False:
                self.closeNavegator()
                return funcionou,msgErro,codErro,'__registerSinsInKmt__2.5',[]

            subGrupoDoSin=str(dadoSin.subGrupo).strip()
            if subGrupoDoSin!='' and dadoSin.subGrupo!=None:
                funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'ctl00_Body_tabCategorias_ucCategorias1_gvCategorias_ctl03_txtCat')
                if funcionou==False:
                    self.closeNavegator()
                    return funcionou,msgErro,codErro,'__registerSinsInKmt__2.6',[]
                time.sleep(0.5)

                funcionou,msgErro,codErro=self.__executarJavaScript__("document.getElementById('ctl00_Body_tabCategorias_ucCategorias1_gvCategorias_ctl03_txtCat').value='"+subGrupoDoSin+"'")
                if funcionou==False:
                    self.closeNavegator()
                    return funcionou,msgErro,codErro,'__registerSinsInKmt__2.7',[]

        conjByTipo=[By.ID]
        conjTextElement=['butAcao2','butContinuar']

        posicao=0
        while posicao<len(conjTextElement):
            funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,conjTextElement[posicao])
            if funcionou==False:
                self.closeNavegator()
                return funcionou,msgErro,codErro,'__registerSinsInKmt__2.8',[]

            funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(By.ID,conjTextElement[posicao])
            if funcionou==False:
                self.closeNavegator()
                return funcionou,msgErro,codErro,'__registerSinsInKmt__2.9',[]
            webElemento.click()
            posicao+=1
            time.sleep(0.2)

        funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'butSim')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__3.0',[]

        funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(By.ID,'label_numeroSin')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__3.1',[]
        numeroSinCorrente=xp.substrRight(str(webElemento.text),6).strip()
        informacoesDoSin=[str(dadoSin.codSol),numeroSinCorrente, str(dadoSin.descricao).strip()]

        funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(By.ID,'butNao') 
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__3.2',[]
        webElemento.click()
        posicao+=1
        time.sleep(0.2)

        funcionou,msgErro,codErro=self.__aguardarAcaoNavegador__(By.ID,'lbutopcao')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__3.3',[]

        funcionou,msgErro,codErro,webElemento=self.__localizarWebElemento__(By.PARTIAL_LINK_TEXT,'Principal')
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'__registerSinsInKmt__3.4',[]
        webElemento.click()
        posicao+=1
        time.sleep(0.2)

        return True,'Sem Erro',0,'__registerSinsInKmt__Fim',informacoesDoSin
          

    def navegateThroughKlassmatt(self,urlKmt,modoNavegacaoOculto,tipoNavegador,loginKmt,senhaKmt,caminhoBaseDadosXls,nomeAbaPlanBaseDados='DADOS'):
        sinsCadastros=[]
        funcionou,msgErro,codErro,dadosParaCadastro,quantidadeRegistros,quantidadeAnexosPorRegistro=self.__obterDadosCadastro__()
        if funcionou==False:
            return funcionou,msgErro,codErro,'navegateThroughKlassmatt 1.0'
        elif quantidadeRegistros<=0:
            return False,'Planilha Vazia',conjuntoErros.pop('ErrBaseVazia'),'navegateThroughKlassmatt 1.1'

        funcionou,msgErro,codErro=self.navegateOnWeb(urlKmt,modoNavegacaoOculto,tipoNavegador)
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro,'navegateThroughKlassmatt 1.2'
        
        funcionou,msgErro,codErro=self.__singInSiteKmt__(loginKmt,senhaKmt)
        if funcionou==False:
            self.closeNavegator()
            return funcionou,msgErro,codErro       

        eachRegistro=0
        while eachRegistro<quantidadeRegistros-1:
            funcionou,msgErro,codErro,localizacaoDoErro,informacaoSinCadastrado=self.__registerSinsInKmt__(dadosParaCadastro[eachRegistro],quantidadeAnexosPorRegistro[eachRegistro])
            if funcionou==False:
                self.closeNavegator()
                return funcionou,msgErro,codErro,'navegateThroughKlassmatt 1.3'

            xp.xlExcluirLinhasCads(caminhoBaseDadosXls,nomeAbaPlanBaseDados)
            txtCreateListSinCad(informacaoSinCadastrado[1])
            eachRegistro+=1

        return funcionou,msgErro,codErro,'navegateThroughKlassmatt Fim'
