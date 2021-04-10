import PySimpleGUI as gui
#import tkinter as tkt
import WebPy as wp
import XlsPy as xp
import TxtPy as tp
import MolPy as mp
import ProgPy as pp
#####################################################################################################################################################################
caminhotxt='config.dat' 
caminhoExportSinsTxt='xbase\sindata.dat'
caminhoAcompanhaSinsTxt='xbase\sinBase.dat'
BasePath='xbase\_sinDataBaseAcomp.dat'
xlBasePathAcomp='xbase\_sinAcompBase.xlsx'
xlBasePathSin='xbase\_sinBase.xlsx'
verificadorContinuaCompilar=1
#####################################################################################################################################################################
try:    
    matrizConfigBasica=tp.LerTxt(caminhotxt)
except:
    gui.popup_error("Arquivo config.dat não encontrado!\n\nContate o suporte:\n\nE-mail: francisco.sousa1993@outlook.com\ntel: (21) 96965-6759",title='Erro: Código 700')
    verificadorContinuaCompilar=0


#####################################################################################################################################################################
def EnviarEmailsSinResp(matriz,totalSins,enviaAuto):
      assinatura = (r"""<P><font color='blue'><b>Atenciosamente,</b></font>
                              <br><font color='green'>Cadastro Klassmatt</font>
                              <br><font color='blue'><b>Gerência Geral de Frota</b>
                              <br>Av. General Justo, 375 - Edifício Bay View - 6º andar
                              <br>Centro - Rio de Janeiro - RJ - 20021-130</font>
                              <br> <a href='www.loginlogistica.com.br'>www.loginlogistica.com.br</a>
                              </p>""")
      
      verificadorMail=[]
      matAux=[]
      erros=0
      SinsErros=''
      matrizErros=[]

      comparador=matriz[0][0]
      cont=0
      for i in matriz:
          if comparador!=i[0]:
              matAux.append(cont-1)
              comparador=i[0]
          cont+=1
      matAux.append(cont-1) 
      cont=0
      verif=-1

      for i in matAux:
          tam=i-cont+1
          m=matriz[cont:(i+1)]              
              
          m=[['CODSOL','SIN','DESCRICAO']]+m          
  
          sinsMessage=GeraHtmlTable(tam+1,m) 
          stdMessage=(r"""<p><font color='blue'>Prezados,<br> Item(s) cadastrado(s), aguardando a geração do(s) código(s).<br>   <br> </font>""")
          stdMessage=stdMessage+sinsMessage+'<br>  </br> <br>  </br>'+assinatura
          codigoSolic=matriz[i][0]
          try: 
              verif=mp.ResponderEmail(codigoSolic,enviaAuto,stdMessage)
          except:
              erros+=1
              SinsErros+=codigoSolic + ';;'
              matrizErros+=matriz[cont:(i+1)]
          else:
              if verif==1:
                verificadorMail.append(codigoSolic)
              else:
                erros+=1
                SinsErros+=codigoSolic + ';;'
                matrizErros+=matriz[cont:(i+1)]

              
          cont=i+1          
      
      SinsErros=xp.substrLeft(SinsErros,xp.findLen(SinsErros)-2)
      return verificadorMail,erros,SinsErros,matrizErros

def EnviarEmailsSinProntos(matriz,totalSins,enviaAuto):
      assinatura = (r"""<P><font color='blue'><b>Atenciosamente,</b></font>
                              <br><font color='green'>Cadastro Klassmatt</font>
                              <br><font color='blue'><b>Gerência Geral de Frota</b>
                              <br>Av. General Justo, 375 - Edifício Bay View - 6º andar
                              <br>Centro - Rio de Janeiro - RJ - 20021-130</font>
                              <br> <a href='www.loginlogistica.com.br'>www.loginlogistica.com.br</a>
                              </p>""")
      
      verificadorMail=[]
      matAux=[]
      erros=0
      SinsErros=''
      matrizErros=[]

      comparador=matriz[0][0]
      cont=0
      for i in matriz:
          if comparador!=i[0]:
              matAux.append(cont-1)
              comparador=i[0]
          cont+=1
      matAux.append(cont-1) 
      cont=0
      verif=-1

      for i in matAux:
          tam=i-cont+1
          m=matriz[cont:(i+1)]              
              
          m=[['CODSOL','SIN','DESCRICAO','STATUS','CODIGO KLASSMATT','OBS','CODRET']]+m
          
  
          sinsMessage=GeraHtmlTable(tam+1,m) 
          stdMessage=(r"""<p><font color='blue'>Prezados,<br> Seguem os códigos dos itens cadastrados:<br>   <br> </font>""")
          stdMessage=stdMessage+sinsMessage+'<br>  </br> <br>  </br>'+assinatura
          codigoSolic=matriz[i][0]
          try: 
              verif=mp.ResponderEmail(codigoSolic,enviaAuto,stdMessage)
          except:
              erros+=1
              SinsErros+=codigoSolic + ';;'
              matrizErros+=matriz[cont:(i+1)]
          else:
              if verif==1:
                verificadorMail.append(codigoSolic)
              else:
                erros+=1
                SinsErros+=codigoSolic + ';;'
                matrizErros+=matriz[cont:(i+1)]

              
          cont=i+1          
      
      SinsErros=xp.substrLeft(SinsErros,xp.findLen(SinsErros)-2)
      return verificadorMail,erros,SinsErros,matrizErros
#####################################################################################################################################################################

def GeraHtmlTable(tam,matriz):
    tabela=r"<br /><p> <table border=5 CELLSPACING=2 CELLPADDING=6>"
    for i in range(tam):
        tabela=tabela + "<tr>"
        for j in range(1,len(matriz[0])):
            if i==0:
                tabela=tabela+"<td><b><font color='black'>" + str(matriz[i][j])+"</font></b></td>"
            else:
                tabela=tabela+"<td>" + str(matriz[i][j])+"</td>"

        tabela=tabela+"</tr>"
    tabela=tabela+"</table>"
    return tabela
#####################################################################################################################################################################

def TelaSin(linTotal,dados):

        #Layout:
        headers=['CODIGO','SIN','DESCRICAO']
        hlyt=[[gui.Text(h,size=(20,2) if h.title!='CODIGO' else (15,2),font=15,pad=(0,0)) for h in headers]]
        
        if linTotal!=0:
            linlyt=[
                        [gui.Input(
                                dados[i][j],
                                size=(75,2) if j==2 else ((25,2) if j==0 else (10,2)),
                                pad=(0,0),text_color="Green",
                                background_color='Black',
                                key=str("A"+str(i)+'x'+str(j))
                                ) for j in range(3)
                         ] for i in range(linTotal)]+[[gui.Text('')]]
        else:
            linlyt=[
                        [gui.Input(
                                '',
                                size=(75,2) if j==2 else ((25,2) if j==0 else (10,2)),
                                pad=(0,0),text_color="Green",
                                background_color='Black',
                                key=str("A"+str(i)+'x'+str(j))
                                ) for j in range(3)
                         ] for i in range(10)]+[[gui.Text('')]]
        hlyt=[[
                gui.Text('  ',size=(10,4)),
                gui.Text('KLASSMATT', font=65), 
                gui.Text('  ',size=(25,6)),
                gui.Text('Enviar solicitação por e-mail:',font=15),
                gui.Button('Enviar',key='btEnviar',button_color=('white','blue')),                
                gui.Checkbox("Automatico",key='ckAuto')

               ]] + hlyt + linlyt
        lyt=[[gui.Column(hlyt, size=(810,500), scrollable=True, vertical_scroll_only=True,expand_y=True)]]

        return lyt
#####################################################################################################################################################################
def TelaSinsAcomp(linTotal,dados):
    #Layout:
        headers=['Codigo','SIN','Descricao','Status','Codigo Klassmatt','Observação','Codigo Retorno']
        #hlyt=[[gui.Text(h,size=(19,2) if h.title=='Descricao' else (15,2),font='Arial 8',pad=(0,0)) for h in headers]]
        hlyt=[[gui.Text(headers[0],font='Arial 8',pad=(8,0),justification='left'),
               gui.Text(headers[1],font='Arial 8',pad=(25,0),justification='left'),
               gui.Text(headers[2],font='Arial 8',pad=(75,0),justification='left'),
               gui.Text(headers[3],font='Arial 8',pad=(30,0),justification='right'),
               gui.Text(headers[4],font='Arial 8',pad=(6,0),justification='left'),
               gui.Text(headers[5],font='Arial 8',pad=(50,0),justification='left'),
               gui.Text(headers[6],font='Arial 8',pad=(50,0),justification='left')
             ]]

        if linTotal!=0:
            linlyt=[
                        [gui.Input(
                                dados[i][j],
                                size=(37,2) if j==2 or j==5 else ((16,2) if j==4 else (10,2)),
                                pad=(0,0),
                                font='Arial 8',
                                text_color="Green",
                                background_color='Black',
                                key=str("MATSIN"+str(i)+'x'+str(j))
                                ) for j in range(7)
                         ] for i in range(linTotal)]+[[gui.Text('')]]
        else:
            linlyt=[
                        [gui.Input(
                                '',
                                font='Arial 8',
                                size=(37,2) if j==2 or j==5 else ((16,2) if j==4 else (10,2)),
                                pad=(0,0),
                                text_color="Green",
                                background_color='Black',
                                key=str("MATSIN"+str(i)+'x'+str(j))
                                ) for j in range(7)
                         ] for i in range(15)]+[[gui.Text('')]]
        hlyt=[[
                gui.Text('  ',size=(2,4)),
                gui.Text('KLASSMATT', font=65), 
                gui.Checkbox("Exibir Navegador",key='ckModoExibido'),
                gui.Text('  ',size=(8,6)),
                gui.Text('Verificar Sins:',font=12),
                gui.Button('        ',key='btVerificarAC',button_color=('white','green')),
                gui.Text('  ',size=(6,6)),
                gui.Text('Enviar codigos por e-mail:',font=12),
                gui.Button('Enviar',key='btEnviarAC',button_color=('white','blue')),                
                gui.Checkbox("Automatico",key='ckAutoAC')

               ]] + hlyt + linlyt
        lyt=[[gui.Column(hlyt, size=(810,500), scrollable=True, vertical_scroll_only=True,expand_y=True)]]

        return lyt
#####################################################################################################################################################################


class Tela:
    def __init__(self):
        #gui.theme('DarkAmber')
        gui.theme('Darkgrey9')
        self.__ExibNovamente__=0

        #Layout Menu:

        if verificadorContinuaCompilar==0:
            return

        
        lytMenu=[
                [gui.Text('  ')],
                [   
                    gui.Text('  '),
                    gui.Frame(
                                    ' Usuário: ',
                                    [
                                      [gui.Text('  ')],
                                      [
                                        gui.Text('  '),
                                        gui.Text('Login:',font=15),
                                        gui.Input(key='tbLogin',font=15,size=(25,10),default_text=matrizConfigBasica[2]),
                                        gui.Text(' ',size=(8,2)),gui.Text('Senha:',font=15),
                                        gui.Input(key='tbSenha',size=(25,10),font=15,password_char="*",default_text=matrizConfigBasica[3]),
                                        gui.Text('  ',size=(5,2))
                                      ],
                                      [gui.Text('  ')]
                                    ],
                                    font=15,

                              )
                    
                ],
                [gui.Text('  ')],
                [
                    gui.Text('  '),
                    gui.Frame(
                                    ' Localizar arquivo: ',
                                    [
                                        [gui.Text('  ')],
                                        [
                                            gui.Text('Caminho:', font=15),
                                            gui.Input(key="tbCaminho",disabled=True,background_color='red',text_color='black',size=(62,15),font=15),
                                            gui.Button('Importar',key='btImportar',size=(10,2),font=15,button_color=('white','darkblue'))                                                           
                                     
                                        ],
                                        [gui.Text('  ')]
                                    ],
                                    font=15
                        
                        )
                ],
                [gui.Text('  ')],
                
                [
                    gui.Text('  '),
                    gui.Frame(
                                    ' Executar: ',
                                    [
                                        [gui.Text('  ')],
                                        [
                                            gui.Text('Navegador:', font=15),
                                            gui.Combo(['Chrome','Mozila'],key='cbNavegador',default_value='Mozila',size=(15,6),font=20),
                                            gui.Text('  ',size=(1,2)),
                                            gui.Checkbox('Silencioso',key='ckSilencioso',font=15,default=True),
                                            gui.Text('  ',size=(20,2)),
                                            gui.Checkbox('Lembrar',key='ckLembrar',font=15),
                                            gui.Button('Ir',size=(10,2),key='btGo',font=15,button_color=('white','darkblue'))
                                     
                                        ],
                                        [gui.Text('  ')]
                                    ],
                                    font=15
                        
                        )
                ],
            ]

        #LAYOUT SINS RETORNO - CADASTRADOS 1
        mSins,tamLinha,tamColuna=xp.xlLerPlanilhaGeral(xlBasePathSin)
        lytSin=TelaSin(tamLinha,mSins)
        self.__totalSins__=tamLinha
        self.__lyt__=lytMenu

        #LAYOUT SINS ACOMPANHAMENTO - CADASTRADOS 2
        mSinsAcomp,LinhastamAcomp,colunasTamAcomp=xp.xlLerPlanilhaGeral(xlBasePathAcomp)
        lytSinAcomp=TelaSinsAcomp(LinhastamAcomp,mSinsAcomp)
        self.__totalAcomp__=LinhastamAcomp
        self.__lytAcomp__=lytSinAcomp

        #Layout Geral:
        lyt=[[gui.TabGroup([[gui.Tab('MENU',lytMenu),gui.Tab('SINS GERADOS',lytSin),gui.Tab('ACOMPANHAMENTO',lytSinAcomp)]])]]

        #Janela:
        
        self.wnd=gui.Window('LOGIN KLASSMATT',lyt)

    def __limparLayout__(self,tipoLayout):
        if tipoLayout==1:
            strCodMatriz='A'
            totalLinha=self.__totalSins__
            totalColuna=3
        else:
            strCodMatriz='MATSIN'
            totalLinha=self.__totalAcomp__
            totalColuna=7

        for linha in range(totalLinha):
            for coluna in range(totalColuna):
                #self.wnd[strCodMatriz+str(linha)+'x'+str(coluna)].Update('-')
                self.wnd.find_element(strCodMatriz+str(linha)+'x'+str(coluna)).update('')             

    def __qtdLinhasLayout__(self,tipoLayout):
        if tipoLayout==1:
            strCodMatriz='A'
            totalLinha=self.__totalSins__
            totalColuna=3
        else:
            strCodMatriz='MATSIN'
            totalLinha=self.__totalAcomp__
            totalColuna=7

        qtdLinhas=0
        for linha in range(totalLinha):
            if not self.val[strCodMatriz+str(linha)+'x'+str(1)] in ['',None]:
                qtdLinhas+=1
        return qtdLinhas

    def __obterDadosSinDimensao__(self,tipoLayout,tipoDimensao,numPosicaoDimensao, filtroValor=None,filtroDimensao=None):

        if tipoLayout==1:
            strCodMatriz='A'
            totalLinha=self.__totalSins__
            totalColuna=3
        else:
            strCodMatriz='MATSIN'
            totalLinha=self.__totalAcomp__
            totalColuna=7
        
        retornoVetorDimensao=[]        

        if tipoDimensao==1:
            for linha in range(totalLinha):
                if numPosicaoDimensao==linha+1:
                    for coluna in range(totalColuna):
                        if filtroValor==None and filtroDimensao==None:
                            retornoVetorDimensao.append(self.val[strCodMatriz+str(linha)+'x'+str(coluna)])
                        else:
                            if not self.val[strCodMatriz+str(linha)+'x'+str(filtroDimensao)] in filtroValor:
                                retornoVetorDimensao.append(self.val[strCodMatriz+str(linha)+'x'+str(coluna)])
        else:
            for coluna in range(totalColuna):                
                if numPosicaoDimensao==coluna+1:
                    for linha in range(totalLinha):
                        if filtroValor==None and filtroDimensao==None:
                            retornoVetorDimensao.append(self.val[strCodMatriz+str(linha)+'x'+str(coluna)])
                        else:
                            if not self.val[strCodMatriz+str(linha)+'x'+str(filtroDimensao)] in filtroValor:
                                retornoVetorDimensao.append(self.val[strCodMatriz+str(linha)+'x'+str(coluna)])



        return retornoVetorDimensao

    def __obterDadosLayout__(self,tipoLayout):
        if tipoLayout==1:
            strCodMatriz='A'
            totalLinha=self.__totalSins__
            totalColuna=3
        else:
            strCodMatriz='MATSIN'
            totalLinha=self.__totalAcomp__
            totalColuna=7
        
        retornoDadosLayout=[]
        for linha in range(totalLinha):
            retornoDadosLayout.append([])
            for coluna in range(totalColuna):
                retornoDadosLayout[linha].append(self.val[strCodMatriz+str(linha)+'x'+str(coluna)])

        return retornoDadosLayout
    
    def __obterDadosAcompNaoAprovados__(self):
        dadosAcomp=self.__obterDadosLayout__(2)
        retornoDadosNaoAprovados=[]
        for eachElement in dadosAcomp:
            if eachElement[6]!=1:
                retornoDadosNaoAprovados.append(eachElement)
        return retornoDadosNaoAprovados


    def __InserirRegistrolayout__(self,tipoLayout,dadosNovos,comecarInicio=False):

        if tipoLayout==1:
            strCodMatriz='A'
            totalLinha=self.__totalSins__
            totalColuna=3
        else:
            strCodMatriz='MATSIN'
            totalLinha=self.__totalAcomp__
            totalColuna=7
        if comecarInicio==False:
            linha=self.__qtdLinhasLayout__(tipoLayout)-1
        else:
            linha=0

        try:
            for eachElement in dadosNovos:
                coluna=0
                for eachValue in eachElement:
                    self.wnd.find_element(strCodMatriz+str(linha)+'x'+str(coluna)).update(eachValue)
                    coluna+=1
                linha+=1

        except Exception as e:
            print(e)

    def __atualizarlayout__(self,tipoLayout,dadosAtualizados):

        if tipoLayout==1:
            strCodMatriz='A'
            totalLinha=self.__totalSins__
            totalColuna=3
        else:
            strCodMatriz='MATSIN'
            totalLinha=self.__totalAcomp__
            totalColuna=7

        for eachElement in dadosAtualizados:
            for linha in range(totalLinha):
                if eachElement[1]==self.val[strCodMatriz+str(linha)+'x'+str(1)]:
                    for coluna in range(totalColuna):
                        self.wnd.find_element(strCodMatriz+str(linha)+'x'+str(coluna)).update(eachElement[coluna])        

        
    def __obterSinsNaoAprovados__(self,matrizSinsAcomp,excluirSinCancelado=False):
        matSinNaoAprovado=[]
        for eachElement in matrizSinsAcomp:
            if eachElement[6]!=1 and eachElement[4]!='APROVADO':     
                if excluirSinCancelado==True:
                    if eachElement[4]!='CANCELADO':
                        matSinNaoAprovado.append(eachElement)
                else:
                    matSinNaoAprovado.append(eachElement)

        return matSinNaoAprovado

    def Show(self):
      
        #conteudo1=['###BASE SINS:\n','###\n','###\n','###\n']
        #conteudo2=['### ARQUIVOS CADASTRADOS - EM ESPERA PARA ENVIO\n###\n###CODSOL	 SIN	DESCRICAO	STATUS      CODKLASSMATT		OBS\n###\n']
        while True:
            self.event,self.val=self.wnd.read()
            
            #Fechar programa
            if self.event==gui.WIN_CLOSED or self.event=='Cancel':
               break
            
           #Importar dados
            if self.event=='btImportar':
                file=gui.filedialog.askopenfilename()
                self.wnd.find_element('tbCaminho').update(file)              
               
            #Rotina de cadastro
            if self.event=='btGo':
                if str(self.val['tbLogin']).strip()=='' or str(self.val['tbSenha']).strip()=='':
                    gui.popup_error('Favor preencher todos os campos!',title='Atenção')
                elif str(self.val['tbCaminho']).strip()=='':
                    gui.popup_error('Favor selecionar o arquivo para importação!',title='Atenção')
                else:
                    login=str(str(self.val['tbLogin']).rstrip()).upper()
                    senha=str(self.val['tbSenha']).rstrip()
                    
                    if self.val['ckLembrar']==True:                       
                        tp.EditTxt(caminhotxt,'','',login,senha)
                    if self.val['cbNavegador']=='Mozila':
                        tpNav=2
                    else:
                        tpNav=1

                    try:                     
                        ret=wp.KmtCadastrarItens(self.val['tbCaminho'],'DADOS',tpNav,[login,senha],matrizConfigBasica[1],self.val['ckSilencioso'])                        
                    except:                        
                        ret=[5,'']
                    #Tratamento de erros:
                    
                    if ret[0]==0:
                        gui.popup_error('Usuário ou senha invalido!',title='Erro Login: Codigo 1001')
                    elif ret[0]==2:
                        gui.popup_error('Erro na URL ou autenticação do SITE Klassmatt.\n\nVerificar com o Suporte:\nE-mail: francisco.sousa@loginlogistica.com.br',title='Erro Acesso ao site: Código 1000')
                    elif ret[0]==1:
                        gui.popup_error('Erro em um objeto HTML:\n\n'+ret[1],title='Erro em objeto: Código 1003')
                    elif ret[0]==3:
                        gui.popup_error(ret[1],title='Erro midia anexa: Código 1004')
                    elif ret[0]==5:
                        gui.popup_error('Ocorreu um erro inesperado!\nTente novamente',title='Erro: Código 900')
                    elif ret[0] in [100,101,102,103,104]:
                        gui.popup_error(ret[1],title='Erro: Codigo ' + str(904+ret[0]))
                    else:
                        try:
                            xp.xlsExportSins(ret)
                            #Guardar SINS
                            
                            #tp.EditTxtBaseSins(caminhoExportSinsTxt,ret,conteudo1)
                            xp.xlInserirDadosPlanGeral(xlBasePathSin,ret)
                            #tp.saveTxtAcompanhamento(basePath,ret,True)
                            xp.xlInserirDadosPlanGeral(xlBasePathAcomp,ret)
                            
                            self.__ExibNovamente__=1
                        except:
                            gui.popup_error('Erro ao exportar os SINS! Operação abortada!','Erro: Código 800')
                        else:
                            gui.popup_ok('Finalizado com sucesso!!!',title='Informação',auto_close_duration=0.5)
                            break

            if self.event=='btVerificarAC':
               
             try:
                vetNumSins=self.__obterDadosSinDimensao__(2,2,2,[1],6)
                vetCodKmt=self.__obterDadosSinDimensao__(2,2,1,[1],6)
                self.__limparLayout__(2)
                

                if self.val['ckModoExibido']==True:
                    modoOculto=False
                else:
                    modoOculto=True

                try:
                    retConsultaSins,numErroConsulta,erroConsulta=wp.RetornarInfoSin(matrizConfigBasica[1],[matrizConfigBasica[2],matrizConfigBasica[3]],vetNumSins,vetCodKmt,modoOculto)
                except Exception as excecao:
                    gui.popup_error('Erros:\n'+str(excecao)+'\n'+str(erroConsulta),title='Erro: Codigo 18')

                matrizDadosAcomp=self.__obterSinsNaoAprovados__(retConsultaSins)
                vetNumSins=xp.obterVetColunaDeMatriz(matrizDadosAcomp,1)

                xp.xlAtualizarRegistroBaseSin(xlBasePathAcomp,vetNumSins,matrizDadosAcomp)
                xp.ordernarPlanAcomp(xlBasePathAcomp)

                matrizDadosAcomp,linha,Coluna=xp.xlLerPlanilhaGeral(xlBasePathAcomp)
                
                self.__InserirRegistrolayout__(2,matrizDadosAcomp,True)
                gui.popup_ok('Concluido!',title='informação!',auto_close_duration=0.2)

             except Exception as e:
                    gui.popup_error('Erro: '+ e.args[0],title="ERRO: CODIGO 125")
                    

            if self.event=='btEnviarAC': 

                matriz=[]
                
                self.wnd.TKroot.title('Carregando...')

                pp.progress(50,['Carregando..','Aguarde..'])             

                matriz=self.__obterDadosLayout__(2)
                matrizSinsNaoExluidos=self.__obterSinsNaoAprovados__(matriz)
                
                try:
                    verificadorMail,erros,sinsErros,matrizErros=EnviarEmailsSinProntos(matriz,self.__totalAcomp__,self.val['ckAutoAC'])
                    self.wnd.TKroot.title('LOGIN KLASSMATT')

                except:
                    gui.popup_error('Erro: Ao tentar enviar o email, ocorreu algo inesperado!','Erro: Código 399')
                else:
                    self.__limparLayout__(2)
                    
                    if erros!=0:

                        gui.popup_error('Erro: Email - houveram itens não encontrados! Verifique o codigo da solicitação:\n\nCodSol:\n\n'+str(sinsErros).replace(';;','\n'),title='Erro: Código 400')
                        i=0
                        
                        conteudoNovotxtAcomp=[]
                        for elemento in matrizErros:
                            conteudoNovotxtAcomp.append([])
                            for coluna in range(0,7):
                                self.wnd.find_element('MATSIN'+str(i)+'x'+str(coluna)).update(elemento[coluna])
                                conteudoNovotxtAcomp[i].append(self.val['MATSIN'+str(i)+'x'+str(coluna)])
                            i+=1
                    else:
                        gui.popup_ok('E-mail enviado com sucesso!!!',title='Informação',auto_close_duration=0.5)

                
                try:
                    matrizDadosAcomp=self.__obterDadosSinDimensao__(2,2) #Ajustar a partir daqui
                    xp.xlAtualizarRegistroBaseSin(xlBasePathAcomp,)
                    
                except:
                   gui.popup_error('Erro: Não foi possivel alterar o arquivo sinAcomp.Dat',title='Erro: Código 700')

               

            if self.event=='btEnviar':        
          
                self.wnd.TKroot.title('Carregando...')

                pp.progress(50,['Carregando..','Aguarde..'])             

                matriz=self.__obterDadosLayout__(1)
                
                try:
                    verificadorMail,erros,sinsErros,matrizErros=EnviarEmailsSinResp(matriz,self.__totalSins__,self.val['ckAuto'])                    

                except:
                    gui.popup_error('Erro: Ao tentar enviar o email, ocorreu algo inesperado!','Erro: Código 399')
                else:
                    self.__limparLayout__(1)
                    
                    if erros!=0:

                        gui.popup_error('Erro: Email - houveram itens não encontrados! Verifique o codigo da solicitação:\n\nCodSol:\n\n'+str(sinsErros).replace(';;','\n'),title='Erro: Código 400')
                        
                        self.__atualizarlayout__(matrizErros)
                        
                    else:
                        gui.popup_ok('E-mail enviado com sucesso!!!',title='Informação',auto_close_duration=0.5)

                self.wnd.TKroot.title('LOGIN KLASSMATT')
                try:
                    respostaExecucao, erroEmRemover=xp.xlRemoverRegistroBaseSin(xlBasePathSin,[''],True)
                    if respostaExecucao==0:
                        gui.popup_error("Ocorreu o seguinte erro:\n" + str(erroEmRemover),title='Erro: Código 138')
                    else:
                        gui.popup_ok('Concluido!!!',title='Informação',auto_close_duration=0.5)
                    if erros!=0:
                        xp.xlInserirDadosPlanGeral(xlBasePathSin,matrizErros)
                except:
                    gui.popup_error('Erro: Não foi possivel alterar o banco de dados SINBASE',title='Erro: Código 700')

