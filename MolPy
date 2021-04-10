import win32com as wcp
import win32com.client as clt
import ProgPy as pp

ErrosOutlook={'OlEmailNaoEncontrado':0,'OlEmailOk':1}

def ResponderEmail(codSol,enviaAuto=False,messageMail=''):
    outlook= clt.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts= clt.Dispatch("Outlook.Application").Session.Accounts
    verificador=0

    pp.progress(120,['Conectando..'])
    
    #SETA CAIXA PRINCIPAL
    mainAcount=outlook.Folders(accounts[0].DeliveryStore.DisplayName)
    
    kmtAcount=outlook.Folders["Cadastro Klassmatt"]
    
    kmtInbox=kmtAcount.Folders["Inbox"]
    
    mainInbox=mainAcount.Folders['Inbox']    
    
    #LOOP - KLASSMATT:
    msg=kmtInbox.Items
    tam=msg.count-1
    i=tam
    
    message=None
   
    while i>tam-300:
        
        if str(str(msg[i].Subject).upper()).find(str(str(codSol).strip()).upper())>-1:
            message=msg[i]
            break
        i-=1
        

    if message!=None:
        responder=message.ReplyAll()
        responder.BodyFormat=2  # olFormatHTM
        responder.HTMLBody=messageMail+responder.HTMLBody
        responder.To=str(str(responder.To).replace("cadastro.klassmatt@loginlogistica.com.br","",1)).replace("Cadastro Klassmatt","",1)
        responder.CC=str(str(responder.CC).replace("cadastro.klassmatt@loginlogistica.com.br","",1)).replace("Cadastro Klassmatt","",1)
        responder.CC= responder.CC + "; cadastro.klassmatt@loginlogistica.com.br"
        

        if enviaAuto==True:
            responder.Send()
        else:
            responder.Display()

        verificador=1

    

    if verificador==0:

        #LOOP - MAIN INBOX:
        msg=mainInbox.Items
        tam=msg.count-1
        i=tam   

        message=None
        while i>tam-300:
            
            if str(str(msg[i].Subject).upper()).find(str(codSol).upper())>-1:
                message=msg[i]
                break
            i-=1
            

        if message!=None:
            responder=message.ReplyAll()
            responder.BodyFormat=2  # olFormatHTM
            responder.Body=messageMail+responder.Body
            responder.To=str(str(responder.To).replace("cadastro.klassmatt@loginlogistica.com.br","",1)).replace("Cadastro Klassmatt","",1)
            responder.CC=str(str(responder.CC).replace("cadastro.klassmatt@loginlogistica.com.br","",1)).replace("Cadastro Klassmatt","",1)
            responder.CC="cadastro.klassmatt@loginlogistica.com.br"
            

            if enviaAuto==True:
                responder.Send()
            else:
                responder.Display()

            verificador=1

    if verificador==0:
        return 0 #ErrosOutlook.pop('OlEmailNaoEncontrado')
    else:
        return 1 #ErrosOutlook.pop('OlEmailOk')
