import win32com as wcp
import win32com.client as clt
import ProgPy as pp
from PyDefines import *

class pyEmail:
    def __init__(self,codSol=STRNULL,sendAuto=FALSE,messageMail=STRNULL):
        self.outlook=clt.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.accounts= clt.Dispatch("Outlook.Application").Session.Accounts
        self.mainAcount=self.outlook.Folders(self.accounts[0].DeliveryStore.DisplayName)
        self.kmtAcount=self.outlook.Folders["Cadastro Klassmatt"]
        self.kmtInbox=self.kmtAcount.Folders["Inbox"]
        try:
            self.mainInbox=self.mainAcount.Folders['Inbox']   
        except:
            self.mainInbox=self.mainAcount.Folders['Caixa de Entrada']  

        self.codSol=codSol.upper().strip()
        self.sendAuto=sendAuto
        self.messageMail=messageMail

    def getEmail(self,code,withAttachment=FALSE,attachmentMustBeXls=FALSE):
        emailsKmt=self.kmtInbox.items
        lenghtKmt=emailsKmt.count-1
        contKmt=lenghtKmt

        emailsMain=self.mainInbox.items
        lenghtMain=emailsMain.count-1
        contMain=lenghtMain

        emailItem=NULL
        cont=0
        try:
            found=FALSE
            while cont<500:
                i=contKmt-cont
                subject=str(emailsKmt[i].Subject).upper().strip()
                if subject.find(self.codSol)>EOF:
                    emailItem=emailsKmt[i]
                    if withAttachment==TRUE:
                        if emailItem.attachments.count>0:
                            if attachmentMustBeXls==TRUE:                                
                                for eachItem in emailItem.attachments:
                                    if str(eachItem.FileName).find('xls')>EOF:
                                        found=TRUE
                                        break
                                if found==TRUE:
                                    break
                            else:
                                break
                    else:                        
                        break

                i=contMain-cont
                subject=str(emailsMain[i].Subject).upper().strip()
                if subject.find(self.codSol)>EOF:
                    emailItem=emailsKmt[i]
                    if withAttachment==TRUE:
                        if emailItem.attachments.count>0:
                            if attachmentMustBeXls==TRUE:
                                for eachItem in emailItem.attachments:
                                    if str(eachItem.FileName).find('xls')>EOF:
                                        found=TRUE
                                        break
                                if found==TRUE:
                                    break

                            else:
                                break
                    else:                        
                        break
                cont+=1
        except Exception as msgErr:
            return FALSE,ERR_EMAIL_NOT_FOUND,msgErr,NULL
        if cont>=500 or emailItem==NULL:
            return FALSE,ERR_EMAIL_NOT_FOUND,'Email não encontrado',NULL
        
        return TRUE,ERR_NO_ERRO,STRNULL,emailItem

    def saveAttachement(self):
        response=self.getEmail(self.codSol,TRUE,TRUE)
        if response[0]==FALSE:
            return response
        emailItem=response[3]
        attachments=emailItem.attachments

        if attachments.count<=0:
            return FALSE,ERR_WITHOUT_ANY_ATTACHMENT,'Sem anexo',STRNULL

        for eachItem in attachments:
            if str(eachItem.FileName).find('xls')>EOF:
                break
        path=PATH_STANDARD_TEMP_DIRECTORY+"\\"+eachItem.FileName
        formatFile=path[-4::]
        if formatFile=='.xls':
            path=path+'x'

        try:
            eachItem.SaveAsFile(path)
        except Exception as msgErr:
            return FALSE,ERR_NO_PERMISSION_ACCESS_OUTLOOK,msgErr,STRNULL
        return TRUE,ERR_NO_ERRO,STRNULL,path

    def sendEmail(self):      

        response=self.getEmail(self.codSol)
        if response[0]==FALSE:
            return response[0:3]

        emailItem=response[3]

        replyEmail=emailItem.ReplyAll()
        replyEmail.BodyFormat=2
        replyEmail.HTMLBody=self.messageMail+replyEmail.HTMLBody
        replyEmail.To=str(str(replyEmail.To).replace("cadastro.klassmatt@loginlogistica.com.br","",1)).replace("Cadastro Klassmatt","",1)
        replyEmail.CC=str(str(replyEmail.CC).replace("cadastro.klassmatt@loginlogistica.com.br","",1)).replace("Cadastro Klassmatt","",1)
        replyEmail.CC= replyEmail.CC + "; cadastro.klassmatt@loginlogistica.com.br"

        try:
            if self.sendAuto==TRUE:
                replyEmail.Send()
            else:
                replyEmail.Display()
        except Exception as msgErr:
            return FALSE,ERR_EMAIL_NOT_SEND,msgErr

        return TRUE,ERR_NO_ERRO,STRNULL

    def __changeFromTextToHTML__(self,table):

        tableText=r"<br /><p> <table border=5 CELLSPACING=2 CELLPADDING=6>"
        tableText+="<tr>"
        for eachField in table[0].keys():
            tableText+="<td><b><font color='black'>"+str(eachField).upper()+"</font></b></td>"        

        postion=0
        for eachDct in table:
            tableText+="<tr>"
            for position in range(len(eachDct)):
                tableText+="<td>" + str(list(eachDct.values())[position]).upper() + "</td>"
            tableText+="</tr>"
        tableText+="</table>"
        return tableText

    def setText(self,table):
        self.messageMail+=self.__changeFromTextToHTML__(table)
        self.messageMail+=EMAIL_SIGNATURE
