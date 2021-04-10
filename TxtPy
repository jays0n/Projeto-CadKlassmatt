import XlsPy as xp
import random as rd

def sinHash(forma,valor):
    aux=['z','s','a','$','O','p',"w",'@','w',">",'<','#','%']
    if forma==1: #Oculta senha
        senHash=''        
        texto=str(valor)
        for elem in texto:
            senHash=str(aux[rd.randrange(len(aux))]) + senHash + str(hex(ord(elem)+3))+str(aux[rd.randrange(len(aux))])
        senHash=senHash.replace('x','u&')

    else: #Converte Senha
        texto=str(valor)
        for elem in aux:
            texto=str(texto.replace(elem,'')).strip()
        texto=texto.replace('u&','x')
        senHash=''
        
        for k in range(len(texto)):
            if (k+1)%4==0:
                elem=texto[(k-3):(k+1)]  
                senHash=senHash+chr(int(elem,16)-3)

    return senHash 


def LerTxt(caminho):
    ret=[]
    arquivo=open(caminho,'r')
    linhas=arquivo.readlines()

    for i in linhas:
        a=str(i)[9:]
        b=str(i)[0:9]
        
        if b=="@caminho@":  
            ret.append(a)
        elif b=="@urlsite@":
            ret.append(a)
        elif b=="@loginxx@":
            ret.append(a)
        elif b=="@senhaxx@":
            a=sinHash(2,a)
            ret.append(a)
    arquivo.close()
    return ret

def EditTxt(caminhoTxt,path,UrlSite,login,senha):

    arquivo=open(caminhoTxt,'r')
    linhas=arquivo.readlines()
    ret=[]

    for i in linhas:

        b=str(i)[0:9]

        if b=="@caminho@":  
            if path!='':
                ret.append(b+path+"\n")
            else:
                ret.append(i)
        elif b=="@urlsite@":
            if UrlSite!='':
                ret.append(b+UrlSite+"\n")
            else:
                ret.append(i)
            
        elif b=="@loginxx@":
            if login!='':
                ret.append(b+login+"\n")
            else:
                ret.append(i)
            
        elif b=="@senhaxx@":
            if senha!='':
                ret.append(b+sinHash(1,senha)+"\n")
            else:
                ret.append(i)
            
        else:
            ret.append(i)

    arquivo=open(caminhoTxt,'w')
    arquivo.writelines(ret)
    arquivo.close()


       
def LerTxtBaseSins(caminho):
    ret=[]
    arquivo=open(caminho,'r')
    k=0
    linhas=arquivo.readlines()

    for i in linhas:

        b=str(i)[0:3]

        if b!='###' and b!='' and b!='\n':  
            k+=1
            variavel=[]
            linhaElement=str(i).replace(';#;;#;',';#; - ;#;')

            for elemento in str(linhaElement).split(';#;'):

                if elemento=='' or elemento==None or elemento==' ' or elemento=='\n':
                    variavel.append(' - ')
                else:
                    variavel.append(elemento)
                
            ret.append(variavel)

    arquivo.close()
    ret.sort()
    return ret,k



def EditTxtBaseSins(caminho,dados,conteudoBaseArq):
    conteudo=conteudoBaseArq
    k=0
    ret=dados

   

    if ret!='':
        ret.sort()

        for elem in ret:
            k+=1
            aux=''
            for var in elem:
                aux=aux+str(var)+';#;'
            aux=xp.substrLeft(aux,xp.findLen(aux)-3)+'\n'
            conteudo.append(aux)
    
    arquivo=open(caminho,'w')

    arquivo.writelines('')  

    arquivo.writelines(conteudo)      

    arquivo.close()

    
    return k


def EditTxtGeral(caminho, novoConteudo):
    arquivo=open(caminho,'w')
    
    arquivo.writelines('')  
        
    arquivo.writelines(novoConteudo)      

    arquivo.close()

def EditTxtAcompanhamento(caminho,consultas, codKmt):
    conteudo=['###BASE SINS ACOMP:\n###\n###\n###\n']
    for i in range(len(codKmt)):
        aux=codKmt[i]+";#;"+consultas[i][0]+";#;"+consultas[i][1]+";#;"+consultas[i][2]+";#;"+consultas[i][3]+";#;"+consultas[i][4]+";#;"+consultas[i][5]+'\n'
        conteudo.append(aux)
    EditTxtGeral(caminho,conteudo)

def saveTxtAcompanhamento(caminho,conteudoNovo,materConteudoAnt=False):
    
    if materConteudoAnt==False:
        conteudo=['###BASE SINS ACOMP:\n###\n###\n###\n']        
    else:
        arquivo=open(caminho,'r')
        conteudo=arquivo.readlines()
        arquivo.close()
    
    k=0
    ret=conteudoNovo  

    if ret!='':
        ret.sort()

        for elem in ret:
            k+=1
            aux=''
            for var in elem:
                aux=aux+str(var)+';#;'
            aux=xp.substrLeft(aux,xp.findLen(aux)-3)+'\n'
            conteudo.append(aux)
    
    arquivo=open(caminho,'w')

    arquivo.writelines('')  

    arquivo.writelines(conteudo)      

    arquivo.close()
    
    return k

def txtCreateListSinCad(sinVal):
    caminho=xp.os.path.expanduser('~')+"\\Desktop\\SAIDATXTKLASSMATT.txt"
    try:
        arquivo=open(caminho,'r')
    except:
        arquivo=open(caminho,'w+')
        arquivo.close()
        arquivo=open(caminho,'r')
    
    conteudo=arquivo.readlines()
    conteudo.append('\n'+str(sinVal))
    arquivo.close()
    arquivo=open(caminho,'w')
    EditTxtGeral(caminho, conteudo)
