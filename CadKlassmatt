
import JanPy as jp
import webNetPy as wp

#gt=jp.Tela()
url=r'https://login.klassmatt.com.br/Login.aspx'
caminho=r"C:\Users\franc\Desktop\ENTRADAPADRAOKMT.xlsx"
myNet=wp.navegadorKmt(caminho,'DADOS')
funcionou,msgErro,codErro, linhaDoErro=myNet.navegateThroughKlassmatt(url,False,'1','francisco.sousa','Fjes1321','')

print(msgErro)


try:
    jp.pp.progress(100,['Iniciando..'])
    #gt.Show()

except:
    jp.gui.popup_error('Erro ao tentar criar o Layout, contate o administrador do programa',title='Erro 404')
    
while gt.__ExibNovamente__==1:
    jp.pp.progress(300,['Aguarde..'])
    gt=jp.Tela()
    try:
        gt.Show()
    except:
        jp.gui.popup_error('Erro ao tentar criar o Layout, contate o administrador do programa',title='Erro 404')
