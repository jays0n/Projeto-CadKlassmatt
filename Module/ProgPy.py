import PySimpleGUI as gui
import random

def progress(val,textos):

    textList = textos
    progressDur = val

    layout = [[gui.Text(textList[0], key="progtext"),gui.Text(''),gui.Text('   0%', key="Kpercent",auto_size_text=True),gui.Text('%')],
              [gui.ProgressBar(progressDur, orientation='h', size=(20, 20), key='progbar')],
              [gui.Cancel()]]

    # create the Window
    #gui.theme('DarkAmber')
    window = gui.Window('Execução', layout)
    # loop that would normally do something useful
    for i in range(progressDur):
        # check to see if the cancel button was clicked and exit loop if clicked
        
        event, values = window.Read(timeout=0)

        window.Element('Kpercent').Update(str(round(100*i/progressDur,1)))
        if event == 'Cancel' or event is None:
            break
            # update bar with loop value +1 so that bar eventually reaches the maximum
        if i ==progressDur:
            window.Element('progtext').Update("Concluindo...")
        if i == progressDur * .2:
            window.Element('progtext').Update(textList[random.randrange(len(textList))])
        if i == progressDur * .3:
            window.Element('progtext').Update(textList[random.randrange(len(textList))])
        if i == progressDur * .4:
            window.Element('progtext').Update(textList[random.randrange(len(textList))])
        if i == progressDur * .5:
            window.Element('progtext').Update(textList[random.randrange(len(textList))])
        if i == progressDur * .6:
            window.Element('progtext').Update(textList[random.randrange(len(textList))])
        if i == progressDur * .7:
            window.Element('progtext').Update(textList[random.randrange(len(textList))])
        if i == progressDur * .8:
            window.Element('progtext').Update(textList[random.randrange(len(textList))])

        window.Element('progbar').UpdateBar(i + 1)
    # done with loop... need to destroy the window as it's still open
    window.Element('Kpercent').Update('100%')
    window.Close()
