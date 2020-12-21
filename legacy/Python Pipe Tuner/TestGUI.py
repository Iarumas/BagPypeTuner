import PySimpleGUI as sg

sg.ChangeLookAndFeel('GreenTan')

form = sg.FlexForm('Pipe Tuner', default_element_size=(40, 1))

column1 = [[sg.Text('Column 1', background_color='#d3dfda', justification='center', size=(10,1))],
           [sg.Spin(values=('Spin Box 1', '2', '3'), initial_value='Spin Box 1')],
           [sg.Spin(values=('Spin Box 1', '2', '3'), initial_value='Spin Box 2')],
           [sg.Spin(values=('Spin Box 1', '2', '3'), initial_value='Spin Box 3')]]
layout = [
    [sg.Text('Reference Frequency', size=(20, 1), font=("Helvetica", 20))],
    [sg.InputText('Pitch A: ',size=(25,1))],
    [sg.InputText('Pitch Bb: ',size=(25,1))],
    [sg.Button(button_text='<<<'),sg.Button(button_text='<<'),sg.Button(button_text='<'),
     sg.Button(button_text='>'),sg.Button(button_text='>>'),sg.Button(button_text='>>>')],
    [sg.Checkbox('My first checkbox!'), sg.Checkbox('My second checkbox!', default=True)],
    [sg.Radio('My first Radio!     ', "RADIO1", default=True), sg.Radio('My second Radio!', "RADIO1")],
    [sg.Multiline(default_text='This is the default Text should you decide not to type anything', size=(35, 3)),
     sg.Multiline(default_text='A second multi-line', size=(35, 3))],
    [sg.InputCombo(('Combobox 1', 'Combobox 2'), size=(20, 3)),
     sg.Slider(range=(1, 100), orientation='h', size=(34, 20), default_value=85)],
    [sg.Listbox(values=('Listbox 1', 'Listbox 2', 'Listbox 3'), size=(30, 3)),
     sg.Slider(range=(1, 100), orientation='v', size=(5, 20), default_value=25),
     sg.Slider(range=(1, 100), orientation='v', size=(5, 20), default_value=75),
     sg.Slider(range=(1, 100), orientation='v', size=(5, 20), default_value=10),
     sg.Column(column1, background_color='#d3dfda')],
    [sg.Text('_'  * 80)],
    [sg.Text('Choose A Folder', size=(35, 1))],
    [sg.Text('Your Folder', size=(15, 1), auto_size_text=False, justification='right'),
     sg.InputText('Default Folder'), sg.FolderBrowse()],
    [sg.Submit(), sg.Cancel()]
     ]

button, values = form.Layout(layout).Read()
#sg.MsgBox(button, values)
print(button, values)