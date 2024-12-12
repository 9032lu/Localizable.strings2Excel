import PySimpleGUI as sg
import json
sg.set_options(font=('Arial Bold', 16))
layout = [
   [sg.Text('Settings', justification='left')],
   [sg.Text('User name', size=(10, 1), expand_x=True),
   sg.Input(key='-USER-')],
   [sg.Text('email ID', size=(10, 1), expand_x=True),
   sg.Input(key='-ID-')],
   [sg.Text('Role', size=(10, 1), expand_x=True),
   sg.Input(key='-ROLE-')],
   [sg.Button("LOAD"), sg.Button('SAVE'), sg.Button('Exit')]
]
window = sg.Window('User Settings Demo', layout, size=(500, 200))
# Event Loop
while True:
   event, values = window.read()
   if event in (sg.WIN_CLOSED, 'Exit'):
      break
   if event == 'LOAD':
      f = open("settings.txt", 'r')
      settings = json.load(f)
      window['-USER-'].update(value=settings['-USER-'])
      window['-ID-'].update(value=settings['-ID-'])
      window['-ROLE-'].update(value=settings['-ROLE-'])
   if event == 'SAVE':
      settings = {'-USER-': values['-USER-'],
      '-ID-': values['-ID-'],
      '-ROLE-': values['-ROLE-']}
      f = open("settings.txt", 'w')
      json.dump(settings, f)
      f.close()
window.close()
