import easygui
import win32com.client as win32
import os

options = ["New Well", "Edit User Profile", "WellSight Systems Location", "Exit"]
wellType = ["CPEC Bakken", "CPEC Upper Shaunavon", "CPEC Lower Shaunavon"]

button = easygui.buttonbox("Select one of the following:", "Well Files Creator 0.9.1", choices = options, default_choice="New Well", cancel_choice="Exit")
if button == options[0]:
  templateChoice = easygui.choicebox(msg='Select Well Type:', title='Well Files Creator', choices = wellType)
  # Start corresponding program depending on user selection
  if templateChoice == wellType[0]:
    # Start Bakken template
    os.startfile("wfcBakken.py")
  elif templateChoice == wellType[1]:
    # Start Upper Shaunavon template
    os.startfile("wfcUShaun.py")
  elif templateChoice == wellType[2]:
    # Start Lower Shaunavon template
    os.startfile('wfcLShaun.py')
  else:
    os.sys.exit(0)
  #os.startfile("wellfilescreator.exe") #Change to .exe for final build
elif button == options[1]:
  os.startfile("create_profile.py") #Change to .exe for final build
elif button == options[2]:
  os.startfile("choose_directory.py") #Change to .exe for final build
else:
  os.sys.exit(0)