import easygui, os
import win32com.client as win32
import user_profile

hztl_file_dir = easygui.fileopenbox(title = "Select the location for HORIZONTAL.LOG.exe", default = "C:\Program Files (x86)\WellSight Systems\starlog\*")
hztl_file_dir += " -h"
vert_file_dir = easygui.fileopenbox(title = "Select the location for STRIP.LOG.exe", default = "C:\Program Files (x86)\WellSight Systems\starlog\*")
vert_file_dir += " -s"

user_profile = open('lib\wellsight_dir.py', 'w') # Change back to wellsight_dir.py
user_profile.write("hztl_file_dir = r\"%s\"\n" % hztl_file_dir)
user_profile.write("vert_file_dir = r\"%s\"\n" % vert_file_dir)

os.startfile("wfc.exe") #Change back to main_menu.py