# wellfilescreatorBakken.py - Populates wellfiles for CPEC Bakken wells and saves to directory

''' 
To Do List:
Add field for user entered month (maybe use a drop down menu and convert to number format) and sample point depth. 
Enter this info into striplogs: 
  Month in spud date and TD date (note: this will have to be changed by user if well crosses into another month)
  Depth Range (=Sample Point depth - 10 m)
  Sample Point memo on TVD and Build logs
Enter this info into report:
  Month in all blank dates using bookmarks in the word document template
  
Possibly gather prog depth info from Stick.xls file. This must be a non-mandatory option as user isn't always given the file. Use if statement here
'''

import easygui #import easygui for Windows boxes
from pywinauto import application, tests, findwindows, timings
import win32com.client as win32
import re
from os.path import abspath
import user_profile, wellsight_dir

# Define paths and window names for use with pywinauto
hztl_log_path = wellsight_dir.hztl_file_dir # Horizontal striplog path
vert_log_path = wellsight_dir.vert_file_dir # Vertical striplog path
hztl_login_title = u'Login - HORIZONTAL.LOG 5.3c'
vert_login_title = u'Login - STRIP.LOG 5.3c'
hztl_template_path = abspath("Templates\CPEC Bakken\Lateral Template.hlg")
build_template_path = abspath("Templates\CPEC Bakken\Vertical-Build Template.slg")
tvd_template_path = abspath("Templates\CPEC Bakken\TVD-Gamma Template.slg")
build_window_name = 'STRIP.LOG - Vertical-Build Template.slg'
tvd_window_name = 'STRIP.LOG - TVD-Gamma Template.slg'
wellpath_path = abspath("Templates\CPEC Bakken\Wellpath Template.xlsx")
isopach_path = abspath("Templates\CPEC Bakken\Isopach Template.xls")
report_path = abspath("Templates\CPEC Bakken\Report Template.docx")
hazard_path = abspath("Templates\CPEC Bakken\Hazard Template.docx")
vertbuild_header = "Vertical-Build"
tvd_header = "TVD-Gamma"
username = user_profile.username
geo2name = user_profile.geo2name

# Fix brackets for SendKeys
if "(" and "(" in hztl_template_path:
  hztl_template_path = hztl_template_path.replace("(", "{(}")
  hztl_template_path = hztl_template_path.replace(")", "{)}")
if "(" and "(" in build_template_path:
  build_template_path = build_template_path.replace("(", "{(}")
  build_template_path = build_template_path.replace(")", "{)}")
if "(" and "(" in tvd_template_path:
  tvd_template_path = tvd_template_path.replace("(", "{(}")
  tvd_template_path = tvd_template_path.replace(")", "{)}")

# Username bracket fix
if "(" and ")" in username:
  username_striplog = username
  username_striplog = username_striplog.replace("(", "{(}")
  username_striplog = username_striplog.replace(")", "{)}")
if "(" and ")" in geo2name:
  geo2_striplog = geo2name
  geo2_striplog = geo2_striplog.replace("(", "{(}")
  geo2_striplog = geo2_striplog.replace(")", "{)}")
  
# User prompt for the save location for the well files
save_dir = easygui.diropenbox(title = "Select the location to save your well files") + '\\'

# User prompt for the well information. Input goes into string list that can be accessed to populate well files
fieldNames = ['Well Name:', 'License:', 'U.W.I.:', 'AFE:', 'Surface Location:', 'Ground Elevation:', 'Kelly Bushing:'] # Define fieldNames, text that sits beside each user box
wellInfo = [] # List which holds the user-inputed data
wellInfo = easygui.multenterbox("Fill in Well Information Below", "Well Info", fieldNames)

# Used to break up the well name at the first digit. Puts into separate strings so it can populate the isopach spreadsheet
wellname_first = wellInfo[0]
wellname_end = ""
if wellInfo[0] is not "":
  wellname_break_search = re.search("\d", wellInfo[0])
  if wellname_break_search:
    wellname_break = wellname_break_search.start()
    wellname_first = wellInfo[0][:wellname_break-1]
    wellname_end = wellInfo[0][wellname_break:]
  
# Break up Surface Location string to input into Isopach Spreadsheet
lsd = ""
sec = ""
twnsp = ""
rng = ""
mer = ""
if wellInfo[4] is not "":
  well_loc_break = wellInfo[4].split("-")
  lsd = well_loc_break[0]
  sec = well_loc_break[1]
  twnsp = well_loc_break[2]
  # Get Meridian value, split at either w or W
  finder = ['w','W']
  for char in finder:
    rng_mer = well_loc_break[3].split(char)
    if len(rng_mer) >= 2:
      rng = rng_mer[0]
      mer_full = rng_mer[1]
      mer_search = re.search("\d", rng_mer[1])
      mer = mer_full[mer_search.start()]
# Round KB to 1 decimal place
if wellInfo[6] is not "":
  KB_rounded = round(float(wellInfo[6]), 1)
else:
  KB_rounded = 0

app = application.Application() # Define app for use in following functions

# Open WellSight Systems Horizontal and Vertical logs using pywinauto
def OpenWellSight(file_path):
  app.start_(file_path) # Start the app from function parameter
  #window_wait = timings.WaitUntilPasses(10, 0.5, lambda: app.window_(title=window_title)) # Wait until Login window opens then begin the below
  return

# Open and control Excel. Creates Wellpath and Isopach spreadsheets and saves them in user selected directory
def ControlExcel():
  excel = win32.gencache.EnsureDispatch('Excel.Application')
  excel.Visible = True
  # Open Wellpath Template
  wb_wp = excel.Workbooks.Open(wellpath_path)
  ws_wellpath = excel.Worksheets('Well Path')
  tb = ws_wellpath.Shapes.AddTextbox(1,220,5,250,30)
  tb.TextFrame2.TextRange.Characters.Text = wellInfo[0] + " Well Path Cross Section"
  tb.Line.Visible = False
  tb.Fill.Visible = False
  tb.TextFrame2.TextRange.Font.Size = 9.3
  tb.TextFrame2.TextRange.Font.Name = "Arial"
  tb.TextFrame2.TextRange.Font.Bold = True
  tb.TextFrame.HorizontalAlignment = 2
  ws_refdepth = excel.Worksheets('Reference Depths')
  ws_refdepth.Cells(2,3).Value = wellInfo[6]
  # Save Wellpath
  excel.ActiveWorkbook.SaveAs(save_dir + wellInfo[0] + " - Wellpath.xlsx")
  # Open Isopach Template
  wb_iso = excel.Workbooks.Open(isopach_path)
  ws_iso = excel.Worksheets('Isopach')
  ws_iso.Cells(1,3).Value = wellname_first
  ws_iso.Cells(2,3).Value = wellname_end
  ws_iso.Cells(3,3).Value = wellInfo[6] + " m"
  ws_iso.Cells(6,3).Value = lsd
  ws_iso.Cells(6,4).Value = sec
  ws_iso.Cells(6,5).Value = twnsp
  ws_iso.Cells(6,6).Value = rng
  ws_iso.Cells(6,7).Value = mer
  # Save Isopach
  excel.ActiveWorkbook.SaveAs(save_dir + wellInfo[0] + " - Isopach.xls")
  return

# Open and control Word. Creates Report and Hazard Assessment documents and saves them in user selected directory
def ControlWord():
  word = win32.gencache.EnsureDispatch("Word.Application")
  word.Visible = True
  report = word.Documents.Open(report_path)
  # Add Well Name and License to cover page
  report_bookmarks = ['WellNameCover1', 'WellNameCover2', 'LicenseCover', 'GeoNameCover', 'Geo2Name', 'PrepForCover', 'WellName', 'UWI', 'License', 'Ground', 'KB', 'Rig', 'KBGeoMark', 'WellNameSurveys', 'LocationSurveys', 'RigSurveys', 'KBSurveys', 'WellNameSummary', 'WellNameEnd', 'GeoNameEnd']
  report_datalist = [wellname_first, wellname_end, 'Well License #: ' + wellInfo[1], username, geo2name, user_profile.reporting_to, wellInfo[0], wellInfo[2], wellInfo[1], wellInfo[5] + ' m', wellInfo[6] + ' m', user_profile.rig, KB_rounded, wellInfo[0], wellInfo[4], user_profile.rig, KB_rounded, wellInfo[0], wellInfo[0], username]
  for i in range(20):
    doc_insert = report.Bookmarks(report_bookmarks[i]).Range
    doc_insert.InsertAfter(report_datalist[i])
  # Add Well Name to header
  word.ActiveDocument.Sections(1).Headers(win32.constants.wdHeaderFooterPrimary).Range.Text = wellInfo[0]
  # Save Report
  word.ActiveDocument.SaveAs(save_dir + wellInfo[0] + " - Report.docx")
  # Open Hazard Assessment Template
  hazard = word.Documents.Open(hazard_path)
  haz_name_insert = hazard.Bookmarks("WellName").Range
  haz_name_insert.InsertAfter(wellInfo[0])
  geo_name_insert = hazard.Bookmarks("GeoName").Range
  geo_name_insert.InsertAfter(username)
  geo2_name_insert = hazard.Bookmarks("Geo2Name").Range
  geo2_name_insert.InsertAfter(geo2name)
  word.ActiveDocument.SaveAs(save_dir + wellInfo[0] + " - Hazard Assessment.docx")
  # Save Hazard Assessment
  return
  
""" To control the horizontal striplogs. Gets past the login screen then opens up the template log for the Lateral log. 
It then goes into Edit -> Headers and enters in the information from the
main user prompt and tabs through the boxes then closes.""" 
def ControlHztl(login_title):
  timings.WaitUntilPasses(10, 0.5, lambda: app.window_(title=login_title))
  app[login_title].TypeKeys('~') # Press enter on login box to get to main window
  app[u'HORIZONTAL.LOG - [not saved]'].Wait('ready')
  app[u'HORIZONTAL.LOG - [not saved]'].MenuSelect("File -> Open")
  app[u'Open Log File'].Wait('ready')
  app[u'Open Log File'].TypeKeys(hztl_template_path + "{ENTER}", with_spaces=True)
  app[u'HORIZONTAL.LOG - Lateral Template.hlg'].Wait('ready') # Not sure if this is working
  app[u'HORIZONTAL.LOG - Lateral Template.hlg'].MenuSelect("Edit->Headers") # 'Edit' same as above, figure this out
  app[u'TMainHeaderForm'].Wait('ready')
  app[u'TMainHeaderForm'].TypeKeys(wellInfo[0] + " {(}Lateral{)}" + "{TAB}" + wellInfo[4] + "{TAB}" + wellInfo[1] + 
  "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}" + wellInfo[5] + "{TAB}" + wellInfo[6] +"%i", with_spaces=True)
  app[u'Other Information'].Wait('ready')
  # Add 'rig' to the below screen. Two tabs then one down arrow. Need code for down arrow in sendkeys
  app[u'Other Information'].TypeKeys("{TAB}Supervision:{SPACE}" + user_profile.companyman + "{ENTER}Rig:{SPACE}" + user_profile.rig + "{TAB}{TAB}" + user_profile.directional + "{ENTER}DD: " + user_profile.dd1 + "{ENTER}DD: " + user_profile.dd2 + "{ENTER}MWD: " + user_profile.mwd1 + "{ENTER}MWD: "
   + user_profile.mwd2 + "{TAB}{TAB}AFE:{SPACE}" + wellInfo[3] + "{ENTER}UWI:{SPACE}" + wellInfo[2] + "{TAB}" + "{ENTER}", with_spaces=True)
  app[u'TMainHeaderForm'].Wait('ready')
  app[u'TMainHeaderForm'].TypeKeys("%o")
  app[u'Operator'].Wait('ready')
  app[u'Operator'].TypeKeys("Crescent Point Energy Corp. {(}Prepared for{SPACE}" + user_profile.reporting_to + "{)}" + "{ENTER}", with_spaces=True)
  app[u'TMainHeaderForm'].Wait('ready')
  app[u'TMainHeaderForm'].TypeKeys("%g")
  app[u'Geologist'].Wait('ready')
  app[u'Geologist'].TypeKeys(username_striplog + ",{SPACE}" + geo2_striplog + "{ENTER}", with_spaces=True)
  app[u'TMainHeaderForm'].Wait('ready')
  app[u'TMainHeaderForm'].TypeKeys("%c")
  app[u'HORIZONTAL.LOG - Lateral Template.hlg'].MenuSelect("File -> Save As")
  app[u'Save Log File AsDialog'].TypeKeys(save_dir + wellInfo[0] + " - Lateral.hlg{ENTER}", with_spaces=True)
  return
  
""" To control the vertical striplogs. Gets past the login screen then opens up the template log for both the Vertical-Build
and the TVD-Gamma logs depending on what's called. It then goes into Edit -> Headers and enters in the information from the
main user prompt and tabs through the boxes then closes.""" 
def ControlVert(login_title, window_name, template_path, log_type):
  timings.WaitUntilPasses(10, 0.5, lambda: app.window_(title=login_title))
  app[login_title].TypeKeys('~')
  app['STRIP.LOG - [not saved]'].Wait('ready')
  app['STRIP.LOG - [not saved]'].MenuSelect("File -> Open")
  app[u'Open Log File'].Wait('active ready')
  app[u'Open Log File'].TypeKeys(template_path + "{ENTER}", with_spaces=True)
  app[window_name].Wait('ready')
  app[window_name].MenuSelect("Edit->Headers")
  app[u'TMainHeaderForm'].Wait('ready')
  app[u'TMainHeaderForm'].TypeKeys(wellInfo[0] + " {(}" + log_type + "{)}" + "{TAB}" + wellInfo[4] + "{TAB}" + wellInfo[1] + 
  "{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}" + wellInfo[5] + "{TAB}" + wellInfo[6] +"%i", with_spaces=True)
  app[u'Other Information'].Wait('ready')
  app[u'Other Information'].TypeKeys("{TAB}Supervision:{SPACE}" + user_profile.companyman + "{ENTER}Rig:{SPACE}" + user_profile.rig + "{TAB}{TAB}" + user_profile.directional + "{ENTER}DD: " + user_profile.dd1 + "{ENTER}DD: " + user_profile.dd2 + "{ENTER}MWD: " + user_profile.mwd1 + "{ENTER}MWD: "
   + user_profile.mwd2 + "{TAB}{TAB}AFE:{SPACE}" + wellInfo[3] + "{ENTER}UWI:{SPACE}" + wellInfo[2] + "{TAB}" + "{ENTER}", with_spaces=True)
  app[u'TMainHeaderForm'].Wait('ready')
  app[u'TMainHeaderForm'].TypeKeys("%o")
  app[u'Operator'].Wait('ready')
  app[u'Operator'].TypeKeys("Crescent Point Energy Corp. {(}Prepared for{SPACE}" + user_profile.reporting_to + "{)}" + "{ENTER}", with_spaces=True)
  app[u'TMainHeaderForm'].Wait('ready')
  app[u'TMainHeaderForm'].TypeKeys("%g")
  app[u'Geologist'].Wait('ready')
  app[u'Geologist'].TypeKeys(username_striplog + ",{SPACE}" + geo2_striplog + "{ENTER}", with_spaces=True)
  app[u'TMainHeaderForm'].Wait('ready')
  app[u'TMainHeaderForm'].TypeKeys("%c")
  app[window_name].MenuSelect("File -> Save As")
  app[u'Save Log File AsDialog'].TypeKeys(save_dir + wellInfo[0] + " - " + log_type + ".slg{ENTER}", with_spaces=True)
  return

OpenWellSight(hztl_log_path)
ControlHztl(hztl_login_title)
OpenWellSight(vert_log_path)
ControlVert(vert_login_title, build_window_name, build_template_path, vertbuild_header)
OpenWellSight(vert_log_path)
ControlVert(vert_login_title, tvd_window_name, tvd_template_path, tvd_header)
ControlExcel()
ControlWord()