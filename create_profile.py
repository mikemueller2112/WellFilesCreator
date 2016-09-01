import easygui, os, user_profile

createprofile_field = ['Your name:', 'Professional Designation (ex, G.I.T./P.Geol etc.):', 'Second Geologist name:', 
'Second Geologist Professional Designation:', 'Rig Contractor (ex, Precision Drilling/Ensign etc.):', 'Rig #:', 'Company Man:', 
'Directional Service Company (ex, Phoenix, Pacesetter, etc.):', 'Directional Driller #1 Name:', 'Directional Driller #2 Name:', 
'MWD #1 Name:', 'MWD #2 Name:', 'Prepared for/Reporting to']
createprofile = []
createprofile = easygui.multenterbox("Create user profile by filling in the below:", "Create User Profile", createprofile_field, 
user_profile.createprofile)

username = createprofile[0] + " (" + createprofile[1] + ")"
geo2name = createprofile[2] + " (" + createprofile[3] + ")"
rig = createprofile[4] + " Rig #" + createprofile[5]
companyman = createprofile[6]
directional = createprofile[7]
dd1 = createprofile[8]
dd2 = createprofile[9]
mwd1 = createprofile[10]
mwd2 = createprofile[11]
reporting_to = createprofile[12]

user_profile = open('lib\user_profile.py', 'w') # Change back to user_profile.py
user_profile.write("username = \"%s\"\n" % username)
user_profile.write("geo2name = \"%s\"\n" % geo2name)
user_profile.write("rig = \"%s\"\n" % rig)
user_profile.write("companyman = \"%s\"\n" % companyman)
user_profile.write("directional = \"%s\"\n" % directional)
user_profile.write("dd1 = \"%s\"\n" % dd1)
user_profile.write("dd2 = \"%s\"\n" % dd2)
user_profile.write("mwd1 = \"%s\"\n" % mwd1)
user_profile.write("mwd2 = \"%s\"\n" % mwd2)
user_profile.write("reporting_to = \"%s\"\n" % reporting_to)
user_profile.write("createprofile = %s\n" % createprofile)

os.startfile("wfc.exe") #Change back to main_menu.py