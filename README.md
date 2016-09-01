README
======

------------------------
Well Files Creator 0.9.3
------------------------

Well Files Creator is designed to automate the creation of some standard well files used for wellsite geology personnel. It allows
for the user to input well-specific data (ex, Well Name, Surface Location, UWI, etc.) into one screen with the program populating
and creating three striplogs (Vertical-Build, TVD-Gamma, and Lateral), two spreadsheets (Wellpath and Isopach) and two
documents (Report and Hazard Assessment). These files will be saved in a folder selected by the user.

Copyright 2016 - Mike Mueller Geological Consulting Ltd.
========================================================

REQUIREMENTS
------------

Must have Wellsight Systems v5.3c and Microsoft Office installed.

INSTALLATION
------------

1. Run "Setup.exe"
2. Follow prompts for installation location
3. Enjoy!

INSTRUCTIONS
------------

1. After installation, run the program. You will be prompted with a screen with three options. "New Well", "Edit User Profile" and
   "WellSight Systems Location". To begin, you will need to create a User Profile and select the location of the WellSight Systems
   .exe files. Click on either to begin.

2. Select either "Edit User Profile" or "WellSight Systems Location".
    "Edit User Profile":
      - Fill in the required information. If you have previously entered information here, you can edit it and it will save it.
        You will need to enter the information as you would like it to appear on the well files. Enter your name, then your professional 
        designation (G.I.T., P.Geol, B.Sc., etc) based on how you would enter it into your striplog headers, report and hazard assessment. 
        It will combine these and put the professional designation in brackets as is required. Then enter Rig Contractor - for example, 
        Precision Drilling, Ensign, etc. and in the box below, enter the Rig #. These will also be combined to be entered in the correct 
        order on the well files. Enter the names for the personel as well as the directional company (this will be, for example, Pacesetter 
        Directional Drilling, Phoenix Technology Services, etc.). These will be filled out in the well files as well.
    "WellSight Systems Location"
      - After clicking this option, you will be prompted to select the location for WellSight Systems HORIZONTAL.LOG.exe and STRIP.LOG.exe.
        These are generally located in C:\Program Files (x86)\WellSight Systems\starlog\ (note: if you are running Windows 32 bit, it will be
        C:\Program Files\WellSight Systems\starlog\). If these locations ever change, you will have to point it to the new location.
 
3. After the above have been setup, you can select "New Well" to create the well files. First you will be presented with a menu where you can
   select the project you are working on. Here you will be greeted with a prompt asking you to choose the destination folder for the well files. 
   After this, you will be asked to input well data specific to your well. Input the data as you would in your report (ex, including all dashes, 
   backslashes, proper capitalization). Once you press "OK", it will cycle through and create your Lateral striplog, Vertical-Build striplog, 
   TVD-Gamma striplog, Wellpath and Isopach spreadsheets and then the Report and Hazard Assessment documents. These will be saved in the folder 
   you selected earlier.

   
NOTE: Once User Profile and WellSight Systems Location have been setup, you can immediately go into "New Well" and the data from your previous
entry will be saved. If anything changes with those, you will have to edit them and from that point it will be saved again.


OTHER INFORMATION
-----------------

You may wish to edit your templates to better fit your needs. This may include, editing the "blank" formation top texts 
(ex, xxx Lower Watrous(xx m TVD, xx m MSL)) or making changes to the "blank" bit information, etc. that you can fill out when necessary. In order to 
do this, you can go to the program file directory (default location is C:\Program Files (x86)\Well Files Creator 0.9.3\Templates. Here you will see 
the template files. It is important that you don't add in/fill in the information that you are prompted to as this may cause errors in the way the 
program runs. It is possible to add in information on the striplog (such as blank holders for formation tops/bit info, etc.) and this will not cause 
any issues with Well Files Creator. Make sure you save and do not change the filenames.

CHANGELOG v0.9.3
----------------

- Added menu to select the project user is drilling. Currently contains, CPEC Bakken, CPEC Upper Shaunavon, and CPEC Lower Shaunavon. This creates
the files specific to the project and the files are populated with the regular formation tops, etc.

CHANGELOG v0.9.2
----------------

- Added Second Geologist name to User Profile. This will input into the corresponding areas in the striplogs, report and hazard assessment.
- Reporting to Geologist will now be input into the striplogs


BUGS/CONTACT
------------

This program is still in Beta, there may still be issues with performance. Please notify me of any bugs, hiccups with performance
or any suggestions for further improvement.

mikemueller.geology@gmail.com
