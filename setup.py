# setup.py - Create .exe from the scripts

from distutils.core import setup
import sys, os
import py2exe

setup(

    windows = [
        {
            "script": "main_menu.py",  
            "icon_resources": [(0, "mm.ico")],
            "uac_info": "requireAdministrator"     ### Icon to embed into the PE file.
        }
    ],
    zipfile="lib/bar.zip", 
    options={"py2exe": {"skip_archive": True,}}
)
setup(

    windows = [
        {
            "script": "choose_directory.py",  
            "icon_resources": [(0, "mm.ico")],
            "uac_info": "requireAdministrator"     ### Icon to embed into the PE file.
        }
    ],
    zipfile="lib/bar.zip", 
    options={"py2exe": {"skip_archive": True}}
) 

setup(

    windows = [
        {
            "script": "create_profile.py",  
            "icon_resources": [(0, "mm.ico")],
            "uac_info": "requireAdministrator"     ### Icon to embed into the PE file.
        }
    ],
    zipfile="lib/bar.zip", 
    options={"py2exe": {"skip_archive": True}}
) 
setup(

    windows = [
        {
            "script": "wfcBakken.py",  
            "icon_resources": [(0, "mm.ico")],
            "uac_info": "requireAdministrator"     ### Icon to embed into the PE file.
        }
    ],
    zipfile="lib/bar.zip", 
    options={"py2exe": {"skip_archive": True}}
) 
setup(

    windows = [
        {
            "script": "wfcUShaun.py",  
            "icon_resources": [(0, "mm.ico")],
            "uac_info": "requireAdministrator"     ### Icon to embed into the PE file.
        }
    ],
    zipfile="lib/bar.zip", 
    options={"py2exe": {"skip_archive": True}}
) 
setup(

    windows = [
        {
            "script": "wfcLShaun.py",  
            "icon_resources": [(0, "mm.ico")],
            "uac_info": "requireAdministrator"     ### Icon to embed into the PE file.
        }
    ],
    zipfile="lib/bar.zip", 
    options={"py2exe": {"skip_archive": True}}
) 