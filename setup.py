import sys, os

from cx_Freeze import setup, Executable

company_name = "faronics"
product_name = "DiallerAgent"

# build_exe_options = {"include_files": ["google_sheet_common.py"]}
build_exe_options = {
    "packages": ["subprocess", "zipfile", "datetime", "dropbox", "warnings", "pandas", "googleapiclient", "httplib2", "numpy",
                 "oauth2client", "time","tkinter","os","_tkinter"],
 	"optimize":"1", 
 	"force": "-f",
    "includes": ["idna.idnadata",'appdirs',"packaging","packaging.version", "packaging.specifiers",
    			 "packaging.requirements","tkinter","win10toast","dropbox"],
    "include_files": ["favicon.ico","Area Code_tz.xlsx","geckodriver.exe","storage_apakhare.json", "storage_robert.json", "google_sheet_common.py",
    "storage_afernandes.json","storage_joshua.json","common.py","browserlaunch.py", "storage.json", "storage_andysingh.json", "storage_abhay.json",
    "storage_spreadsheet_footprint.json"],
    }

# "optimize":"1",

# if sys.platform == 'win32':
#     base = 'Win32GUI'
# else:
#     base = None

os.environ['TCL_LIBRARY'] = r'C:\Users\admin\AppData\Local\Programs\Python\Python37\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\admin\AppData\Local\Programs\Python\Python37\tcl\tk8.6'

setup(
    name="Faronics Dialler",
    # publisher="Faronics Corporation",
    version="0.2",
    shortcutName="Faronics Dialler",
    description="Faronics Dialler - desc",
    options={"build_exe": build_exe_options},
    executables=[Executable("faronics_dialler.py",base="Win32GUI")]
)
