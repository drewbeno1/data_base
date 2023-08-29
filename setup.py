from setuptools import setup

APP = ['dataBasev01.pyw']
DATA_FILES = []
OPTIONS = {
 'iconfile':'database.ico',
 'argv_emulation': True,
 'packages': ['tkinter', 'openpyxl', 'pandas'] 
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)