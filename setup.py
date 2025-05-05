import sys
from cx_Freeze import setup, Executable


base = "Console"

options = {
       'build_exe': {
           'includes': ['selenium', 'openpyxl','time'],
           'include_files': [],
       }
   }


executables = [Executable('scraper1.py', base=base,icon="OIP.ico")]

setup(name='Scraper',
      version='1.2',
      description='no description',
      options=options,
      executables=executables)
