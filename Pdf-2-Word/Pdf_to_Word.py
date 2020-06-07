# -*- coding: utf-8 -*-
"""
Created on Sun Jun  7 18:06:47 2020

@author: srira
"""

import autoit
import glob

program='Winword.EXE' # Find the winword.exe path in program files for MS Word

pdf_files=glob.glob('pdf/*.pdf') # Save the pdfs to be converted to this folder
pdf_files=[pdf.split('\\')[-1] for pdf in pdf_files]
for pdf in pdf_files:
    autoit.run(program+' '+pdf)
    autoit.win_wait_active("[CLASS:OpusApp]",100)
    autoit.send("{F12}")
    autoit.win_wait_active("Save As",20)
    autoit.control_click("Save As",'Button8')
    autoit.win_close("[CLASS:OpusApp]")