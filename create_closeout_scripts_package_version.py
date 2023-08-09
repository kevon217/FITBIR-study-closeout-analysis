##### .PY --> .IPYNB / CLOSEOUT ANALYSIS SCRIPTS PACKAGE #####

# This script converts .py closeout analysis files to .ipynb from the command line and zips up the neccessary files for a closeout
# analysis package to be shared with Ops members 

### LOAD PACKAGES ###

import os 
import shutil
from datetime import date

### INITIALIZE ###

version_date = date.today().strftime("%Y-%m-%d")
version_name = 'Version-4.1'
name_package = 'Closeout-Analysis-Scripts_ipynb_' + version_name + '_' + version_date

dir_output = 'C:\\Users\\kevarmen1\\Documents\\BRICS\\Study_Closeout\\'
dir_package = os.path.join(dir_output, name_package)
os.mkdir(dir_package)
dir_scripts =  'C:\\Users\\kevarmen1\\Documents\\BRICS\\Sourcetree_BRICS\\Operations\\Closeout Analysis Scripts\\' # need \\ for cmd line (\ is escape character)

fn_py_reform = 'Reformat_QT_2_DR_validation_format.py'
fn_py_studycloseout = 'StudyCloseout.py'
fn_ipynb_reform = 'Reformat_QT_2_DR_validation_format.ipynb'
fn_ipynb_studycloseout = 'StudyCloseout.ipynb'

fn_py_faf = 'fitbir_analysis_functions.py'
fn_py_val = 'Analyze_Validation_Error_Log.py'
fn_word_temp = 'FITBIR Study Closeout Report_Template.docx'

### COMMAND LINE: CONVERT .PY --> .IPYNB ###

dir_closeout = os.system('cd ' + dir_scripts)
os.system('p2j ' + '-o ' + fn_py_reform)
os.system('p2j ' + '-o ' + fn_py_studycloseout)

### SAVE CLOSEOUT PACKAGE VERSION ###

shutil.copy(fn_ipynb_reform, dir_package)
shutil.copy(fn_py_reform, dir_package)
shutil.copy(fn_ipynb_studycloseout, dir_package) 
shutil.copy(fn_py_studycloseout, dir_package) 
shutil.copy(fn_py_faf, dir_package) 
shutil.copy(fn_py_val, dir_package) 
shutil.copy(fn_word_temp, dir_package) 


