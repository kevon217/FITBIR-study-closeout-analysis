### CONVERT PY FILES IN DIRECTORY TO IPYNB FILES AND CREATE A VERSIONED PACKAGED ###
    
    # This script converts .py files in a directory to .ipynb files and packages them into a custom directory based on user input for "package name" and "version number"


# LOAD PACKAGES

import os 
import shutil
import glob
from datetime import date
import tkinter as tk # for dialog boxes
from tkinter import filedialog
root = tk.Tk()
root.withdraw()

def create_folder(folder_path): # creates new folders every time script is run
    adjusted_folder_path = folder_path
    folder_found = os.path.isdir(adjusted_folder_path)
    counter = 0
    while folder_found == True:
        counter = counter + 1
        adjusted_folder_path = folder_path + ' (' + str(counter) + ')'
        folder_found = os.path.isdir(adjusted_folder_path)
    os.mkdir(adjusted_folder_path)
    return adjusted_folder_path

# INITIALIZE

date_version = date.today().strftime("%Y-%m-%d")
print('Please enter ipynb package name: ')
name_package = input()
print('Please enter ipynb package version #: ')
name_version = input()
dirname_package = f"{name_package}_ipynb_Version-{name_version}_{date_version}"

dirparent_output = filedialog.askdirectory(title = 'Choose parent directory for ipynb package output')
dir_package = create_folder(os.path.join(dirparent_output, dirname_package))
dir_scripts = filedialog.askdirectory(title = 'Choose directory with py scripts to be converted to ipynb')


# LOOP THROUGH SCRIPTS DIRECTORY --> CONVERT PY TO IPYNB --> COPY TO PACKAGE DIRECTORY

os.chdir(dir_scripts)
pyFiles = sorted(glob.glob('*.py'))
for py in pyFiles:
    fn_py, fn_ext = os.path.splitext(py)
    fn_ipynb = fn_py + '.ipynb'
    dir_closeout = os.system('cd ' + dir_scripts)
    os.system('p2j ' + '-o ' + py)
    shutil.copy(py, dir_package)
    shutil.copy(fn_ipynb, dir_package)
    
    

