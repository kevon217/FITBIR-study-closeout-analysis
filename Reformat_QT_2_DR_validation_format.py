# RE-FORMAT FILES FOR USE IN VALIDATION TOOL  

# This script can take Query Tool results and re-format them so they can be run through submission tool for validation. It does the following:
# - Insert form structure name in cell A1
# - Adds "record" to cell B1
# - Keeps DatasetID instead of "x" as a record marker
# - Removes form structure name from data element headers e.g., BVMTRV1.0.Main.GUID --> Main.GUID
# - Saves Study ID and Dataset ID in a separate csv file so validation errors can be referenced by dataset ID
# - Creates logging file summarizing actions taken. 

# This script can also take Data Repository files and create a separate directory for validation and subsequent use in analysis script.

# This script can also take a csv file with a list of DatasetIDs to filter upon
# - Cell A1 should have "Dataset" as variable header 
# - Cells A2-AX should have the DatasetIDs you want to include (1 per row) e.g., FITBIR-DATA0007723

# Note: this will not allow imaging csvs to validate unless filepaths have been changed appropriately.

# Load packages
import pandas as pd
import numpy as np
import glob
import os
import chardet
import sys
import datetime
import tkinter as tk # for dialog boxes
from tkinter import messagebox
from tkinter import filedialog
import importlib
from loguru import logger # for logging script actions
logger.remove()

# Choose starting directory where reformat script files live
root = tk.Tk()
root.withdraw()
start_dir = filedialog.askdirectory(title = 'Select Starting directory where script lives')
os.chdir(start_dir)

# Load fitbir_analysis_functions 
import fitbir_analysis_functions as faf # contains custom functions for loading/processing FITBIR csv datafiles 
importlib.reload(faf)

# Select closeout analysis directory
import tkinter as tk
from tkinter import filedialog
root = tk.Tk()
root.withdraw()

# Select directory with Unflattened files
unflat_dir = filedialog.askdirectory(title = 'Select Directory with Unflattened Results') 
os.chdir(unflat_dir)

# QT or DR validation
qt_dr_1_0 = messagebox.askyesno("Python","Are your unflattened files from Query Tool (yes) or data repository (no)?")

# Create directories
val_dir = faf.create_folder(os.path.join(os.getcwd(),'CSV_files_validation'))
val_csv_dir = faf.create_folder(os.path.join(val_dir, 'Validation_CSVs'))

# Create log file
sep = ","  # for separating log messages with comma for opening as csv and filtering
logger_filepath = val_dir + "/Reformatting_Log_{time}.log" # need to figure out how to get rid of milliseconds in log file name 
logger.add(logger_filepath, format = "{time:YYYY-MM-DD HH:mm:ss},{level},{message}", backtrace = False, diagnose = False)
logger.add(sys.stderr, colorize = True, format = "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> <level>{message}</level>")
logger.info(sep.join(['CSV_files_validation path:', val_dir]))

# Filter by Dataset ID (if from QT)?
if qt_dr_1_0:
    filter_ID_1_0 = messagebox.askyesno("Python","Do you want to filter by dataset ID")
    if filter_ID_1_0:
        filter_ID_file =  filedialog.askopenfilename(title= 'Choose file with list of Dataset IDs for filter',initialdir = os.path.dirname(unflat_dir))  
        filter_ID_list = pd.read_csv(filter_ID_file,low_memory=False).iloc[:,0].dropna().unique().tolist()
        logger.info(sep.join(['DatasetID Filter file:', filter_ID_file]))
    else:
        logger.info('No data were filtered by Dataset IDs')
    
# REFORMATTING

max_col_excel = 16384 
max_row_excel = 1048576 
logger.info('STARTING CSV REFORMATTING...')
start_time = datetime.datetime.now() 
idx = 1 
if qt_dr_1_0: 
    # QT files loop
    filenames = sorted(glob.glob('query_result_*.csv'))
    for filename in filenames:
        formName = filename.split('query_result_')[1].split('_20')[0] # get form structure name. double check name was parsed correctly.
        logger.info(sep.join(['REFORMATTING UNFLATTENED QT CSV:', filename]))
        
        # Names of new files
        val_file = val_csv_dir + '\\' + formName + '_validation.csv' # add idx in cases where multiple forms with same formName
        
        # read in file and create dataframes for both files
        df = pd.read_csv(filename,low_memory=False, dtype= 'object')
        col_guid = [col for col in df.columns if '.GUID' in col][0]
        if filter_ID_1_0:
            df.loc[:,'Dataset'] = df.loc[:,'Dataset'].ffill() # forward fill dataset ID for filtering
            df = faf.id_filter(df, filename, filter_ID_list).reset_index(drop=True)
            rg_index = df.loc[pd.isna(df[col_guid]), :].index # find where nans are in GUID col so dataset ID can be removed later
            df.loc[rg_index,'Dataset'] = np.nan
        
        # Reformat unflat QT csv files function
        if df.empty:
            logger.warning(sep.join(['No validation csv or report files created. Form is empty:', filename]))
        else:
            df_val_csv = faf.unflat_reformat(df, formName)
            df_val_csv.to_csv(val_file, index = False, header = False) # write csv file
            logger.info(sep.join(['Validation csv file created:', filename]))
        
        idx = idx + 1 
       
else:
    # DR files loop
    filenames = sorted(glob.glob('*.csv'))
    for filename in filenames:
        logger.info(sep.join(['COPYING DATA REPOSITORY CSV TO VALIDATION DIRECTORY:', filename]))
        
        # names of new files
        val_file = os.path.join(val_csv_dir, filename) # don't append '_validation.csv'. this will keep DR filename consistent with formName given to validation csv file
        
        # read in file and create dataframes for both files
        with open(filename, 'rb') as f:
            result = chardet.detect(f.read()) # check character encoding
        df = pd.read_csv(filename,low_memory=False, header = None, encoding = result['encoding'], dtype= 'object')
        df_val_csv = df.copy()
        
        if df_val_csv.shape[0] > 2: # csvs with data will have > 2 rows
            # write csv file
            df_val_csv.to_csv(val_file, index = False, header = False)
            logger.info(sep.join(['Validation csv file created:', filename]))
        else:
            logger.warning(sep.join(['No validation csv or report files created. Form is empty:', filename]))   

# Close log file
logger.info(sep.join(['Script completion time:', str(datetime.datetime.now() - start_time)]))  
logger.remove()
os.chdir(start_dir)
