####### ANALYZE RESULT DETAILS FILE #######

def main(resultDetails_file):
    import pandas as pd
    import numpy as np
    import glob
    import os
    import csv
    from loguru import logger
    import re
    
    ####### INITIALIZE SETTINGS #######
    
    sep = "," # for separating log messages with comma for opening as csv and filtering
    logger.info('Running Analyze_Validation_Error_Log.py file...')
    val_dir_csvs = os.path.dirname(resultDetails_file)
    os.chdir(val_dir_csvs)

    ####### TRANSFORM 'resultDetails.txt' TO DATAFRAME #######
    
    log_lol = list(csv.reader(open(os.path.basename(resultDetails_file), 'r'), delimiter='\t'))
    ncols = max(len(elem) for elem in log_lol)
    if ncols == 4:
        df_val= pd.DataFrame(log_lol, columns = ['Empty','FS File','Result','Details'])
    elif ncols == 5:   
        df_val= pd.DataFrame(log_lol, columns = ['Empty','FS File','Result','Details','Empty_1']) # for error logs that have retired element warning
    else:
        df_val= pd.DataFrame(log_lol, columns = ['Empty','FS File','Result','Details','Empty_1','Empty_2'])
    df_val.fillna("",inplace=True)
    if "Empty_1" in df_val:
        df_val["Details"] = df_val["Details"] + df_val["Empty_1"]
    header_list =  ['FS File','Result','Details','GUID','Validation Row','All Submitted Data Column','Data Element','Issue'] # add additional columns for filtering on GUID/row/column/data element/issue
    df_val= df_val.reindex(columns = header_list)   
    df_val=  df_val.dropna(axis =0, how='all') # drop empty rows
    
    # get validation file names from Validation_csv directory + compare with resultDetails file
    val_files_dir = sorted(glob.glob('*.csv')) # in directory
    val_files_log = list(df_val['FS File'].unique()) # that were validated
    val_files = sorted(list(set(val_files_dir) & set(val_files_log))) # find intersection of files in directory and that were validated based on log

    # need to remove rows that have value "ALERT:"
    rmv = set(val_files_log) - set(val_files) 
    rmv = list(filter(None, rmv)) # remove empty set first
    idx_rmv = df_val[df_val['FS File'].isin(rmv)].index # remove rows with 'ALERT' message
    df_val= df_val.drop(labels=idx_rmv, axis=0).reset_index(drop = True) 
    
    # find row index ranges where val_file is located in resultDetails dataframe
    nrows = len(df_val)
    idx_val_files = df_val[df_val['FS File'].isin(val_files)]
    val_files = idx_val_files['FS File'].tolist() # matches indexes
    idx_start = idx_val_files.index.tolist()
    idx_end = idx_start.copy()
    idx_end.append(nrows)
    idx_end.pop(0)
    idx_range = list(zip(idx_start,np.array(idx_end)))
    
    # convert val_file name to FS name
    val_FS_name = []
    for file in val_files:
        formName = file.split('_validation.csv')[0]
        val_FS_name.append(formName)
        
    # create dictionary of validation files and idx_range
    dict_fn_index = dict(zip(val_FS_name,idx_range))
    
    # find # of errors for each FS / output results to Dataset_IDs csvs
    error_list = []; # keeps track of # errors per validated FS
    
    for i in range(0,len(dict_fn_index)):
         key = val_FS_name[i]
         logger.info(sep.join(['Extracting validation results for:', key]))
         idx_temp = dict_fn_index.get(key)
         df_val_filt = df_val.loc[idx_temp[0]:idx_temp[1]-1].copy() # will make edits to filtered version of df_val
         warn_idx = df_val_filt[df_val_filt['Result'] == 'WARNINGS'].index
         err_idx = df_val_filt[df_val_filt['Result'] == 'ERRORS'].index
         if len(warn_idx) == 0: # means there are NO warnings
             if len(err_idx) == 0:
                 n_errors = 0
             else:
                 n_errors = idx_temp[1] - err_idx[0] - 1 # verify this number
         else:
             if len(err_idx) == 0:
                 n_errors = 0
             else:
                 n_errors = warn_idx[0] - err_idx[0]-1
         error_list.append(n_errors)
         logger.info(sep.join(['Number of validation errors:', str(n_errors)]))
    
         # reorganize result details for spreadsheet
         if len(err_idx) != 0:
             if len(warn_idx) !=0:
                 df_val_filt.loc[err_idx[0]:warn_idx[0]-1, 'Result'] = "ERROR" # forward fill RESULT type for filtering
             else:
                 df_val_filt.loc[err_idx[0]:df_val_filt.index.stop, 'Result'] = "ERROR" # forward fill RESULT type for filtering
         if len(warn_idx) != 0:
             df_val_filt.loc[warn_idx[0]:, 'Result'] = "WARNING" 
         
         # parse resultDetails.txt file and paste guid/row/col/DE/issue to separate cols for easier filtering
         guid = re.compile(r"(?<=guid )(\w+)", re.IGNORECASE)
         row = re.compile(r"(?<=row )(\w+)")
         col = re.compile(r"(?<=column )(\w+)")
         data_element = re.compile(r'(?<=data element )"(.*?)"')
         issue_de = re.compile( r'(?<="\s)(.+)') # these errors have DE in quotes followed by string "is"
         issue_dup = re.compile( r'(?<=are\s)(.+)') # these duplicate errors have the string "are"
         issue_de_str = df_val_filt['Details'].str.extract(issue_de) 
         issue_dup_str = df_val_filt['Details'].str.extract(issue_dup)
         issue_de_rows_notnull = [index for index, row in issue_de_str.iterrows() if row.notnull().any()]
         issue_dup_rows_notnull = [index for index, row in issue_dup_str.iterrows() if row.notnull().any()]
           
         df_val_filt['GUID'] = df_val_filt['Details'].str.extract(guid)
         df_val_filt['Validation Row'] = df_val_filt['Details'].str.extract(row)
         df_val_filt['All Submitted Data Column'] = df_val_filt['Details'].str.extract(col)
         df_val_filt['Data Element'] = df_val_filt['Details'].str.extract(data_element)
         df_val_filt.loc[issue_de_rows_notnull,['Issue']] = issue_de_str.loc[issue_de_rows_notnull,0]
         df_val_filt.loc[issue_dup_rows_notnull,['Issue']] = issue_dup_str.loc[issue_dup_rows_notnull,0]
      
         
         # put parsed results into original df_val dataframe for use in study closeout analysis reporting
         df_val.iloc[dict_fn_index.get(key)[0]:dict_fn_index.get(key)[1],:] = df_val_filt
         
    dict_FS_errors = dict(zip(val_FS_name,error_list))
    
    return dict_fn_index, dict_FS_errors, df_val

if __name__ == "__main__":
    dict_FS_errors = main()
        
    