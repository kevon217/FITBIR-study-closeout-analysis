# Study Closeout Analysis

Data submitted to the FITBIR data repository must undergo a study closeout analysis before they are shared with the FITBIR community. The closeout analysis script checks for the following:

1. Demographic data have been submitted for all participants.
2. Duplicate entries.
3. Partial data (i.e., a subset of participants are missing values for a given data point).
4. Extra-validation scoring algorithms errors.
5. Correspondence between GUIDs and reported subject IDs (if submitted) throughout all submissions.

Overall, these checks ensure that data for each participant are consistent, properly reported, and as complete as possible. These detailed findings are reported to the study team and will be corrected by the study team if necessary before sharing the data.

## Steps for Closeout Analysis

### Step 1: Create Analysis Folder

- Create a new folder which will house your study's analysis, e.g., `Closeout_Analyses_Study_XYZ`.
- **NOTE**: DO NOT use the closeout scripts directories to house your particular study's data or analysis.

### Step 2: Download Study Data

#### FOR QUERY TOOL DATA:

- Download BOTH unflattened and flattened study data from the QT to your study's analysis folder in separate subfolders.
- It’s recommended you include in the subfolder name whether it’s unflattened or flattened, e.g., `Unflattened_Data` and `Flattened_Data`, and store both in a directory designated for the particular study’s closeout analysis, e.g., `Closeout_Analyses_Study_XYZ`.
- **NOTE**: No need to download associated files. You should still download imaging csvs so that our checks can process those as well. Imaging Ops will perform a more detailed analysis of image contents.

#### FOR DATA REPOSITORY DATA:

- Download all the needed submissions from the study's profile in the Data Repository to your study's analysis folder, e.g., `Closeout_Analyses_Study_XYZ`.
- It's recommended you store these in a subfolder named `Data_Repository_Data`.
- **NOTE**: By default, all submitted data from the Data Repository is already unflattened and you won’t be dealing with flattened data.
- **NOTE**: No need to download associated files.

### Step 3: Open your Python IDE

- Open your Jupyter notebook/lab or other Python IDE.
- Navigate to the `Closeout_Analysis_Scripts_VERSION_(#)_(date)` directory that you have created to keep the scripts as provided above.

### Step 4 (Optional): Reformat of the Unflattened Datasets to be run in the Validation Tool

- If you are planning on running the files through the Validation Tool (and they are from the Query Tool), you'll need to reformat them into the proper format for validation.
- Run `Reformat_QT_2_DR_validation_format.ipynb` or `.py`.
  - A dialog box opens up prompting you to select the directory where the unflattened QT results (or data repository files) are, e.g., `Closeout_Analyses_Study_XYZ\Unflattened_Data`.
  - A dialog box opens up asking if you want to filter by dataset IDs.
  - If yes, choose the csv file with a list of dataset IDs in the first column (1 per row) with the header being "Dataset". Example:

    ```
    Dataset
    FITBIR-DATA0007723
    FITBIR-DATA0007742
    FITBIR-DATA0007801
    FITBIR-DATA0007743
    FITBIR-DATA0007786
    ```

- **NOTE**: Make sure you also filter by the same dataset IDs when running the second script (`StudyCloseout.ipynb` or `.py`) or else the row/column information will not align properly.
- This script will create a subfolder called `CSV_files_validation` with the following contents:
  - `\Validation_CSVs` (folder): Contains reformatted QT data saved as CSVs files, which are used to run validation in the submission tool.
    - **SUGGESTION**: You should store your validation resultsDetails.txt file here, e.g., `CSV_files_validation\Validation_CSVs\resultDetails.txt`.
  - `\Reformatting_Log_timestamp.txt` (text file): This is a log file detailing steps of the reformatting script for your reference. You can open it up as excel and delimit by csv to browse through it more easily and filter. No need to share this with the study team.

### Step 5 (Optional): Validate form structures in Validation Tool

- Run validation in the submission tool.
  - In either the Webstart or Javascript submission tool, click browse and choose the `CSV_files_validation\Validation_CSVs` subfolder to load unflattened files.
  - Validate all form structures with extra-validation rules together (exclude imaging-related form structures).
  - **NOTE**: You may want to just run validation on all forms as an extra safety measure.
  - **NOTE**: Imaging files will always throw errors since (1) you won't be downloading the associated files and (2) the filepath specified in filepath data elements are relative to the submitter's local computer.
  - Export results as `resultDetails.txt` (or name of choice) to the same subfolder, e.g., `CSV_files_validation\Validation_CSVs`.
  - **NOTE**: Make sure you are downloading the results for all form structures validated, not just an individual one. In the Javascript tool, you will have to select all the form structures you validated before exporting the results.

### Step 6: Run Closeout Analysis

- Run `StudyCloseout.ipynb` or `.py`.
  - Follow the on-screen prompts to select directories, validate files, and filter by dataset IDs.
  - Enter study ID and study name.
  - **NOTE**: Ensure consistency in dataset IDs and file selections across all steps.

### Step 7: Analyze Closeout Results

- Inspect contents of the newly created `Closeout_Analysis` folder located in the folder designated as your closeout analysis, e.g., `Closeout_Analyses_Study_XYZ\Closeout_Analysis`.
- Review the provided files and templates for a comprehensive analysis and reporting.

### Step 8: Write Report and Send to Team

- Inspect the summary table.
- Complete `FITBIR Study Closeout Report_Template.docx`.
- Zip `Closeout_Analysis` and send to the team for review.
