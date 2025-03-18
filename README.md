# AO-Reporting-Program
Report creation module for Alpha One Coaching College.

## Code Usage
### Data Preprocessing
Module for reading input excels and aggregating into one Reading, Mathematical Reasoning, Thinking Skills and Writing (where applicable) dataframe. The format is standardised, with Writing having its own format.

### Missing Info
Class containing two main functions - return missing exam information and missing emails. 

These are outputted as excels into the working directory when requested by the user. 

### PDF Module
Defining extra PDF settings on top of the FPDF class and also adding extra functions to use when generating each pdf. Module mostly contains all functions that are run to generate pdf elements. This includes table layouts in the pdf, images, etc. 

### Report Creation
Report creation takes in the outputs from Data Preprocessing and transforms into a format for use in creating reports through an aggregate dataframe. The idea here is that every row contains information for a student on all info - including: 
- Name
- Email
- Individual Marks
- Aggregate marks
- Indiviudal Rankings
- Aggregate Ranking
- Marks for each question in each paper as columns (1 if correct, 0 if incorrect)

The program can then generate a report based on a each row of the dataframe. The dataframe is ordered such that students with complete information come first, and incomplete students (with missing) are at the bottom. 

### Send emails
Module to send emails of pdfs out the parents if there is an email attached to the row in the aggregate dataframe created in Report Creation.

### Program UI 
This is main program that gets run. UI layout, interactables and settings are defined in UIMainWindow.py, and this module handles the backend. Any changes to functions in the previous modules should flow through automatically through this script (as it just imports functions from those scripts), but any changes to UI need to be updated in UIMainWindow. Any new functionality that needs to be attached to the added UI needs to then be connected in this module. 

### To Run
Ensure that you're running Python3.10.

For Mac:
```
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```
To compile into exe:
```
pyinstaller --onefile --add-data "pdf_resources:pdf_resources" --add-data "percentile_bands:percentile_bands" program_ui.py
```
Edit .ui file with:
```
designer
```