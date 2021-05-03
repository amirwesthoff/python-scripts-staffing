import sys
import glob
import os
import pandas as pd
import warnings

warnings.filterwarnings("ignore")

# ----- GLOBAL VARIABLES ------

OUTPUT_CSV = './output.csv'
OUTPUT_XSLX = './output.xlsx'
COLUMN_HEADERS = [
    'File',
    'Keyword(s) Used',
    'Team Request Id',
    'Team Request Name',
    'Position Name',
    'Role Notes',
]

COLUMNS = ['', 'Team Request Name', 'Position Name', 'Role Notes']

# ----- HELPER FUNCTIONS ------


def checkForExistingAndRemove():
    if os.path.isfile(OUTPUT_CSV) or os.path.isfile(OUTPUT_XSLX):
        print('File "output.csv" and/or "output.xlsx" already exist(s) in this directory. Continuing will remove these files.')
        check = input('Do you want to continue? (y/n) : ')
        if check == 'n'.lower():
            sys.exit()
        if check == 'y'.lower():
            print('Removing files...')
            if os.path.isfile(OUTPUT_CSV):
                os.remove(OUTPUT_CSV)
            if os.path.isfile(OUTPUT_XSLX):
                os.remove(OUTPUT_XSLX)


def createOutputFile():
    print('Creating output.csv...')

    empty_df = pd.DataFrame(columns=[COLUMN_HEADERS])
    empty_df.to_csv(OUTPUT_CSV, index=False)


def processFiles():
    print('Processing files...')

    file_list = glob.glob("files/*.xlsx")
    num_files = len(file_list)

    counter = 0

    for file in file_list:
        df = pd.read_excel(file, header=5)

        selection = df[df[COLUMNS[CHOICE_COLUMN]].str.contains(CHOICE_QUERY)]

        if not selection.empty:
            hits = selection.shape[0]
            counter += hits
            print('Found %d hit(s)...' % hits)

        selection.insert(0, COLUMN_HEADERS[1], CHOICE_QUERY)
        selection.insert(0, COLUMN_HEADERS[0], file)

        selection['Role Notes'].replace(to_replace=[r"_x000D_"], value=[
            ""], regex=True, inplace=True)

        modified_selection = selection[COLUMN_HEADERS]

        modified_selection.to_csv(
            OUTPUT_CSV, index=False, header=False, mode='a', encoding='utf-8')

    print('Creating output.xlsx...')
    csv = pd.read_csv(OUTPUT_CSV)
    csv.to_excel(OUTPUT_XSLX, index=False)

    print('Processing completed. %d file(s) processed, %d hit(s) found for keyword(s) "%s" in column "%s".' %
          (num_files, counter, CHOICE_QUERY, COLUMNS[CHOICE_COLUMN]))


# ----- MAIN ------

print('\n')
print('Bulk Keyword Search through BU_NL_OPEN_REQUEST_TO RESPOND **-**-****.xlsx')
print('-------------------------------------------------------------------------')
print('\n')
print('!!! Make sure the files through which you want to search are in the folder named "files".')
print('!!! This script results in two files: "output.csv" and "output.xlsx".')
print('!!! The keyword search is case sensitive.')
print('!!! Use the pipe operator for multiple keywords (e.g. "Jira|requirements management"), although finding one is enough for a match.')
print('!!! Type CTRL + C to break the program at any time.')
print('\n')

CHOICE_COLUMN = int(
    input('On which column to do you want to perform the search (1 - Team Request Name, 2 - Position Name, 3 - Role Notes)? (1/2/3) '))
CHOICE_QUERY = input(
    'Search for keyword(s): ')

checkForExistingAndRemove()

createOutputFile()

processFiles()
