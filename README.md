# Installation steps:

1. Make sure [Python](https://www.python.org/downloads/) is installed properly and works from the command line (e.g. run ```python --version``` in a command line terminal)
2. From a command line terminal, navigate to the folder where contents have been unzipped (e.g. ```cd C:\workspace\python-scripts-staffing```)
3. Run ```pip install -r requirements.txt```
4. Open a Python interpreter by running ```python```
5. At the Python interpreter prompt (>>>), run commands ```import nltk``` and then ```nltk.download('all')```
6. After downloading has finished, exit the Python interpreter by running ```exit()```
7. Make sure the staffing files 'BU_NL_OPEN_REQUEST_TO RESPOND {date}.xslx' are in the same folder as the scripts.
8. Run ```python bulk-keyword-search.py``` for the bulk keyword search.
9. To see a number of top common words, run ```python word-analysis.py``` after that.
10. Keep in mind that if you open 'output.csv' and/or 'output.xlsx', you might need to close them before running the scripts again.