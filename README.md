# Outlook topic summarizer
This project is a simple POC how NLTK, spaCy and gensim can be used to extract a summary of most common topics
from bunch of emails. Since win32com Python library has compatibility issues with Python 3.10.* I needed to 
extract .json file from emails by using VBA directly from Outlook. In order to use this POC you would need to:
1. Create a macros in Outlook by and use code inside ExportEmails.bas. 
2. Setup 24th line in ExportEmails.bas inside GetEmails subroutine to the proper path to your repository.
3. Execute the macros manually.
4. Execute Python script main.py

