# Sales_Agreement_DATACLEAN

**USAGE**
Finds cells in a row with the same sales agreement date and populates the blank cells in between these two cells with that agreement date. 

**HOW TO USE**
1. Export your .xlsx file
2. Copy the path to your file and assign it to 'file_input' in clean.py
3. Run the program-- blank cells in between cells with the same left and right dates should be populated and resulting .xlsx file will be exported in cleaned.xlsx
   
**ADDITIONAL COMMENTS**
- you will get a 'FutureWarning' when the program is running. This can be ignored-- allow the program to run until processed with exit code 0.
- cleaned.xlsx should be created locally in the project directory of this program
  
