# Sales_Agreement_DATACLEAN

**USAGE**
- same_left_right_date(df)
   - populates blank cells in between two cells that have the same sales agreement date
     
- propagate_right(df)
   - populates blank cells at the end of a row. Identifies the rightmost cell populated and fills in blank cells to the right whose column date is prior to the agreement date identified

- choose_date(df)
  - populates blank cells in between cells that have different agreement dates and makes a decision between which date to choose (agreement date in left or right cell)

- needs_revisiion(df)
  - flags rows with a question mark and highlights the cells in them that needs revision
     
 - mark_done(df) **mainly for my own usage
   - marks rows as 'done' if there are no blank cells between the first populated cell and last populated cell

**HOW TO USE**
1. Export your .xlsx file
2. Copy the path to your file and assign it to 'file_input' in clean.py
3. Run the functions-- resulting .xlsx file will be exported in cleaned.xlsx
   
**ADDITIONAL COMMENTS**
- you will get a 'FutureWarning' when the program is running. This can be ignored-- allow the program to run until processed with exit code 0.
- cleaned.xlsx will be created locally in the project directory of this program automatically
- functions should be used one at a time for best results-- I used same_left_right, propagate_right, and then choose_date
  
