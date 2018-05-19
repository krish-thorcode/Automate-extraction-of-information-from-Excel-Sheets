# Handling-Excel-files-with-Python
Automating the boring extraction of information from huge excel files of VIT time table

*Python third-party modules required:* openpyxl

*Usage:*

1. for both summer_sem.py and summer_sem_updated.py

  ----------these two scripts assume that the sheet in the input xlsx file has non empty values atleast in 2nd row--------

  python `<python script name with extension>` `<xlsx file name with extension>` `<sheet name inside xlsx file>` `<day required>`

  example:
  python summer_sem_updated.py sample_summer_sem.xlsx 'Input Data' Tuesday

2. for both normal_sem.py and normal_sem_updated.py

  -----------these two scripts can run on both empty as well as non empty sheets-----------

  ----for a specific day
  python `<python script name with extension>` `<xlsx file name with extension>` `<sheet name inside xlsx file>` `<day required>`

  example:
  python normal_sem_updated.py sample_normal_sem.xlsx renamed_sheet Monday

  ----for all day_slots
  python `<python script name with extension>` `<xlsx file name with extension>` `<sheet name inside xlsx file>` all

  example:
  python normal_sem_updated.py sample_normal_sem.xlsx renamed_sheet all
