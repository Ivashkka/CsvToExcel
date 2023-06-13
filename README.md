# CsvToExcel
convert csv to excel and filter columns and rows by config txt files

This script can widely filter excel files or convert csv to excel and then filter it.

First of all script instantly deletes columns from excel listed in bad_columns.txt (replace data in this file to yours).

Then script filteres rows by custom configs.
To do this you need to create txt file in the same folder as script and call it <column_name>.txt (replace <column_name> with uors column name).
Then insert data you want from script to delete it from excel. For example:
    file Animal.txt:
    Cow
    Sheep
    Dolphin
This will remove all rows in excel wich contains string Cow or Sheep or Dolphin in column named Animal.
To delete all raws except ones with Cow / Sheep / Dolphin in column Animal you simple need to name txt file !Animal.txt
instead of Animal.txt.

Then you can mark columns to delete once again at the end for script by adding them into last_bad_columns.txt (replace data in this file to yours),
and then script will delete this columns at the end of filtering (previous step).

And the and of work you can pass to script names of colomns you want to rename by putting them into rename_columns.txt (replace data in this file to yours)<
and right before saving all data, script will rename columns as you wrote.

At the end of work script will save data to out.xlsx file.
