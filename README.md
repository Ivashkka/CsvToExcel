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
    Dolphin (each item on a new line).
This will remove all rows in excel wich contains string Cow or Sheep or Dolphin in column named Animal.
To delete all raws except ones with Cow / Sheep / Dolphin in column Animal you simple need to name txt file !Animal.txt
instead of Animal.txt.

Then you can mark columns to delete once again at the end for script by adding them into last_bad_columns.txt (replace data in this file to yours),
and then script will delete this columns at the end of filtering (previous step).

And the and of work you can pass to script names of colomns you want to rename by putting them into rename_columns.txt (replace data in this file to yours)<
and right before saving all data, script will rename columns as you wrote.

At the end of work script will save data to out.xlsx file.

------------------------------------------------------------------------------------------------------------------------------------------------------------------

конвертировать csv в excel и отфильтровать содержимое с использованием конфигов txt

В начале скрипт удаляет все ненужные колонки перечисленные в bad_columns.txt (замените данные в этом файле на свои).

Далее скрипт фильтрует все строки и удаляет ненужные по ключевым значениям, которые вы указываете в файле txt,
который нужно создать в дирректории скрипта и назвать его <имя колонки>.txt (замените <имя колонки> на свое имя).
Например файл с именем Animal.txt и содержимым:
    Cow
    Sheep
    Dolphin (каждый элемент на новой строке).
Это удалит все строки в таблице, в которых есть значение Cow или Sheep или Dolphin в колонке Animal.
Для того, чтобы наоборот удалить все, но оставить строки с содержимым Cow / Sheep / Dolphin просто назовите файл !Animal.txt вместо Animal.txt.

Далее можно еще раз указать колонки для удаления уже после фильтрации в файле last_bad_columns.txt для того, чтобы скрипт удалил их уже после этапа фильтрации данных
(замените данные в этом файле на свои).

И в конце можно указать как переименовать колонки перед сохранением в файле rename_columns.txt, чтобы скрипт переименовал указанные вами колонки перед сохранением.
В конце скрипт сохранит данные в файл out.xlsx.
