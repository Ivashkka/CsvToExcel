import asterisk_excel_lib
BCNAME = "./bad_columns.txt"
LBCNAME = "./last_bad_columns.txt"
RENAME = "./rename_columns.txt"
data = asterisk_excel_lib.data()

def main():
    drop_garbage(BCNAME)
    keep_rows()
    drop_rows()
    drop_garbage(LBCNAME)
    rename_columns_on_save()
    data.save()

def drop_garbage(path): # instantly drop usless columns
    print("deliting usless columns ...")
    bf_f = open(path, 'r')
    bad_columns = bf_f.readlines()
    bf_f.close()
    for i in range(0, len(bad_columns)):
        bad_columns[i] = bad_columns[i].strip()
        data.del_column(bad_columns[i])
    print(f"deleted {len(bad_columns)} columns")

def rename_columns_on_save(): # sets new names of columns when file saves
    print("renaming columns ...")
    rf_f = open(RENAME, 'r')
    rename_list = rf_f.readlines()
    rf_f.close()
    for i in range(0, len(rename_list)):
        rename_list[i] = rename_list[i].strip()
        rename_list[i] = rename_list[i].split("=")
        data.rename_column(rename_list[i][0], rename_list[i][1])
    print(f"renamed {len(rename_list)} columns")

def drop_rows():
    print("searching bad data in rows ...")
    for i in data.get_columns():
        if asterisk_excel_lib.fexist(f"{i}.txt"):
            print(f"cleaning bad data from {i} ...")
            bd_f = open(f"{i}.txt", 'r')
            bad_data = bd_f.readlines()
            bd_f.close()
            for j in range(0, len(bad_data)):
                bad_data[j] = bad_data[j].strip()
            dcount = data.drop_rows_by_value(i, bad_data)
            print(f"deleted {dcount} rows")

def keep_rows():
    print("deliting all rows except good once ...")
    for i in data.get_columns():
        if asterisk_excel_lib.fexist(f"!{i}.txt"):
            print(f"storing anly good data at {i} ...")
            gd_f = open(f"!{i}.txt", 'r')
            good_data = gd_f.readlines()
            gd_f.close()
            for j in range(0, len(good_data)):
                good_data[j] = good_data[j].strip()
            dcount = data.keep_rows_by_value(i, good_data)
            print(f"deleted {dcount} rows")



if __name__ == "__main__":
    main()
