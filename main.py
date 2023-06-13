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

# import openpyxl

# wb = openpyxl.load_workbook('data.xlsx')
# ws = wb.active
# bad_col = ["linkedid", "sequence", "peeraccount", "recordingfile", "dst_cnam", "outbound_cnam", "outbound_cnum", "did", "userfield",
#            "accountcode", "amaflags", "lastdata", "lastapp", "dstchannel", "channel", "clid", "src", "uniqueid"]
# #print('Total number of rows: '+str(ws.max_row)+'. And total number of columns: '+str(ws.max_column))
# dcontext_char = None
# dcontext_id = None
# duration_char = None
# billsec_char = None
# bad_context = ["ivr-1", "from-internal-xfer", "play-system-recording", "ivr-2", "ivr-3"]

# i = 1
# while True:
#     col_name = ws[str(chr(64 + i))+"1"].value
#     print(f"col_name: {col_name}; index: {i}; char: {str(chr(64 + i))}1")
#     if col_name == None:
#         break
#     if col_name in bad_col:
#         ws.delete_cols(i, 1)
#         print("deleted!")
#     else:
#         i += 1

# i = 1
# print("\n\n\n")
# for i in range(1, ws.max_column+1):
#     col_name = ws[str(chr(64 + i))+"1"].value
#     print(f"col_name: {col_name}; index: {i}")

# for i in range(1, ws.max_column+1):
#     col_name = ws[str(chr(64 + i))+"1"].value
#     if col_name == "dcontext":
#         dcontext_char = str(chr(64 + i))
#         dcontext_id = i
#         break

# for i in range(1, ws.max_column+1):
#     col_name = ws[str(chr(64 + i))+"1"].value
#     if col_name == "duration":
#         duration_char = str(chr(64 + i))
#         break

# for i in range(1, ws.max_column+1):
#     col_name = ws[str(chr(64 + i))+"1"].value
#     if col_name == "billsec":
#         billsec_char = str(chr(64 + i))
#         break

# for i in range(1, ws.max_column+1):
#     col_name = ws[str(chr(64 + i))+"1"].value
#     if col_name == "billsec":
#         billsec_char = str(chr(64 + i))
#         break

# i = 1
# while True:
#     cell_value = ws[dcontext_char+str(i)].value
#     print(f"cell_value: {cell_value}; char: {dcontext_char}{i}")
#     if cell_value == None:
#         break
#     if cell_value in bad_context:
#         ws.delete_rows(i, 1)
#         print("deleted!")
#     else:
#         i += 1


# print("\n\n\n")
# for i in range(2, ws.max_row+1):
#     cell_value = ws[duration_char+str(i)].value
#     if cell_value == None:
#         break
#     cell_value = f"{int(cell_value) // 60} min {int(cell_value) % 60} sec"
#     print(f"cell_value: {cell_value}; char: {duration_char}{i}")
#     ws[duration_char+str(i)].value = cell_value



# print("\n\n\n")
# for i in range(2, ws.max_row+1):
#     cell_value = ws[billsec_char+str(i)].value
#     if cell_value == None:
#         break
#     cell_value = f"{int(cell_value) // 60} min {int(cell_value) % 60} sec"
#     print(f"cell_value: {cell_value}; char: {billsec_char}{i}")
#     ws[billsec_char+str(i)].value = cell_value


# for i in range(1, ws.max_column+1):
#     ws.column_dimensions[str(chr(64 + i))].width = 30

# #ws.auto_filter.ref = f"{dcontext_char}1:{dcontext_char}{ws.max_row}"
# ws.delete_cols(dcontext_id, 1)
# wb.save("out.xlsx")
