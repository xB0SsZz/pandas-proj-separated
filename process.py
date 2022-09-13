import pandas as pd
import xlsxwriter
import sys
import os
import croped_delete


def append_records(read_df, header, output_file, col, arg1, arg2, arg3, rows_to_delete):

    args = list()
    if arg1 != "NULL":
        args.append(arg1)
    if arg2 != "NULL":
        args.append(arg2)
    if arg2 != "NULL":
        args.append(arg2)

    out_df = pd.DataFrame(columns=header)
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')
    # in this block of code I check if there is a column with the name specified in the data_config_file.txt
    column = 0
    for i in range(0, len(read_df.columns)):
        c = read_df.columns[i]
        if c.replace(" ", "") == col:
            print(f"Searching on column {col}")
            column = i
    if column == 0:
        print("No column named " + col)
        return

    # here starts the iteration through every row of example_croped.xlsx
    # I just explained what happens for one of the args, but it's the exact same procedure for the other two
    for j in range(0, len(read_df)):
        flag = True
        for a in range(0, len(args)):
            if str(args[a]).lower() in str(read_df.iloc[j][read_df.columns[column]]).lower():
                flag = True
            else:
                flag = False
        if flag:
            # if arg1 is in row j and selected column, I store every value from this row in a list
            row = {}
            for k in range(0, len(read_df.columns)):
                row.update({f"{read_df.columns[k]}":read_df.iloc[j][read_df.columns[k]]})
            update_df = pd.DataFrame(row, index=[0])
            out_df = pd.concat([out_df, update_df], ignore_index=True)
            if j not in rows_to_delete:
                rows_to_delete.append(j)
        
    # when the output dataframe is complete, I save it using the writer I created before
    out_df.to_excel(writer, index=False, header=True)
    writer.save()
    return rows_to_delete

def main():
    # reading excel name from command line arguments
    excel_name = sys.argv[1]

    # reading the excel
    croped_df = pd.read_excel(excel_name)

    # getting column names to add to the new files that will be created
    header = croped_df.columns.values.tolist()

    # creating an array to store all the row indexes I will be deleting from example_croped.xlsx
    rows_to_delete = [-1]

    filenames = list()

    # reading data_config_file.txt:
    with open('data_config_file.txt', 'r') as f:
        # iterating through all the lines in data_config_file.txt
        for line in f:
            if not line.startswith("#"):
                # separating each argument of the line in different variables
                output_file = line.partition("|")[0]
                col = (line.partition("|")[2].partition("|")[0])
                arg1 = (line.partition("|")[2].partition("|")[2].partition("|")[0])
                arg2 = (line.partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[0])
                arg3 = (line.partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[0])

                # removing spaces from the end and the start of the strings
                output_file = output_file.strip()
                col = col.strip()
                arg1 = arg1.strip()
                arg2 = arg2.strip()
                arg3 = arg3.strip()
                arg3 = arg3.replace("\n", "")

                print("---------------------------------------------------")
                print("Saving on file |" + output_file + "|")
                print("Searching on column |" + col + "| for |" + arg1 + "|" + arg2 + "|" + arg3 + "|")

                # checking if file already exists, if it does, the dataframe I will use will be from the already existing xlsx file
                # if not, create a new dataframe

                if output_file in filenames:
                    read_df = pd.read_excel(output_file)
                    rows_to_delete = append_records(read_df, header, output_file, col, arg1, arg2, arg3, rows_to_delete)
                else:
                    filenames.append(output_file)
                    rows_to_delete = append_records(croped_df, header, output_file, col, arg1, arg2, arg3, rows_to_delete)


    # delete the rows (goes to croped_delete.py file)
    croped_delete.delete_rows(excel_name, rows_to_delete)





if __name__ == '__main__':
    main()