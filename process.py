import pandas as pd
import numpy as np
import xlsxwriter
import sys
import os


def main():
    excel_name = sys.argv[1]

    croped_df = pd.read_excel(excel_name)

    print(croped_df)

    header = croped_df.columns.values.tolist()

    

    rows_to_delete = [-1]
    # reading data_config_file.txt:
    with open('data_config_file.txt', 'r') as f:
        for line in f:
            output_file = line.partition("|")[0]
            col = (line.partition("|")[2].partition("|")[0])
            arg1 = (line.partition("|")[2].partition("|")[2].partition("|")[0])
            arg2 = (line.partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[0])
            arg3 = (line.partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[0])
            output_file = output_file.replace(" ", "")
            
            col = col.replace(" ", "")
            arg1 = arg1.replace(" ", "")
            arg2 = arg2.replace(" ", "")
            arg3 = arg3.replace(" ", "")
            arg3 = arg3.replace("\n", "")

            print("---------------------------------------------------")
            print("Saving on file |" + output_file + "|")
            print("Searching on column |" + col + "| for |" + arg1 + "|" + arg2 + "|" + arg3 + "|")
            
            # checking if file already exists
            if os.path.exists(output_file):
                out_df = pd.read_excel(output_file)
            else:
                out_df = pd.DataFrame(columns=header)
                
            
            writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

            column = 0

            for i in range(0, len(croped_df.columns)):
                c = croped_df.columns[i]
                if c.replace(" ", "") == col:
                    column = i
            if column == 0:
                print("No column named " + col)
                return
            for j in range(0, len(croped_df)):
                if arg1 in str(croped_df.iloc[j][croped_df.columns[column]]):
                    row = list()
                    for k in range(0, len(croped_df.columns)):
                        row.append(croped_df.iloc[j][croped_df.columns[k]])
                    print("INDEX: " + str(len(out_df.index)))
                    out_df.loc[len(out_df.index)] = row
                    if j not in rows_to_delete:
                        rows_to_delete.append(j)
                if arg2 in str(croped_df.iloc[j][croped_df.columns[column]]):
                    row = list()
                    for k in range(0, len(croped_df.columns)):
                        row.append(croped_df.iloc[j][croped_df.columns[k]])
                    out_df.loc[len(out_df.index)] = row
                    if j not in rows_to_delete:
                        rows_to_delete.append(j)
                if arg3 in str(croped_df.iloc[j][croped_df.columns[column]]):
                    row = list()
                    for k in range(0, len(croped_df.columns)):
                        row.append(croped_df.iloc[j][croped_df.columns[k]])
                    out_df.loc[len(out_df.index)] = row
                    if j not in rows_to_delete:
                        rows_to_delete.append(j)
            
            out_df.to_excel(writer, index=False, header=True)
            writer.save()
            




if __name__ == '__main__':
    main()