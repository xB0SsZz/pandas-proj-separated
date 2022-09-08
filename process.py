import pandas as pd
import xlsxwriter
import sys
import os
import croped_delete


def main():
    # reading excel name from command line arguments
    excel_name = sys.argv[1]

    # reading the excel
    croped_df = pd.read_excel(excel_name)

    # getting column names to add to the new files that will be created
    header = croped_df.columns.values.tolist()

    # creating an array to store all the row indexes I will be deleting from example_croped.xlsx
    rows_to_delete = [-1]

    # reading data_config_file.txt:
    with open('data_config_file.txt', 'r') as f:
        # iterating through all the lines in data_config_file.txt
        for line in f:

            # separating each argument of the line in different variables
            output_file = line.partition("|")[0]
            col = (line.partition("|")[2].partition("|")[0])
            arg1 = (line.partition("|")[2].partition("|")[2].partition("|")[0])
            arg2 = (line.partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[0])
            arg3 = (line.partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[2].partition("|")[0])

            # removing spaces from the strings
            output_file = output_file.replace(" ", "")
            col = col.replace(" ", "")
            arg1 = arg1.replace(" ", "")
            arg2 = arg2.replace(" ", "")
            arg3 = arg3.replace(" ", "")
            arg3 = arg3.replace("\n", "")

            print("---------------------------------------------------")
            print("Saving on file |" + output_file + "|")
            print("Searching on column |" + col + "| for |" + arg1 + "|" + arg2 + "|" + arg3 + "|")
            
            # checking if file already exists, if it does, the dataframe I will use will be from the already existing xlsx file
            # if not, create a new dataframe
            if os.path.exists(output_file):
                out_df = pd.read_excel(output_file)
            else:
                out_df = pd.DataFrame(columns=header)
                
            # creating the writer to save to the output file
            writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

            # in this block of code I check if there is a column with the name specified in the data_config_file.txt
            column = 0
            for i in range(0, len(croped_df.columns)):
                c = croped_df.columns[i]
                if c.replace(" ", "") == col:
                    print(f"Searching on column {col}")
                    column = i
            if column == 0:
                print("No column named " + col)
                return

            # here starts the iteration through every row of example_croped.xlsx 
            # I just explained what happens for one of the args, but it's the exact same procedure for the other two
            for j in range(0, len(croped_df)):
                if arg1 in str(croped_df.iloc[j][croped_df.columns[column]]):
                    # if arg1 is in row j and selected column, I store every value from this row in a list
                    row = list()
                    for k in range(0, len(croped_df.columns)):
                        row.append(croped_df.iloc[j][croped_df.columns[k]])
                    # and here I assign the value of the row to the end of the output file
                    print(f"Appending {row} to {output_file}.")
                    out_df.loc[len(out_df.index)] = row
                    # and finnally if this row isn't already in the array of rows to delete, I add it
                    # this row might be already in the array if chosen by other line of data_config_file.txt, so if taht is the case I don't add it to the array
                    if j not in rows_to_delete:
                        rows_to_delete.append(j)
                if arg2 in str(croped_df.iloc[j][croped_df.columns[column]]):
                    row = list()
                    for k in range(0, len(croped_df.columns)):
                        row.append(croped_df.iloc[j][croped_df.columns[k]])
                    print(f"Appending {row} to {output_file}.")
                    out_df.loc[len(out_df.index)] = row
                    if j not in rows_to_delete:
                        rows_to_delete.append(j)
                if arg3 in str(croped_df.iloc[j][croped_df.columns[column]]):
                    row = list()
                    for k in range(0, len(croped_df.columns)):
                        row.append(croped_df.iloc[j][croped_df.columns[k]])
                    print(f"Appending {row} to {output_file}.")
                    out_df.loc[len(out_df.index)] = row
                    if j not in rows_to_delete:
                        rows_to_delete.append(j)
            # when the output dataframe is complete, I save it using the writer I created before
            print(f"{output_file} has {str(len(out_df))} rows")
            out_df.to_excel(writer, index=False, header=True)
            writer.save()
    

    # delete the rows (goes to croped_delete.py file)
    croped_delete.delete_rows(excel_name, rows_to_delete)
            




if __name__ == '__main__':
    main()