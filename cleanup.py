import pandas as pd
import xlsxwriter
import numpy as np
import sys
import os



def main():
    # reading excel name from command line arguments
    excel_name = sys.argv[1]

    # checking if the file exists in current directory, if not, program terminates
    if not os.path.exists(excel_name):
        sys.exit(f'The excel file {excel_name} doesnt exist')

    # reading excel through pandas, skipping one row, so the first part of the cleanup is already done
    df = (pd.read_excel(excel_name, skiprows=1).dropna(how='all', axis=1))

    # create the writer to write to excel_name_croped.xlsx
    writer_croped = pd.ExcelWriter(excel_name.partition('.')[0] + "_croped.xlsx", engine='xlsxwriter')


    # removing column W and saving it with the writer:
    df.drop(df.columns[22], axis=1, inplace=True)

    df.to_excel(writer_croped, index=False, header=True)
    writer_croped.save()

if __name__ == '__main__':
    main()