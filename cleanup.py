import pandas as pd
import xlsxwriter
import numpy as np
import sys



def main():
    excel_name = sys.argv[1]

    df = pd.read_excel(excel_name)

    # removing row 1 (Here I just save the original excel without the header, which is the first row):
    df.to_excel(excel_name.partition('.')[0] + "_croped.xlsx", index=False, header=False)
    croped_df = pd.read_excel(excel_name.partition('.')[0] + "_croped.xlsx")

    # removing column W:
    croped_df.drop(croped_df.columns[22], axis=1, inplace=True)
    croped_df.to_excel(excel_name.partition('.')[0] + "_croped.xlsx", index=False, header=True)

    writer_croped = pd.ExcelWriter(excel_name.partition('.')[0] + "_croped.xlsx", engine='xlsxwriter')
    croped_df.to_excel(writer_croped, index=False, header=True)
    writer_croped.save()

if __name__ == '__main__':
    main()