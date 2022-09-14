import pandas as pd
import numpy as np


def delete_rows(excel_name, rows_to_delete):
    # as I only send the excel name, I need to open the file here again, and create it's writer to save after deleting the rows
    df = pd.read_excel(excel_name)
    writer_croped = pd.ExcelWriter(excel_name, engine='xlsxwriter')

    # here I order the array of rows to delete in descending order, and this is essencial,
    # because if I delete a row that is not at the end, all the other rows will go up (decrease the index)
    # and that would mess everything up, I need to make sure the rows are deleted from last to first
    rows_to_delete = np.sort(rows_to_delete)[::-1]

    # here I iterate through every row index and delete it
    # if you notice, I only go until len(rows_to_delete) - 1
    # that's because I initialized the array with [-1], so the last value, as it is in descending order, will be -1, and there is no row -1
    for i in range(0, len(rows_to_delete) - 1):
        df.drop(labels=rows_to_delete[i], axis=0, inplace=True)

    # finnally, saving the croped file without the rows I deleted using it's writer
    df.to_excel(writer_croped, index=False, header=True)
    writer_croped.save()