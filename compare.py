import pandas as pd
import openpyxl as xl

def compare_dataframes(df_old, df_new):
    """
    Compares two dataframes and return a dictionary of dataframes containing the differences
    param df_old:   old dataframe
    param df_new:   new dataframe
    return:         a dictionary of dataframes containing the differences
    """

    dict_diff = {}
    for sheet_name in df_old.keys():
        dict_diff[sheet_name] = {}
        df_old_sheet = df_old[sheet_name]
        df_new_sheet = df_new[sheet_name]
        # get the list of rows in both dataframes
        rows_old = df_old_sheet.index
        rows_new = df_new_sheet.index

        # get the list of rows that are in both dataframes
        rows_common = list(set(rows_old).intersection(set(rows_new)))
        # get the list of rows that are only in the old dataframe
        rows_deleted = list(set(rows_old) - set(rows_common))
        # get the list of rows that are only in the new dataframe
        rows_new = list(set(rows_new) - set(rows_common))
        # create a dataframe for each type of difference
        dict_diff[sheet_name]['new'] = df_new_sheet.loc[rows_new]
        dict_diff[sheet_name]['deleted'] = df_old_sheet.loc[rows_deleted]

        # modified rows
        df_modified = pd.DataFrame()
        for row in rows_common:
            if not df_old_sheet.loc[row].astype(object).equals(df_new_sheet.loc[row].astype(object)):
                old_row = df_old_sheet.loc[row].to_frame(name='Old Row').transpose()
                new_row = df_new_sheet.loc[row].to_frame(name='New Row').transpose()
                df_modified = pd.concat([old_row, new_row], ignore_index=True)
        dict_diff[sheet_name]["modified"] = df_modified
    return dict_diff

def create_excel_file(dict_diff):
    """
    Creates an excel file containing the differences between two dataframes
    param dict_diff:  a dictionary of dataframes containing the differences
    """

    # Create a new Excel file to store the results
    result_file = xl.Workbook()
    result_sheet = result_file.active

    # Add headers to the result sheet
    first_key = list(dict_diff.keys())[0]
    second_key = list(dict_diff[first_key].keys())[0]
    col_names = list(dict_diff[list(dict_diff.keys())[0]][list(dict_diff[list(dict_diff.keys())[0]].keys())[0]].columns)
    result_sheet.append(["File Status", "Sheet Name"] + col_names)

if __name__ == "__main__":
    # load the two excel files
    df_old = pd.read_excel("data1.xlsx", sheet_name=None)
    df_new = pd.read_excel("data2.xlsx", sheet_name=None)
    # compare the two dataframes
    dict_diff = compare_dataframes(df_old, df_new)
    # create excel file containing the diff
    create_excel_file(dict_diff)
    