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
        rows_added = df_new_sheet.index

        # get the list of rows that are in both dataframes
        rows_common = list(set(rows_old).intersection(set(rows_added)))
        # get the list of rows that are only in the old dataframe
        rows_deleted = list(set(rows_old) - set(rows_common))
        # get the list of rows that are only in the new dataframe
        rows_added = list(set(rows_added) - set(rows_common))
        # create a dataframe for each type of difference
        dict_diff[sheet_name]['added'] = df_new_sheet.loc[rows_added]
        dict_diff[sheet_name]['deleted'] = df_old_sheet.loc[rows_deleted]

        # modified rows
        df_modified = pd.DataFrame()
        for row in rows_common:
            if not df_old_sheet.loc[row].astype(object).equals(df_new_sheet.loc[row].astype(object)):
                old_row = df_old_sheet.loc[row].to_frame(name='Old').transpose()
                new_row = df_new_sheet.loc[row].to_frame(name='New').transpose()
                df_modified = pd.concat([old_row, new_row], ignore_index=True)
        dict_diff[sheet_name]["modified"] = df_modified
    return dict_diff

def create_excel_file(dict_diff):
    """
    Creates an Excel file listing the differences between two dataframes
    param dict_diff: a dictionary of dataframes containing the differences
    """

    # Create a new Excel file to store the results
    result_file = xl.Workbook()
    result_sheet = result_file.active

    # Add headers to the result sheet
    first_key = list(dict_diff.keys())[0]
    second_key = list(dict_diff[first_key].keys())[0]
    col_names = list(dict_diff[first_key][second_key].columns)
    result_sheet.append(["File Status", "Sheet Name"] + col_names)

    is_old = True
    
    # Iterate through differences and append data to the result sheet
    for sheet_name, diff_data in dict_diff.items():
        for diff_type, diff_df in diff_data.items():
            if diff_type == 'modified':
                for index, row in diff_df.iterrows():
                    if is_old:
                        result_sheet.append(["Old", sheet_name] + row.tolist())
                        is_old = False
                    else:
                        result_sheet.append(["New", sheet_name] + row.tolist())
                        is_old = True 
            elif diff_type == 'added':
                for index, row in diff_df.iterrows():
                    result_sheet.append(["Added", sheet_name] + row.tolist())
            elif diff_type == 'deleted':
                for index, row in diff_df.iterrows():
                    result_sheet.append(["Deleted", sheet_name] + row.tolist())

    # Save the result file
    result_file.save("result.xlsx")

if __name__ == "__main__":
    
    # load the two excel files
    df_old = pd.read_excel("data1.xlsx", sheet_name=None)
    df_new = pd.read_excel("data2.xlsx", sheet_name=None)
    
    # compare the two dataframes
    dict_diff = compare_dataframes(df_old, df_new)

    # create excel file containing the diff
    create_excel_file(dict_diff)
    