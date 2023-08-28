import os
import pandas as pd
import openpyxl


def concat_singlesheet_excelfiles(file_paths):
    dfs = []
    if file_paths:
        for file in file_paths:
            df = pd.read_excel(file, na_filter=False)
            dfs.append(df)
        df = pd.concat(dfs, ignore_index=True)
        return df


def concat_multiplesheets_excelfiles(file_paths):
    df_list = []
    if file_paths:
        for file in file_paths:
            dfs = pd.read_excel(file, na_filter=False, sheet_name=None)
            for _, df in dfs.items():
                df_list.append(df)
        df = pd.concat(df_list, ignore_index=True)
        return df


def multiplesheets_excelfiles(file_paths):
    df_list = []
    if file_paths:
        for file in file_paths:
            dfs = pd.read_excel(file, na_filter=False, sheet_name=None)
            for _, df in dfs.items():
                df_list.append(df)
        df = pd.concat(df_list, ignore_index=True)
        return df


def df_to_excel(df, save_path):
    excel_writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
    df = df.applymap(
        lambda value: str(value) if isinstance(value, (int, float)) and len(str(value)) >= 10 else value)

    df.to_excel(excel_writer, sheet_name='mergeResult', index=False)
    excel_writer.close()


def table_to_excel(file_path, header, data):
    df = pd.DataFrame(data, columns=header)
    df.to_excel(file_path, sheet_name='mergeResult', index=False)


def merge_to_excel_download(load_paths, save_path):
    try:
        df_list = []
        sheet_names = []
        excel_writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
        for file in load_paths:
            excel = pd.ExcelFile(file)
            sheet_count = len(excel.sheet_names)
            for i in range(sheet_count):
                df = pd.read_excel(file, sheet_name=i, na_filter=False, engine='openpyxl')
                sheet_names.append(excel.sheet_names[i])
                df = df.applymap(
                    lambda value: str(value) if isinstance(value, (int, float)) and len(str(value)) >= 10 else value)
                df_list.append(df)
                n = 0
                while len(sheet_names) != len(set(sheet_names)):
                    increse = lambda x: x + 1
                    n = increse(n)
                    sheet_names[len(sheet_names) - 1] = excel.sheet_names[i] + f'...{n}'
        for num, df in enumerate(df_list):
            df.to_excel(excel_writer, sheet_name=sheet_names[num], index=False)
        excel_writer.close()
    except Exception as e:
        print(e)


def return_sheets(file_path):
    excel = pd.ExcelFile(file_path)
    sheet_names = excel.sheet_names
    return sheet_names


def file_change(file_path, sheet_name):
    df = pd.read_excel(file_path, sheet_name=sheet_name, na_filter=False)
    return df


def save_whole_sheet_to_one_file(file_paths, sheet_index):
    dfs = []
    excel_writer = pd.ExcelWriter()