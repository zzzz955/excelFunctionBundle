import os
import pandas as pd

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
    df_list = []
    sheet_names = []
    excel_writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
    for file in load_paths:
        df = pd.read_excel(file, na_filter=False)
        sheet_names.append(os.path.basename(file)[:10] + '...')
        df = df.applymap(
            lambda value: str(value) if isinstance(value, (int, float)) and len(str(value)) >= 10 else value)
        df_list.append(df)
        n = 0
        while len(sheet_names) != len(set(sheet_names)):
            increse = lambda x: x + 1
            n = increse(n)
            sheet_names[len(sheet_names) - 1] = os.path.basename(file)[:10] + f'...{n}'

    for num, df in enumerate(df_list):
        df.to_excel(excel_writer, sheet_name=sheet_names[num], index=False)
    excel_writer.close()