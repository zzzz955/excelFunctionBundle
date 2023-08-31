import os
import pandas as pd
import xlsxwriter
import openpyxl


def concat_singlesheet_excelfiles(file_paths):
    # 단일 시트 엑셀 파일 병합
    dfs = []
    if file_paths:
        for file in file_paths:
            df = pd.read_excel(file, na_filter=False)
            dfs.append(df)
        df = pd.concat(dfs, ignore_index=True)
        return df


def concat_multiplesheets_excelfiles(file_paths):
    # 다중 시트 엑셀 파일 병합
    df_list = []
    if file_paths:
        for file in file_paths:
            dfs = pd.read_excel(file, na_filter=False, sheet_name=None)
            for _, df in dfs.items():
                df_list.append(df)
        df = pd.concat(df_list, ignore_index=True)
        return df


def df_to_excel(df, save_path):
    # 데이터 프레임 엑셀화
    excel_writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
    df = df.applymap(
        lambda value: str(value) if isinstance(value, (int, float)) and len(str(value)) >= 10 else value)

    df.to_excel(excel_writer, sheet_name='mergeResult', index=False)
    excel_writer.close()


def table_to_excel(file_path, header, data):
    # 테이블 데이터 엑셀화
    df = pd.DataFrame(data, columns=header)
    df.to_excel(file_path, sheet_name='mergeResult', index=False)


def merge_to_excel_download(load_paths, save_path):
    # 다중 시트 엑셀 파일 생성 및 다운로드
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
                sheet_names[len(sheet_names) - 1] = excel.sheet_names[i] + f'..{n}'
    for num, df in enumerate(df_list):
        df.to_excel(excel_writer, sheet_name=sheet_names[num], index=False)
    excel_writer.close()


def return_sheets(file_path):
    # 엑셀 파일 내 시트 반환
    excel = pd.ExcelFile(file_path)
    sheet_names = excel.sheet_names
    return sheet_names


def file_change(file_path, sheet_name):
    # 파일의 특정 시트 데이터 프레임 반환
    df = pd.read_excel(file_path, sheet_name=sheet_name, na_filter=False)
    return df


def group_by_data(df, cmb1, cmb2, radio_btn1):
    # 다이얼 로그 값을 받아와 GROUP BY 기능 실행
    group = df.groupby(by=cmb1, as_index=False)[cmb2].agg(radio_btn1)
    group_sorted = group.sort_values(by=cmb2, ascending=False)
    return group_sorted


def duplicate_data(df, cmb1, radio_btn1):
    # 다이얼 로그 값을 받아와 중복 제거 기능 실행
    duplicated = df.drop_duplicates(subset=cmb1, keep=radio_btn1)
    return duplicated


def text_filter(df, criteria_col, criteria_text, checkbox_value):
    df = df.applymap(str)
    if checkbox_value:
        text_filtered_df = df[df[criteria_col] == criteria_text]
    else:
        text_filtered_df = df[df[criteria_col].str.contains(criteria_text)]
    return text_filtered_df

