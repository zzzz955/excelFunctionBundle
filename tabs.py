import os
import pandas as pd
from PyQt5.QtWidgets import QVBoxLayout, QWidget, QPushButton, \
    QLabel, QHBoxLayout, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QComboBox

class Tab1(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        self.layout.addLayout(layout2)

        self.t1label1 = QLabel('rowCount : 0')
        layout2.addWidget(self.t1label1)
        self.t1label2 = QLabel('columnCount : 0')
        layout2.addWidget(self.t1label2)

        self.table_widget = QTableWidget()
        self.reserve_table_widget = QTableWidget()
        self.layout.addWidget(self.table_widget)

        layout3 = QHBoxLayout()
        self.layout.addLayout(layout3)

        self.excel_download_btn = QPushButton()
        self.excel_download_btn.setText('다운로드(단일 시트 내 데이터 병합)')
        layout3.addWidget(self.excel_download_btn)
        self.excel_download_btn.clicked.connect(self.table_to_excel)

        self.excel_download_btn2 = QPushButton()
        self.excel_download_btn2.setText('다운로드(파일별 시트 생성)')
        layout3.addWidget(self.excel_download_btn2)
        self.excel_download_btn2.clicked.connect(self.merge_to_onefile)

        # 테이블 정렬
        self.table_widget.horizontalHeader().sortIndicatorChanged.connect(self.sortTable)
        self.setLayout(self.layout)

        self.file_paths = []

    def single_sheet_excel_file_Conversion(self, file_paths):
        # 엑셀 업로드 함수
        try:
            self.file_paths = file_paths
            dfs = []
            if file_paths:
                for file in file_paths:
                    df = pd.read_excel(file, na_filter=False)
                    dfs.append(df)
                df = pd.concat(dfs, ignore_index=True)
                self.df = df
                self.df_to_table(df)
                self.df_to_reserve_table(df)
        except Exception as e:
            print(e)

    def multiple_sheet_excel_file_Conversion(self, file_paths):
        # 엑셀 업로드 함수
        try:
            self.file_paths = file_paths
            df_list = []
            if file_paths:
                for file in file_paths:
                    dfs = pd.read_excel(file, na_filter=False, sheet_name=None)
                    for _, df in dfs.items():
                        df_list.append(df)
                df = pd.concat(df_list, ignore_index=True)
                self.df = df
                self.df_to_table(df)
                self.df_to_reserve_table(df)
        except Exception as e:
            print(e)

    def df_to_table(self, df):
        # 데이터 프레임 테이블화
        self.table_widget.setRowCount(len(df))
        self.table_widget.setColumnCount(len(df.columns))
        self.table_widget.setHorizontalHeaderLabels(df.columns)
        self.table_widget.horizontalHeader().setSortIndicatorShown(True)
        for r in range(len(df)):
            for c in range(len(df.columns)):
                item = str(df.iloc[r, c])
                self.table_widget.setItem(r, c, QTableWidgetItem(item))
        self.table_widget.resizeColumnsToContents()
        # 행 및 열 개수 노출
        self.t1label1.setText(f'rowCount : {str(self.table_widget.rowCount())}')
        self.t1label2.setText(f'columnCount : {str(self.table_widget.columnCount())}')

    def df_to_reserve_table(self, df):
        # 기존 병합 데이터프레임 값 임시 테이블에 저장
        self.reserve_table_widget.setRowCount(len(df))
        self.reserve_table_widget.setColumnCount(len(df.columns))
        self.reserve_table_widget.setHorizontalHeaderLabels(df.columns)
        for r in range(len(df)):
            for c in range(len(df.columns)):
                item = str(df.iloc[r, c])
                self.reserve_table_widget.setItem(r, c, QTableWidgetItem(item))
        self.reserve_table_widget.resizeColumnsToContents()

    def do_group_by(self, cmb1, cmb2, radio_btn1):
        # 다이얼 로그 값을 받아와 GROUP BY 기능 실행
        df = self.df
        group = df.groupby(by=cmb1, as_index=False)[cmb2].agg(radio_btn1)
        group_sorted = group.sort_values(by=cmb2, ascending=False)
        self.df_to_table(group_sorted)

    def exit_group_by(self):
        # 집계 테이블 원 상태로 복구
        if self.reserve_table_widget.rowCount() > 0:
            self.df_to_table(self.df)

    def sortTable(self, logicalIndex, order):
        # 각 헤더에 맞게 정렬
        if self.table_widget.rowCount() > 0:
            self.table_widget.sortItems(logicalIndex, order)

    def excel_download(self, df):
        # 병합된 엑셀 파일 다운로드 경로 설정
        file_path, _ = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')
        if not file_path:
            return
        try:
            df.to_excel(file_path, sheet_name='mergeResult', index=False)
            result = QMessageBox.question(self, '정보', '엑셀 파일 생성 완료, 생성한 파일을 여시겠습니까?', QMessageBox.Ok | QMessageBox.No,
                                          QMessageBox.No)
            if result == QMessageBox.Ok:
                os.startfile(file_path)
            else:
                return
        except Exception as e:
            QMessageBox.warning(self, '오류', f'파일 저장 중 예외 발생, 예외 내용 : {e}')

    def table_to_excel(self):
        # 테이블 데이터프레임화 및 엑셀 파일 저장
        try:
            if self.reserve_table_widget.rowCount() > 0:
                header_col = []
                for column in range(self.reserve_table_widget.columnCount()):
                    header_item = self.reserve_table_widget.horizontalHeaderItem(column)
                    header_col.append(header_item.text())

                data = []
                for r in range(self.reserve_table_widget.rowCount()):
                    rowdata = []
                    for c in range(self.reserve_table_widget.columnCount()):
                        item = self.reserve_table_widget.item(r, c)
                        rowdata.append(item.text())
                    data.append(rowdata)
                df = pd.DataFrame(data, columns=header_col)
                self.excel_download(df)
        except Exception as e:
            QMessageBox.warning(self, '경고', f'엑셀 파일 변환 중 예외 발생, 예외 내용 : {e}')

    def merge_to_onefile(self):
        try:
            save_file_path, _ = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')

            if self.reserve_table_widget.rowCount() > 0 and save_file_path:
                df_list = []
                sheet_names = []
                excel_writer = pd.ExcelWriter(save_file_path, engine='xlsxwriter')
                for file in self.file_paths:
                    df = pd.read_excel(file, na_filter=False)
                    sheet_names.append(os.path.basename(file)[:10] + '...')
                    df = df.applymap(lambda value: str(value) if isinstance(value, (int, float)) and len(str(value)) >= 10 else value)
                    df_list.append(df)
                    n = 0
                    while len(sheet_names) != len(set(sheet_names)):
                        increse = lambda x: x+1
                        n = increse(n)
                        sheet_names[len(sheet_names)-1] = os.path.basename(file)[:10] + f'...{n}'

                for num, df in enumerate(df_list):
                    df.to_excel(excel_writer, sheet_name=sheet_names[num], index=False)
                excel_writer.close()
                result = QMessageBox.question(self, '정보', '엑셀 파일 생성 완료, 생성한 파일을 여시겠습니까?',
                                              QMessageBox.Ok | QMessageBox.No,
                                              QMessageBox.Ok)
                if result == QMessageBox.Ok:
                    os.startfile(save_file_path)
                else:
                    return
        except Exception as e:
            QMessageBox.warning(self, '경고', f'엑셀 파일 변환 중 예외 발생, 예외 내용 : {e}')

class Tab2(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.layout = QVBoxLayout()
        self.table_widget = QTableWidget()
        self.combo_box = QComboBox()