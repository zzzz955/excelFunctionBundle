import os
import pandas as pd
from openpyxl import Workbook, load_workbook
from PyQt5.QtWidgets import QVBoxLayout, QWidget, QPushButton, \
    QLabel, QHBoxLayout, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QListWidget, QAbstractItemView

class Tab1(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.layout = QVBoxLayout(self)

        layout2 = QHBoxLayout(self)
        self.layout.addLayout(layout2)

        self.t1label1 = QLabel('rowCount : 0')
        layout2.addWidget(self.t1label1)
        self.t1label2 = QLabel('columnCount : 0')
        layout2.addWidget(self.t1label2)

        self.table_widget = QTableWidget()
        self.reserve_table_widget = QTableWidget()
        self.layout.addWidget(self.table_widget)

        layout3 = QHBoxLayout(self)
        self.layout.addLayout(layout3)

        self.do_group_by_btn = QPushButton()
        self.do_group_by_btn.setText('집계 함수 실행')
        layout3.addWidget(self.do_group_by_btn)
        self.do_group_by_btn.clicked.connect(self.connect_dialog)

        self.exit_group_by_btn = QPushButton()
        self.exit_group_by_btn.setText('집계 함수 해제')
        layout3.addWidget(self.exit_group_by_btn)
        self.exit_group_by_btn.clicked.connect(self.exit_group_by)

        layout4 = QHBoxLayout(self)
        self.layout.addLayout(layout4)

        self.excel_upload_btn = QPushButton()
        self.excel_upload_btn.setText('엑셀 파일 업로드')
        layout4.addWidget(self.excel_upload_btn)
        self.excel_upload_btn.clicked.connect(self.excel_upload)

        self.excel_download_btn = QPushButton()
        self.excel_download_btn.setText('엑셀 파일 다운로드')
        layout4.addWidget(self.excel_download_btn)
        self.excel_download_btn.clicked.connect(self.table_to_excel)
        self.table_widget.horizontalHeader().sortIndicatorChanged.connect(self.sortTable)

    def excel_upload(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, '파일 선택', '', 'Excel Files(*.xlsx)')
        if file_paths:
            df = pd.read_excel(file_paths[0], na_filter=False)
            try:
                for file in file_paths[1:]:
                    df2 = pd.read_excel(file, na_filter=False)
                    df = pd.concat([df, df2], ignore_index=True)
                self.df = df
                self.df_to_table(df)
                self.df_to_reserve_table(df)
            except Exception as e:
                print(e)

    def df_to_table(self, df):
        self.table_widget.setRowCount(len(df))
        self.table_widget.setColumnCount(len(df.columns))
        self.table_widget.setHorizontalHeaderLabels(df.columns)
        self.table_widget.horizontalHeader().setSortIndicatorShown(True)
        for r in range(len(df)):
            for c in range(len(df.columns)):
                item = str(df.iloc[r, c])
                self.table_widget.setItem(r, c, QTableWidgetItem(item))
        self.table_widget.resizeColumnsToContents()
        self.t1label1.setText(f'rowCount : {str(self.table_widget.rowCount())}')
        self.t1label2.setText(f'columnCount : {str(self.table_widget.columnCount())}')

    def excel_download(self, df):
        file_path = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')
        if not file_path[0]:
            return
        try:
            df.to_excel(file_path[0], sheet_name='mergeResult', index=False)
            result = QMessageBox.question(self, '정보', '엑셀 파일 생성 완료, 생성한 파일을 여시겠습니까?', QMessageBox.Ok | QMessageBox.No,
                                          QMessageBox.No)
            if result == QMessageBox.Ok:
                os.startfile(file_path[0])
            else:
                return
        except Exception as e:
            QMessageBox.warning(self, '오류', f'파일 저장 중 예외 발생, 예외 내용 : {e}')

    def table_to_excel(self):
        try:
            if self.reserve_table_widget.rowCount()>0:
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

    def df_to_reserve_table(self, df):
        self.reserve_table_widget.setRowCount(len(df))
        self.reserve_table_widget.setColumnCount(len(df.columns))
        self.reserve_table_widget.setHorizontalHeaderLabels(df.columns)
        for r in range(len(df)):
            for c in range(len(df.columns)):
                item = str(df.iloc[r, c])
                self.reserve_table_widget.setItem(r, c, QTableWidgetItem(item))
        self.reserve_table_widget.resizeColumnsToContents()

    def connect_dialog(self):
        try:
            if self.reserve_table_widget.rowCount() > 0:
                self.main_window.open_tab1_dialog()
        except Exception as e:
            print(e)

    def do_group_by(self, cmb1, cmb2, radio_btn1):
        df = self.df
        group = df.groupby(by=cmb1, as_index=False)[cmb2].agg(radio_btn1)
        group_sorted = group.sort_values(by=cmb2, ascending=False)
        self.df_to_table(group_sorted)

    def exit_group_by(self):
        if self.reserve_table_widget.rowCount() > 0:
            self.df_to_table(self.df)

    def sortTable(self, logicalIndex, order):
        if self.table_widget.rowCount() > 0:
            self.table_widget.sortItems(logicalIndex, order)

'''class Tab2(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.layout = QVBoxLayout()'''