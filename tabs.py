import os
import pandas as pd
import dataframes
from PyQt5.QtWidgets import QVBoxLayout, QWidget, QPushButton, \
    QLabel, QHBoxLayout, QFileDialog, QMessageBox, QTableWidget, QTableWidgetItem, QComboBox

class Tab1(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()

        self.t1label1 = QLabel('rowCount : 0')
        self.t1label2 = QLabel('columnCount : 0')
        self.table_widget = QTableWidget()
        self.reserve_table_widget = QTableWidget()
        self.excel_download_btn = QPushButton('다운로드(업로드 파일 기준)')
        self.excel_download_btn2 = QPushButton('다운로드(테이블 데이터 기준)')

        self.layout.addLayout(layout2)
        layout2.addWidget(self.t1label1)
        layout2.addWidget(self.t1label2)
        self.layout.addWidget(self.table_widget)
        self.layout.addLayout(layout3)
        layout3.addWidget(self.excel_download_btn)
        layout3.addWidget(self.excel_download_btn2)
        self.setLayout(self.layout)

        self.excel_download_btn.clicked.connect(self.dataframe_to_excel)
        self.excel_download_btn2.clicked.connect(self.table_to_excel)
        self.table_widget.horizontalHeader().sortIndicatorChanged.connect(self.sortTable)

        self.file_paths = []
        self.df = None



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

    def dataframe_to_excel(self):
        if self.df is not None:
            save_file_path, _ = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')
            if not save_file_path:
                return
            dataframes.df_to_excel(self.df, save_file_path)
            self.open_filepath(save_file_path)

    def table_to_excel(self):
        try:
            # 테이블 데이터프레임화 및 엑셀 파일 저장
            if self.table_widget.rowCount() > 0:
                save_file_path, _ = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')
                if not save_file_path:
                    return
                header = []
                for column in range(self.table_widget.columnCount()):
                    header_item = self.table_widget.horizontalHeaderItem(column)
                    header.append(header_item.text())
                data = []
                for r in range(self.table_widget.rowCount()):
                    rowdata = []
                    for c in range(self.table_widget.columnCount()):
                        item = self.table_widget.item(r, c)
                        rowdata.append(item.text())
                    data.append(rowdata)
                dataframes.table_to_excel(save_file_path, header, data)
                self.open_filepath(save_file_path)
        except Exception as e:
            QMessageBox.warning(self, '경고', f'엑셀 파일을 저장할 수 없습니다. 해당 파일이 열려 있는 상태가 아닌지 확인해 보세요 {e}')

    def open_filepath(self, file_path):
        result = QMessageBox.question(self, '정보', '엑셀 파일 생성 완료, 생성한 파일을 여시겠습니까?',
                                      QMessageBox.Ok | QMessageBox.No,
                                      QMessageBox.Ok)
        if result == QMessageBox.Ok:
            os.startfile(file_path)
        else:
            return

    '''def merge_to_onefile(self):
        if self.reserve_table_widget.rowCount() > 0:
            save_file_path, _ = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')
            if not save_file_path:
                return
            dataframes.merged_to_excel_download(self.file_paths, save_file_path)
            self.open_filepath(save_file_path)'''
class Tab2(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()

        self.t1label1 = QLabel('rowCount : 0')
        self.t1label2 = QLabel('columnCount : 0')
        self.table_widget = QTableWidget()
        self.reserve_table_widget = QTableWidget()
        self.excel_download_btn = QPushButton('다운로드(업로드 파일 기준)')
        self.excel_download_btn2 = QPushButton('다운로드(테이블 데이터 기준)')

        self.layout.addLayout(layout2)
        layout2.addWidget(self.t1label1)
        layout2.addWidget(self.t1label2)
        self.layout.addWidget(self.table_widget)
        self.layout.addLayout(layout3)
        layout3.addWidget(self.excel_download_btn)
        layout3.addWidget(self.excel_download_btn2)
        self.setLayout(self.layout)

        self.excel_download_btn.clicked.connect(self.dataframe_to_excel)
        self.excel_download_btn2.clicked.connect(self.table_to_excel)
        self.table_widget.horizontalHeader().sortIndicatorChanged.connect(self.sortTable)

        self.file_paths = []
        self.df = None

    def sortTable(self, logicalIndex, order):
        # 각 헤더에 맞게 정렬
        if self.table_widget.rowCount() > 0:
            self.table_widget.sortItems(logicalIndex, order)

    def dataframe_to_excel(self):
        if self.df is not None:
            save_file_path, _ = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')
            if not save_file_path:
                return
            dataframes.df_to_excel(self.df, save_file_path)
            self.open_filepath(save_file_path)

    def table_to_excel(self):
        try:
            # 테이블 데이터프레임화 및 엑셀 파일 저장
            if self.table_widget.rowCount() > 0:
                save_file_path, _ = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')
                if not save_file_path:
                    return
                header = []
                for column in range(self.table_widget.columnCount()):
                    header_item = self.table_widget.horizontalHeaderItem(column)
                    header.append(header_item.text())
                data = []
                for r in range(self.table_widget.rowCount()):
                    rowdata = []
                    for c in range(self.table_widget.columnCount()):
                        item = self.table_widget.item(r, c)
                        rowdata.append(item.text())
                    data.append(rowdata)
                dataframes.table_to_excel(save_file_path, header, data)
                self.open_filepath(save_file_path)
        except Exception as e:
            QMessageBox.warning(self, '경고', f'엑셀 파일을 저장할 수 없습니다. 해당 파일이 열려 있는 상태가 아닌지 확인해 보세요 {e}')