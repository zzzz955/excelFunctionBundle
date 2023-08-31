import os
import pandas as pd
import dataframes
from PyQt5.QtWidgets import QVBoxLayout, QWidget, QPushButton, \
    QLabel, QHBoxLayout, QTableWidget, QComboBox, QMessageBox

class Tab1(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window

        # 위젯 추가
        self.layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        self.label1 = QLabel('rowCount : 0')
        self.label2 = QLabel('columnCount : 0')
        self.table_widget = QTableWidget()
        self.excel_download_btn = QPushButton('다운로드(업로드 파일 기준)')
        self.excel_download_btn2 = QPushButton('다운로드(테이블 데이터 기준)')

        # 레이아웃 지정
        self.layout.addLayout(layout2)
        layout2.addWidget(self.label1)
        layout2.addWidget(self.label2)
        self.layout.addWidget(self.table_widget)
        self.layout.addLayout(layout3)
        layout3.addWidget(self.excel_download_btn)
        layout3.addWidget(self.excel_download_btn2)
        self.setLayout(self.layout)

        # 시그널 추가
        self.excel_download_btn.clicked.connect(self.main_window.dataframe_to_excel)
        self.excel_download_btn2.clicked.connect(self.main_window.table_to_excel)
        self.table_widget.horizontalHeader().sortIndicatorChanged.connect(self.main_window.sort_table)

class Tab2(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window

        # 위젯 추가
        self.layout = QVBoxLayout()
        layout1 = QHBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        self.t2label1 = QLabel('<b>파일 정보 : </b>')
        self.file_list_btn = QPushButton('파일 리스트')
        self.t2label2 = QLabel('<b>시트 정보 : </b>')
        self.combobox1 = QComboBox()
        self.label1 = QLabel('rowCount : 0')
        self.label2 = QLabel('columnCount : 0')
        self.table_widget = QTableWidget()
        self.excel_download_btn = QPushButton('다운로드(업로드 파일 기준)')
        self.excel_download_btn2 = QPushButton('다운로드(테이블 데이터 기준)')

        # 레이아웃 지정
        self.layout.addLayout(layout1)
        layout1.addWidget(self.t2label1)
        layout1.addWidget(self.file_list_btn)
        layout1.addWidget(self.t2label2)
        layout1.addWidget(self.combobox1)

        self.layout.addLayout(layout2)
        layout2.addWidget(self.label1)
        layout2.addWidget(self.label2)
        self.layout.addWidget(self.table_widget)

        self.layout.addLayout(layout3)
        layout3.addWidget(self.excel_download_btn)
        layout3.addWidget(self.excel_download_btn2)
        self.setLayout(self.layout)

        # 시그널 추가
        self.file_list_btn.clicked.connect(self.main_window.show_listwidget_exec)
        self.excel_download_btn.clicked.connect(self.main_window.dataframe_to_excel)
        self.excel_download_btn2.clicked.connect(self.main_window.table_to_excel)
        self.table_widget.horizontalHeader().sortIndicatorChanged.connect(self.main_window.sort_table)

        # 변수 초기화
        self.file_path = None

    def combobox1_add_items(self, file_path):
        # 콤보 박스 값 추가 및 시그널 연결
        if self.combobox1.count() > 0:
            self.combobox1.currentTextChanged.disconnect(self.combobox_value_changed)
            self.combobox1.clear()
        self.file_path = file_path
        sheet_names = dataframes.return_sheets(file_path)
        self.combobox1.addItems(sheet_names)
        current_sheet_name = self.combobox1.currentText()
        self.combobox1.currentTextChanged.connect(self.combobox_value_changed)
        return file_path, current_sheet_name

    def combobox_value_changed(self):
        # 콤보 박스 값 변경 시 테이블 최신화
        try:
            current_sheet_name = self.combobox1.currentText()
            df = dataframes.file_change(self.file_path, current_sheet_name)
            self.main_window.tab2_df = df
            self.main_window.df_to_table(df)
        except Exception as e:
            QMessageBox.critical(self, '예외 발생', f'엑셀 파일을 테이블에 불러 올 수 없습니다. {e}'
                                                f'\n1. 파일이 올바른 형식 인지 확인해 주세요.'
                                                f'\n2. 파일 내부 데이터에 문제가 있을 수 있습니다.'
                                                f'\n → [*.*], (*.*), /| 등의 특수 문자를 제거 후 시도해 주세요.')
