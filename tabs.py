import os
import pandas as pd
import dataframes
from PyQt5.QtWidgets import QVBoxLayout, QWidget, QPushButton, \
    QLabel, QHBoxLayout, QTableWidget, QComboBox

class Tab1(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()

        self.label1 = QLabel('rowCount : 0')
        self.label2 = QLabel('columnCount : 0')
        self.table_widget = QTableWidget()
        self.excel_download_btn = QPushButton('다운로드(업로드 파일 기준)')
        self.excel_download_btn2 = QPushButton('다운로드(테이블 데이터 기준)')

        self.layout.addLayout(layout2)
        layout2.addWidget(self.label1)
        layout2.addWidget(self.label2)
        self.layout.addWidget(self.table_widget)
        self.layout.addLayout(layout3)
        layout3.addWidget(self.excel_download_btn)
        layout3.addWidget(self.excel_download_btn2)
        self.setLayout(self.layout)

        self.excel_download_btn.clicked.connect(self.main_window.dataframe_to_excel)
        self.excel_download_btn2.clicked.connect(self.main_window.table_to_excel)
        self.table_widget.horizontalHeader().sortIndicatorChanged.connect(self.main_window.sort_table)

class Tab2(QWidget):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
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

        self.file_list_btn.clicked.connect(self.main_window.show_listwidget_exec)
        self.excel_download_btn.clicked.connect(self.main_window.dataframe_to_excel)
        self.excel_download_btn2.clicked.connect(self.main_window.table_to_excel)
        self.table_widget.horizontalHeader().sortIndicatorChanged.connect(self.main_window.sort_table)
        self.file_path = None

    def combobox1_add_items(self, file_path):
        try:
            if self.combobox1.count() > 0:
                self.combobox1.currentTextChanged.disconnect(self.combobox_value_changed)
                self.combobox1.clear()
            self.file_path = file_path
            sheet_names = dataframes.return_sheets(file_path)
            self.combobox1.addItems(sheet_names)
            current_sheet_name = self.combobox1.currentText()
            self.combobox1.currentTextChanged.connect(self.combobox_value_changed)
            return file_path, current_sheet_name
        except Exception as e:
            print(e)

    def combobox_value_changed(self):
        current_sheet_name = self.combobox1.currentText()
        df = dataframes.file_change(self.file_path, current_sheet_name)
        self.main_window.df_to_table(df)