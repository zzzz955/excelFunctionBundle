import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QTableWidget, QTableWidgetItem
from qt_material import apply_stylesheet
from tabs import *
from dialogs import *

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('엑셀 업무 툴')
        self.setGeometry(0, 0, 800, self.height())
        self.layout = QVBoxLayout()
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        layout2 = QHBoxLayout()
        self.tab_widget = QTabWidget()
        self.single_sheet_excel_upload_btn = QPushButton('단일 시트 엑셀 파일 업로드')
        self.multiple_sheet_excel_upload_btn = QPushButton('다중 시트 엑셀 파일 업로드')
        self.func_button = QPushButton('기능')
        self.close_button = QPushButton('종료')

        self.layout.addLayout(layout2)
        self.layout.addWidget(self.tab_widget)
        layout2.addWidget(self.single_sheet_excel_upload_btn)
        layout2.addWidget(self.multiple_sheet_excel_upload_btn)
        layout2.addWidget(self.func_button)
        self.layout.addWidget(self.close_button)

        self.single_sheet_excel_upload_btn.clicked.connect(self.single_sheet_excel_file_upload)
        self.multiple_sheet_excel_upload_btn.clicked.connect(self.multiple_sheet_excel_file_upload)
        self.func_button.clicked.connect(self.func_Bundle_exec)
        self.tab_widget.currentChanged.connect(self.tab_changed)
        self.close_button.clicked.connect(self.close_app)

        # 탭 노출
        self.tab1 = Tab1(self)
        self.tab2 = Tab2(self)
        self.tab_widget.addTab(self.tab1, '엑셀 파일 병합')
        self.tab_widget.addTab(self.tab2, '단일 시트')
        self.header = []

    def single_sheet_excel_file_upload(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, '파일 선택', '', 'Excel Files(*.xlsx)')
        if file_paths:
            self.single_sheet_excel_file_Conversion(file_paths)

    def multiple_sheet_excel_file_upload(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, '파일 선택', '', 'Excel Files(*.xlsx)')
        if file_paths:
            self.multiple_sheet_excel_file_Conversion(file_paths)

    def single_sheet_excel_file_Conversion(self, file_paths):
        # 엑셀 업로드 함수
        try:
            self.tab_widget.currentWidget().file_paths = file_paths
            df = dataframes.concat_singlesheet_excelfiles(file_paths)
            self.df_to_table(df)
            self.df_to_reserve_table(df)
            self.tab_widget.currentWidget().df = df
        except Exception as e:
            print(e)

    def multiple_sheet_excel_file_Conversion(self, file_paths):
        self.tab_widget.currentWidget().file_paths = file_paths
        df = dataframes.concat_multiplesheets_excelfiles(file_paths)
        self.df_to_table(df)
        self.df_to_reserve_table(df)
        self.tab_widget.currentWidget().df = df

    def df_to_table(self, df):
        # 데이터 프레임 테이블화
        self.header = []
        self.tab_widget.currentWidget().table_widget.setRowCount(len(df))
        self.tab_widget.currentWidget().table_widget.setColumnCount(len(df.columns))
        self.tab_widget.currentWidget().table_widget.setHorizontalHeaderLabels(df.columns)
        self.tab_widget.currentWidget().table_widget.horizontalHeader().setSortIndicatorShown(True)
        for r in range(len(df)):
            for c in range(len(df.columns)):
                item = str(df.iloc[r, c])
                self.tab_widget.currentWidget().table_widget.setItem(r, c, QTableWidgetItem(item))
        self.tab_widget.currentWidget().table_widget.resizeColumnsToContents()
        # 행 및 열 개수 노출
        self.tab_widget.currentWidget().t1label1.setText(f'rowCount : {str(self.tab_widget.currentWidget().table_widget.rowCount())}')
        self.tab_widget.currentWidget().t1label2.setText(f'columnCount : {str(self.tab_widget.currentWidget().table_widget.columnCount())}')
        for column in range(self.tab_widget.currentWidget().table_widget.columnCount()):
            header_item = self.tab_widget.currentWidget().table_widget.horizontalHeaderItem(column)
            self.header.append(header_item.text())

    def df_to_reserve_table(self, df):
        # 기존 병합 데이터프레임 값 임시 테이블에 저장
        self.tab_widget.currentWidget().reserve_table_widget.setRowCount(len(df))
        self.tab_widget.currentWidget().reserve_table_widget.setColumnCount(len(df.columns))
        self.tab_widget.currentWidget().reserve_table_widget.setHorizontalHeaderLabels(df.columns)
        for r in range(len(df)):
            for c in range(len(df.columns)):
                item = str(df.iloc[r, c])
                self.tab_widget.currentWidget().reserve_table_widget.setItem(r, c, QTableWidgetItem(item))

    def func_Bundle_exec(self):
        dialog = func_Bundle(self)
        dialog.exec()

    def exit_group_by(self):
        # 집계 테이블 원 상태로 복구
        if self.tab_widget.currentWidget().reserve_table_widget.rowCount() > 0:
            self.tab_widget.currentWidget().exit_group_by()

    def tab_changed(self):
        self.header = []
        for column in range(self.tab_widget.currentWidget().table_widget.columnCount()):
            header_item = self.tab_widget.currentWidget().table_widget.horizontalHeaderItem(column)
            self.header.append(header_item.text())

    def group_by_dialog(self):
        try:
            # 집계 관련 다이얼 로그 노출
            if self.tab_widget.currentWidget().reserve_table_widget:
                dialog = groupby_Func(self, self.header)
                result = dialog.exec()
                # 다이얼 로그로 부터 값 가져오기
                if result == QDialog.Accepted:
                    cmb1 = dialog.selected_combo_item
                    cmb2 = dialog.selected_combo_item2
                    radio_btn1 = dialog.selected_radio_button
                    # 탭 함수 호출
                    self.do_group_by(cmb1, cmb2, radio_btn1)
                elif result == QDialog.Rejected:
                    return
            else:
                QMessageBox(self,'예외', '현재 탭에선 집계 기능을 사용할 수 없습니다.')
        except Exception as e:
            print(e)

    def insert_col_dialog(self):
        try:
            if self.tab_widget.currentWidget().table_widget:
                dialog = insert_col_Func(self, self.header)
                result = dialog.exec()
                if result == QDialog.Accepted:
                    cmb = dialog.selected_combo_index
                    radio_btn = dialog.selected_radio_button
                    insert_header = dialog.inserted_header_val
                    insert_val = dialog.inserted_data_val
                    self.do_insert_col(cmb, radio_btn, insert_header, insert_val)
                elif result == QDialog.Rejected:
                    return
        except Exception as e:
            print(e)

    def insert_row_dialog(self):
        try:
            if self.tab_widget.currentWidget().table_widget:
                dialog = insert_row_Func(self, self.header)
                result = dialog.exec()
                if result == QDialog.Accepted:
                    row_index = int(dialog.selected_row_index)
                    datas = dialog.insert_list
                    radio_btn = dialog.selected_radio_button
                    self.do_insert_row(row_index, radio_btn, datas)
                elif result == QDialog.Rejected:
                    return
        except Exception as e:
            print(e)

    def do_group_by(self, cmb1, cmb2, radio_btn1):
        # 다이얼 로그 값을 받아와 GROUP BY 기능 실행
        df = self.tab_widget.currentWidget().df
        group = df.groupby(by=cmb1, as_index=False)[cmb2].agg(radio_btn1)
        group_sorted = group.sort_values(by=cmb2, ascending=False)
        self.df_to_table(group_sorted)

    def do_insert_col(self, cmb, radio_btn, insert_header, insert_val):
        self.tab_widget.currentWidget()\
            .table_widget.insertColumn(cmb+radio_btn)
        self.tab_widget.currentWidget()\
            .table_widget.setHorizontalHeaderItem(cmb+radio_btn, QTableWidgetItem(insert_header))
        for row in range(self.tab_widget.currentWidget().table_widget.rowCount()):
            self.tab_widget.currentWidget().table_widget.setItem(row, cmb+radio_btn, QTableWidgetItem(insert_val))

    def do_insert_row(self, row_index, radio_btn, datas):
        try:
            for index, data in enumerate(datas):
                self.tab_widget.currentWidget().table_widget.insertRow(row_index + radio_btn + index - 1)
                for col in range(len(data)):
                    self.tab_widget.currentWidget().table_widget.setItem(
                        row_index + radio_btn + index - 1, col, QTableWidgetItem(data[col]))
        except Exception as e:
            print(e)

    def close_app(self):
        # 앱 종료
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='custom.xml')
    window = MainWindow()
    window.show()
    sys.exit(app.exec())