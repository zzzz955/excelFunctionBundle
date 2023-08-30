import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget, QTableWidget, QTableWidgetItem, QFileDialog
from qt_material import apply_stylesheet
from tabs import *
from dialogs import *

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('엑셀 업무 툴')
        self.setGeometry(0, 0, 800, 800)
        self.layout = QVBoxLayout()
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        # 위젯 추가
        layout2 = QHBoxLayout()
        self.tab_widget = QTabWidget()
        self.single_sheet_excel_upload_btn = QPushButton('엑셀 파일 업로드(단일 시트)')
        self.multiple_sheet_excel_upload_btn = QPushButton('엑셀 파일 업로드(다중 시트)')
        self.func_button = QPushButton('기능')
        self.initialize_func_button = QPushButton('기능 초기화')
        self.close_button = QPushButton('종료')

        # 레이아웃 지정
        self.layout.addLayout(layout2)
        self.layout.addWidget(self.tab_widget)
        layout2.addWidget(self.single_sheet_excel_upload_btn)
        layout2.addWidget(self.multiple_sheet_excel_upload_btn)
        layout2.addWidget(self.func_button)
        layout2.addWidget(self.initialize_func_button)
        self.layout.addWidget(self.close_button)

        # 시그널 추가
        self.single_sheet_excel_upload_btn.clicked.connect(self.single_sheet_excel_file_upload)
        self.multiple_sheet_excel_upload_btn.clicked.connect(self.multiple_sheet_excel_file_upload)
        self.func_button.clicked.connect(self.func_Bundle_exec)
        self.initialize_func_button.clicked.connect(self.initialize_func)
        self.tab_widget.currentChanged.connect(self.tab_changed)
        self.close_button.clicked.connect(self.close_app)

        # 탭 노출
        self.tab1 = Tab1(self)
        self.tab2 = Tab2(self)
        self.tab_widget.addTab(self.tab1, '데이터 병합본')
        self.tab_widget.addTab(self.tab2, '시트별 구분')

        # 변수 초기화
        self.header = []
        self.rows = None
        self.single_file_paths = None
        self.multiple_file_paths = None
        self.tab1_df = None
        self.tab2_df = None
        self.file_paths = None
        self.modeless_dialog = show_listwidget(self)

    def data_update(self):
        # 데이터 최신화 함수
        self.rows = self.tab_widget.currentWidget().table_widget.rowCount()
        self.header = []
        for column in range(self.tab_widget.currentWidget().table_widget.columnCount()):
            header_item = self.tab_widget.currentWidget().table_widget.horizontalHeaderItem(column)
            self.header.append(header_item.text())
        self.tab_widget.currentWidget().label1.setText(
            f'rowCount : {str(self.tab_widget.currentWidget().table_widget.rowCount())}')
        self.tab_widget.currentWidget().label2.setText(
            f'columnCount : {str(self.tab_widget.currentWidget().table_widget.columnCount())}')

    def initialize_func(self):
        # 기능 초기화 함수
        if self.tab_widget.currentIndex() == 0 and self.tab1_df:
            self.df_to_table(self.tab1_df)
        elif self.tab_widget.currentIndex() == 1 and self.tab2_df:
            self.df_to_table(self.tab2_df)
        else:
            return

    def sort_table(self, logical_index, order):
        # 테이블 위젯 정렬 함수
        if self.tab_widget.currentWidget().table_widget.rowCount() > 0:
            self.tab_widget.currentWidget().table_widget.sortItems(logical_index, order)

    def single_sheet_excel_file_upload(self):
        # 단일 시트 엑셀 파일 업로드
        try:
            file_paths, _ = QFileDialog.getOpenFileNames(self, '파일 선택', '', 'Excel Files(*.xlsx)')
            if file_paths:
                self.single_file_paths = file_paths
                if self.tab_widget.currentIndex() == 0:
                    self.single_sheet_excel_file_conversion(file_paths)
                elif self.tab_widget.currentIndex() == 1:
                    self.file_paths = file_paths
                    self.modeless_dialog.list_widget.clear()
                    self.modeless_dialog.list_widget.addItems(file_paths)
                    self.modeless_dialog.show()
        except Exception as e:
            QMessageBox.critical(self, '예외 발생', f'엑셀 파일을 불러 올 수 없습니다. {e}'
                                          f'\n1. 파일이 열려 있는 상태 인지 확인해 주세요.'
                                          f'\n2. 파일이 올바른 형식 인지 확인해 주세요.'
                                          f'\n3. 파일 내부 데이터에 문제가 있을 수 있습니다.'
                                          f'\n → [*.*], (*.*), /| 등의 특수 문자를 제거 후 시도해 주세요.')

    def multiple_sheet_excel_file_upload(self):
        # 다중 시트 엑셀 파일 업로드
        try:
            file_paths, _ = QFileDialog.getOpenFileNames(self, '파일 선택', '', 'Excel Files(*.xlsx)')
            if file_paths:
                self.multiple_file_paths = file_paths
                if self.tab_widget.currentIndex() == 0:
                    self.multiple_sheet_excel_file_conversion(file_paths)
                elif self.tab_widget.currentIndex() == 1:
                    self.file_paths = file_paths
                    self.modeless_dialog.list_widget.clear()
                    self.modeless_dialog.list_widget.addItems(file_paths)
                    self.modeless_dialog.show()
        except Exception as e:
            QMessageBox.critical(self, '예외 발생', f'엑셀 파일을 불러 올 수 없습니다. {e}'
                                          f'\n1. 파일이 열려 있는 상태 인지 확인해 주세요.'
                                          f'\n2. 파일이 올바른 형식 인지 확인해 주세요.'
                                          f'\n3. 파일 내부 데이터에 문제가 있을 수 있습니다.'
                                          f'\n → [*.*], (*.*), /| 등의 특수 문자를 제거 후 시도해 주세요.')

    def single_sheet_excel_file_conversion(self, file_paths):
        # 엑셀 파일 병합
        df = dataframes.concat_singlesheet_excelfiles(file_paths)
        self.df_to_table(df)
        self.tab1_df = df

    def multiple_sheet_excel_file_conversion(self, file_paths):
        # 엑셀 파일 병합
        df = dataframes.concat_multiplesheets_excelfiles(file_paths)
        self.df_to_table(df)
        self.tab1_df = df

    def list_widget_exec(self, file_path):
        # 탭2 파일 목록 노출
        self.tab_widget.setCurrentIndex(1)
        file_path, current_sheet_name = self.tab2.combobox1_add_items(file_path)
        df = dataframes.file_change(file_path, current_sheet_name)
        self.tab2_df = df
        self.df_to_table(df)

    def df_to_table(self, df):
        # 데이터 프레임 테이블 위젯 추가
        self.header = []
        self.tab_widget.currentWidget().table_widget.clear()
        self.tab_widget.currentWidget().table_widget.setRowCount(len(df))
        self.tab_widget.currentWidget().table_widget.setColumnCount(len(df.columns))
        self.tab_widget.currentWidget().table_widget.setHorizontalHeaderLabels(df.columns)
        self.tab_widget.currentWidget().table_widget.horizontalHeader().setSortIndicatorShown(True)
        for r in range(len(df)):
            for c in range(len(df.columns)):
                item = str(df.iloc[r, c])
                self.tab_widget.currentWidget().table_widget.setItem(r, c, QTableWidgetItem(item))
        self.tab_widget.currentWidget().table_widget.resizeColumnsToContents()
        self.tab_widget.currentWidget().label1.setText(f'rowCount : {str(self.tab_widget.currentWidget().table_widget.rowCount())}')
        self.tab_widget.currentWidget().label2.setText(f'columnCount : {str(self.tab_widget.currentWidget().table_widget.columnCount())}')
        for column in range(self.tab_widget.currentWidget().table_widget.columnCount()):
            header_item = self.tab_widget.currentWidget().table_widget.horizontalHeaderItem(column)
            self.header.append(header_item.text())
        self.rows = self.tab_widget.currentWidget().table_widget.rowCount()

    def func_Bundle_exec(self):
        # 기능 다이얼 로그 호출
        dialog = func_Bundle(self)
        dialog.show()

    def show_listwidget_exec(self):
        # 리스트 위젯 다이얼 로그 호출
        self.modeless_dialog.show()

    def tab_changed(self):
        # 탭 변경 시 헤더 데이터 변경
        self.header = []
        for column in range(self.tab_widget.currentWidget().table_widget.columnCount()):
            header_item = self.tab_widget.currentWidget().table_widget.horizontalHeaderItem(column)
            self.header.append(header_item.text())

    def group_by_dialog(self):
        # 집계 함수 다이얼 로그 호출
        dialog = groupby_Func(self, self.header)
        result = dialog.exec()
        if result == QDialog.Accepted:
            cmb1 = dialog.selected_combo_item
            cmb2 = dialog.selected_combo_item2
            radio_btn1 = dialog.selected_radio_button
            self.do_group_by(cmb1, cmb2, radio_btn1)
        elif result == QDialog.Rejected:
            return

    def duplicate_dialog(self):
        # 중복 제거 다이얼 로그 호출
        dialog = duplicate_Func(self, self.header)
        result = dialog.exec()
        if result == QDialog.Accepted:
            cmb1 = dialog.selected_combo_item
            radio_btn1 = dialog.selected_radio_button
            # 탭 함수 호출
            self.do_duplicate(cmb1, radio_btn1)
        elif result == QDialog.Rejected:
            return

    def insert_col_dialog(self):
        # 열 삽입 다이얼 로그 호출
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

    def insert_row_dialog(self):
        # 행 삽입 다이얼 로그 호출
        dialog = insert_row_Func(self, self.header, self.rows)
        result = dialog.exec()
        if result == QDialog.Accepted:
            row_index = int(dialog.selected_row_index)
            datas = dialog.insert_list
            radio_btn = dialog.selected_radio_button
            self.do_insert_row(row_index, radio_btn, datas)
        elif result == QDialog.Rejected:
            return

    def delete_col_dialog(self):
        try:
        # 열 삭제 다이얼 로그 호출
            dialog = delete_col_Func(self, self.header)
            dialog.exec()
        except Exception as e:
            print(e)

    def delete_row_dialog(self):
        # 행 삭제 다이얼 로그 호출
        try:
            dialog = delete_row_Func(self, self.rows)
            dialog.exec()
        except Exception as e:
            print(e)

    def replace_dialog(self):
        # 찾아 바꾸기 다이얼 로그 호출
        dialog = replace_Func(self)
        dialog.show()

    def clear_black_row(self):
        # 빈 행 삭제 다이얼 로그 호출
        blank_rows_index = []
        for row in range(self.tab_widget.currentWidget().table_widget.rowCount()):
            items = []
            for col in range(self.tab_widget.currentWidget().table_widget.columnCount()):
                item = self.tab_widget.currentWidget().table_widget.item(row, col)
                items.append(item.text())
            if all(item == '' for item in items):
                blank_rows_index.append(row)
        for i in reversed(blank_rows_index):
            self.tab_widget.currentWidget().table_widget.removeRow(i)

    def do_group_by(self, cmb1, cmb2, radio_btn1):
        # 다이얼 로그 값을 받아와 집계 기능 실행
        if self.tab_widget.currentIndex() == 0:
            group_sorted = dataframes.group_by_data(self.tab1_df, cmb1, cmb2, radio_btn1)
            self.df_to_table(group_sorted)
        if self.tab_widget.currentIndex() == 1:
            group_sorted = dataframes.group_by_data(self.tab2_df, cmb1, cmb2, radio_btn1)
            self.df_to_table(group_sorted)

    def do_duplicate(self, cmb1, radio_btn1):
        # 다이얼 로그 값을 받아와 중복 제거 기능 실행
        if self.tab_widget.currentIndex() == 0:
            duplicated = dataframes.duplicate_data(self.tab1_df, cmb1, radio_btn1)
            self.df_to_table(duplicated)
        if self.tab_widget.currentIndex() == 1:
            duplicated = dataframes.duplicate_data(self.tab2_df, cmb1, radio_btn1)
            self.df_to_table(duplicated)

    def do_insert_col(self, cmb, radio_btn, insert_header, insert_val):
        # 다이얼 로그 값을 받아와 열 삽입 기능 실행
        self.tab_widget.currentWidget()\
            .table_widget.insertColumn(cmb+radio_btn)
        self.tab_widget.currentWidget()\
            .table_widget.setHorizontalHeaderItem(cmb+radio_btn, QTableWidgetItem(insert_header))
        for row in range(self.tab_widget.currentWidget().table_widget.rowCount()):
            self.tab_widget.currentWidget().table_widget.setItem(row, cmb+radio_btn, QTableWidgetItem(insert_val))
        self.data_update()

    def do_insert_row(self, row_index, radio_btn, datas):
        # 다이얼 로그 값을 받아와 행 삽입 기능 실행
        for index, data in enumerate(datas):
            self.tab_widget.currentWidget().table_widget.insertRow(row_index + radio_btn + index - 1)
            for col in range(len(data)):
                self.tab_widget.currentWidget().table_widget.setItem(
                    row_index + radio_btn + index - 1, col, QTableWidgetItem(data[col]))
            self.data_update()

    def do_delete_col(self, first_col_index, last_col_index):
        # 다이얼 로그 값을 받아와 열 삭제 기능 실행
        for i in range(last_col_index-first_col_index+1):
            self.tab_widget.currentWidget().table_widget.removeColumn(first_col_index)
        self.data_update()

    def do_delete_row(self, first_row_index, last_row_index):
        # 다이얼 로그 값을 받아와 열 삭제 기능 실행
        try:
            if last_row_index-first_row_index >= 0:
                rows = last_row_index-first_row_index
                start_row = first_row_index
            elif last_row_index-first_row_index < 0:
                rows = first_row_index-last_row_index
                start_row = last_row_index
            else:
                QMessageBox.warning(self, '경고', 'row_index를 성공적으로 불러오지 못했습니다.'
                                                '올바른 값이 입력 되었는지 확인해 주세요.')
                return
            for i in range(rows):
                self.tab_widget.currentWidget().table_widget.removeRow(start_row)
            self.data_update()
        except Exception as e:
            print(e)

    def do_replace(self, find_text, replace_text, if_exact):
        for row in range(self.tab_widget.currentWidget().table_widget.rowCount()):
            for col in range(self.tab_widget.currentWidget().table_widget.columnCount()):
                item = self.tab_widget.currentWidget().table_widget.item(row, col)
                if item:
                    text = item.text()
                    if find_text == text:
                        text = replace_text
                        item.setText(text)
                    elif not if_exact:
                        new_text = re.sub(find_text, replace_text, text)
                        item.setText(new_text)

    def dataframe_to_excel(self):
        # 데이터 프레임 기준 엑셀 파일 추출
        if self.tab_widget.currentWidget().table_widget.rowCount() > 0:
            save_file_path, _ = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')
            if not save_file_path:
                return
            if self.tab_widget.currentIndex() == 0:
                dataframes.df_to_excel(self.tab1_df, save_file_path)
            elif self.tab_widget.currentIndex() == 1:
                dataframes.merge_to_excel_download(self.file_paths, save_file_path)
            self.open_filepath(save_file_path)

    def table_to_excel(self):
        # 테이블 데이터 기준 엑셀 파일 추출
        try:
            # 테이블 데이터 프레임 화 및 엑셀 파일 저장
            if self.tab_widget.currentWidget().table_widget.rowCount() > 0:
                save_file_path, _ = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')
                if not save_file_path:
                    return
                header = []
                for column in range(self.tab_widget.currentWidget().table_widget.columnCount()):
                    header_item = self.tab_widget.currentWidget().table_widget.horizontalHeaderItem(column)
                    header.append(header_item.text())
                data = []
                for r in range(self.tab_widget.currentWidget().table_widget.rowCount()):
                    rowdata = []
                    for c in range(self.tab_widget.currentWidget().table_widget.columnCount()):
                        item = self.tab_widget.currentWidget().table_widget.item(r, c)
                        rowdata.append(item.text())
                    data.append(rowdata)
                dataframes.table_to_excel(save_file_path, header, data)
                self.open_filepath(save_file_path)
        except Exception as e:
            QMessageBox.warning(self, '경고', f'엑셀 파일을 저장할 수 없습니다. 해당 파일이 열려 있는 상태가 아닌지 확인해 보세요 {e}')

    def open_filepath(self, file_path):
        # 저장 후 파일 열기 기능
        result = QMessageBox.question(self, '정보', '엑셀 파일 생성 완료, 생성한 파일을 여시겠습니까?',
                                      QMessageBox.Ok | QMessageBox.No,
                                      QMessageBox.Ok)
        if result == QMessageBox.Ok:
            os.startfile(file_path)
        else:
            return

    def close_app(self):
        # 앱 종료
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='custom.xml')
    window = MainWindow()
    window.show()
    sys.exit(app.exec())