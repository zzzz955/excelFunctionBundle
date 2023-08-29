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
        self.setGeometry(0, 0, 800, self.height())
        self.layout = QVBoxLayout()
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        layout2 = QHBoxLayout()
        self.tab_widget = QTabWidget()
        self.single_sheet_excel_upload_btn = QPushButton('엑셀 파일 업로드(단일 시트)')
        self.multiple_sheet_excel_upload_btn = QPushButton('엑셀 파일 업로드(다중 시트)')
        self.func_button = QPushButton('기능')
        self.initialize_func_button = QPushButton('기능 초기화')
        self.close_button = QPushButton('종료')

        self.layout.addLayout(layout2)
        self.layout.addWidget(self.tab_widget)
        layout2.addWidget(self.single_sheet_excel_upload_btn)
        layout2.addWidget(self.multiple_sheet_excel_upload_btn)
        layout2.addWidget(self.func_button)
        layout2.addWidget(self.initialize_func_button)
        self.layout.addWidget(self.close_button)

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
        self.header = []
        self.rows = None
        self.single_file_paths = None
        self.multiple_file_paths = None
        self.df = None
        self.file_paths = None
        self.modeless_dialog = show_listwidget(self)

    def data_update(self):
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
        if self.single_file_paths:
            file_paths = self.single_file_paths.copy()
            self.single_sheet_excel_file_Conversion(file_paths)
        elif self.multiple_file_paths:
            file_paths = self.multiple_file_paths.copy()
            self.multiple_sheet_excel_file_Conversion(file_paths)
        else:
            return

    def sort_table(self, logical_index, order):
        # 각 헤더에 맞게 정렬
        if self.tab_widget.currentWidget().table_widget.rowCount() > 0:
            self.tab_widget.currentWidget().table_widget.sortItems(logical_index, order)

    def single_sheet_excel_file_upload(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, '파일 선택', '', 'Excel Files(*.xlsx)')
        if file_paths:
            self.single_file_paths = file_paths
            if self.tab_widget.currentIndex() == 0:
                self.single_sheet_excel_file_Conversion(file_paths)
            elif self.tab_widget.currentIndex() == 1:
                self.file_paths = file_paths
                self.modeless_dialog.list_widget.clear()
                self.modeless_dialog.list_widget.addItems(file_paths)
                self.modeless_dialog.show()

    def multiple_sheet_excel_file_upload(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, '파일 선택', '', 'Excel Files(*.xlsx)')
        if file_paths:
            self.multiple_file_paths = file_paths
            if self.tab_widget.currentIndex() == 0:
                self.multiple_sheet_excel_file_Conversion(file_paths)
            elif self.tab_widget.currentIndex() == 1:
                self.file_paths = file_paths
                self.modeless_dialog.list_widget.clear()
                self.modeless_dialog.list_widget.addItems(file_paths)
                self.modeless_dialog.show()

    def single_sheet_excel_file_Conversion(self, file_paths):
        # 엑셀 업로드 함수
        df = dataframes.concat_singlesheet_excelfiles(file_paths)
        self.df_to_table(df)
        self.df = df

    def multiple_sheet_excel_file_Conversion(self, file_paths):
        df = dataframes.concat_multiplesheets_excelfiles(file_paths)
        self.df_to_table(df)
        self.df = df

    def list_widget_exec(self, file_path):
        self.tab_widget.setCurrentIndex(1)
        file_path, current_sheet_name = self.tab2.combobox1_add_items(file_path)
        df = dataframes.file_change(file_path, current_sheet_name)
        self.df_to_table(df)

    def df_to_table(self, df):
        # 데이터 프레임 테이블화
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
        # 행 및 열 개수 노출
        self.tab_widget.currentWidget().label1.setText(f'rowCount : {str(self.tab_widget.currentWidget().table_widget.rowCount())}')
        self.tab_widget.currentWidget().label2.setText(f'columnCount : {str(self.tab_widget.currentWidget().table_widget.columnCount())}')
        for column in range(self.tab_widget.currentWidget().table_widget.columnCount()):
            header_item = self.tab_widget.currentWidget().table_widget.horizontalHeaderItem(column)
            self.header.append(header_item.text())
        self.rows = self.tab_widget.currentWidget().table_widget.rowCount()

    def func_Bundle_exec(self):
        dialog = func_Bundle(self)
        dialog.exec()

    def show_listwidget_exec(self):
        self.modeless_dialog.show()

    def tab_changed(self):
        self.header = []
        for column in range(self.tab_widget.currentWidget().table_widget.columnCount()):
            header_item = self.tab_widget.currentWidget().table_widget.horizontalHeaderItem(column)
            self.header.append(header_item.text())

    def group_by_dialog(self):
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

    def duplicate_dialog(self):
        dialog = duplicate_Fucn(self, self.header)
        result = dialog.exec()
        if result == QDialog.Accepted:
            cmb1 = dialog.selected_combo_item
            radio_btn1 = dialog.selected_radio_button
            # 탭 함수 호출
            self.do_duplicate(cmb1, radio_btn1)
        elif result == QDialog.Rejected:
            return

    def insert_col_dialog(self):
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

    def insert_row_dialog(self):
        if self.tab_widget.currentWidget().table_widget:
            dialog = insert_row_Func(self, self.header, self.rows)
            result = dialog.exec()
            if result == QDialog.Accepted:
                row_index = int(dialog.selected_row_index)
                datas = dialog.insert_list
                radio_btn = dialog.selected_radio_button
                self.do_insert_row(row_index, radio_btn, datas)
            elif result == QDialog.Rejected:
                return

    def do_group_by(self, cmb1, cmb2, radio_btn1):
        # 다이얼 로그 값을 받아와 GROUP BY 기능 실행
        group_sorted = dataframes.group_by_data(self.df, cmb1, cmb2, radio_btn1)
        self.df_to_table(group_sorted)

    def do_duplicate(self, cmb1, radio_btn1):
        # 다이얼 로그 값을 받아와 GROUP BY 기능 실행
        duplicated = dataframes.duplicate_data(self.df, cmb1, radio_btn1)
        self.df_to_table(duplicated)

    def do_insert_col(self, cmb, radio_btn, insert_header, insert_val):
        self.tab_widget.currentWidget()\
            .table_widget.insertColumn(cmb+radio_btn)
        self.tab_widget.currentWidget()\
            .table_widget.setHorizontalHeaderItem(cmb+radio_btn, QTableWidgetItem(insert_header))
        for row in range(self.tab_widget.currentWidget().table_widget.rowCount()):
            self.tab_widget.currentWidget().table_widget.setItem(row, cmb+radio_btn, QTableWidgetItem(insert_val))
        self.data_update()

    def do_insert_row(self, row_index, radio_btn, datas):
        for index, data in enumerate(datas):
            self.tab_widget.currentWidget().table_widget.insertRow(row_index + radio_btn + index - 1)
            for col in range(len(data)):
                self.tab_widget.currentWidget().table_widget.setItem(
                    row_index + radio_btn + index - 1, col, QTableWidgetItem(data[col]))
            self.data_update()

    def dataframe_to_excel(self):
        if self.tab_widget.currentWidget().table_widget.rowCount() > 0:
            save_file_path, _ = QFileDialog.getSaveFileName(self, '파일 선택', '', 'Excel File(*.xlsx)')
            if not save_file_path:
                return
            if self.tab_widget.currentIndex() == 0:
                dataframes.df_to_excel(self.df, save_file_path)
            elif self.tab_widget.currentIndex() == 1:
                dataframes.merge_to_excel_download(self.file_paths, save_file_path)
            self.open_filepath(save_file_path)

    def table_to_excel(self):
        try:
            # 테이블 데이터프레임화 및 엑셀 파일 저장
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