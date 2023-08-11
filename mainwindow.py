import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTabWidget
from qt_material import apply_stylesheet
from tabs import *
from dialogs import *

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('엑셀 업무 툴')
        self.setGeometry(0, 0, 600, 800)
        self.layout = QVBoxLayout()
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        layout2 = QHBoxLayout()
        self.layout.addLayout(layout2)

        self.tab_widget = QTabWidget()
        self.layout.addWidget(self.tab_widget)

        self.excel_upload_btn = QPushButton()
        self.excel_upload_btn.setText('단일 시트 엑셀 파일 업로드')
        layout2.addWidget(self.excel_upload_btn)
        self.excel_upload_btn.clicked.connect(self.single_sheet_excel_file_upload)

        self.excel_upload_btn2 = QPushButton()
        self.excel_upload_btn2.setText('다중 시트 엑셀 파일 업로드')
        layout2.addWidget(self.excel_upload_btn2)
        self.excel_upload_btn2.clicked.connect(self.multiple_sheet_excel_file_upload)

        self.func_button = QPushButton()
        self.func_button.setText('기능')
        layout2.addWidget(self.func_button)
        self.func_button.clicked.connect(self.func_Bundle_exec)

        self.close_button = QPushButton('종료', self)
        self.close_button.clicked.connect(self.close_app)
        self.layout.addWidget(self.close_button)

        # 탭 노출
        self.tab1 = Tab1(self)
        self.tab2 = Tab2(self)
        self.tab_widget.addTab(self.tab1, '엑셀 파일 병합')
        self.tab_widget.addTab(self.tab2, '')

    def single_sheet_excel_file_upload(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, '파일 선택', '', 'Excel Files(*.xlsx)')
        if file_paths:
            self.tab_widget.currentWidget().single_sheet_excel_upload(file_paths)

    def multiple_sheet_excel_file_upload(self):
        try:
            file_paths, _ = QFileDialog.getOpenFileNames(self, '파일 선택', '', 'Excel Files(*.xlsx)')
            if file_paths:
                self.tab_widget.currentWidget().multiple_sheet_excel_file_upload(file_paths)
        except Exception as e:
            print(e)

    def func_Bundle_exec(self):
        dialog = func_Bundle(self)
        dialog.exec_()

    def exit_group_by(self):
        # 집계 테이블 원 상태로 복구
        if self.tab_widget.currentWidget().reserve_table_widget.rowCount() > 0:
            self.tab_widget.currentWidget().exit_group_by()

    def group_by_dialog(self):
        # 집계 관련 다이얼 로그 노출
        if self.tab_widget.currentWidget().reserve_table_widget:
            colCount = self.tab_widget.currentWidget().reserve_table_widget.columnCount()
            header_col = []
            for column in range(colCount):
                header_item = self.tab_widget.currentWidget().reserve_table_widget.horizontalHeaderItem(column)
                header_col.append(header_item.text())
            dialog = groupby_Fucn(self, colCount, header_col)
            result = dialog.exec_()

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

    def do_group_by(self, cmb1, cmb2, radio_btn1):
        # 다이얼 로그 값을 받아와 GROUP BY 기능 실행
        df = self.tab_widget.currentWidget().df
        group = df.groupby(by=cmb1, as_index=False)[cmb2].agg(radio_btn1)
        group_sorted = group.sort_values(by=cmb2, ascending=False)
        self.tab_widget.currentWidget().df_to_table(group_sorted)

    def close_app(self):
        # 앱 종료
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='custom.xml')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())