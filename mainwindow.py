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
        self.tab_widget = QTabWidget()
        self.layout.addWidget(self.tab_widget)
        self.close_button = QPushButton('종료', self)
        self.close_button.clicked.connect(self.close_app)
        self.layout.addWidget(self.close_button)

        # 탭 노출
        self.tab1 = Tab1(self)
        self.tab2 = Tab2(self)
        self.tab_widget.addTab(self.tab1, '엑셀 파일 병합 및 집계')
        self.tab_widget.addTab(self.tab2, '')

    def open_tab1_dialog(self):
        # 집계 관련 다이얼 로그 노출
        colCount = self.tab1.reserve_table_widget.columnCount()
        header_col = []
        for column in range(colCount):
            header_item = self.tab1.reserve_table_widget.horizontalHeaderItem(column)
            header_col.append(header_item.text())
        dialog = PopupDialog(self, colCount, header_col)
        result = dialog.exec_()

        # 다이얼 로그로 부터 값 가져오기
        if result == QDialog.Accepted:
            cmb1 = dialog.selected_combo_item
            cmb2 = dialog.selected_combo_item2
            radio_btn1 = dialog.selected_radio_button
            # 탭 함수 호출
            self.tab1.do_group_by(cmb1, cmb2, radio_btn1)

        elif result == QDialog.Rejected:
            return

    def close_app(self):
        # 앱 종료
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    apply_stylesheet(app, theme='custom.xml')
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())