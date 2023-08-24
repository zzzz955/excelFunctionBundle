from PyQt5.QtWidgets import QComboBox, QRadioButton, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QMessageBox, QLineEdit

class func_Bundle(QDialog):
    def __init__(self, main_window):
        super().__init__()
        self.setWindowTitle('기능 모음')
        self.main_window = main_window

        layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        self.do_group_by_btn = QPushButton('집계 함수 실행')
        self.exit_group_by_btn = QPushButton('집계 함수 해제')
        self.do_insert_btn_v = QPushButton('열 일괄 삽입')
        self.do_insert_btn_h = QPushButton('행 일괄 삽입')
        self.exit_dialog_btn = QPushButton('종료')

        layout.addLayout(layout2)
        layout2.addWidget(self.do_group_by_btn)
        layout2.addWidget(self.exit_group_by_btn)

        layout.addLayout(layout3)
        layout3.addWidget(self.do_insert_btn_v)
        layout3.addWidget(self.do_insert_btn_h)
        layout.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)

        self.do_group_by_btn.clicked.connect(self.con_group_by_dialog)
        self.exit_group_by_btn.clicked.connect(self.exit_group_by)
        self.do_insert_btn_v.clicked.connect(self.con_col_insert_dialog)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

        if hasattr(self.main_window.tab_widget.currentWidget(), 'table_widget'):
            self.table_data = self.main_window.tab_widget.currentWidget().table_widget
        else:
            self.table_data=None
        self.button_toggle()

    def button_toggle(self):
        if hasattr(self.main_window.tab_widget.currentWidget(), 'table_widget'):
            self.do_group_by_btn.setEnabled(True)
            self.exit_group_by_btn.setEnabled(True)
        else:
            self.do_group_by_btn.setEnabled(False)
            self.exit_group_by_btn.setEnabled(False)

    def con_group_by_dialog(self):
        # 집계 관련 다이얼로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.group_by_dialog()

    def con_col_insert_dialog(self):
        # 삽입 관련 다이얼로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.insert_col_dialog()

    def exit_group_by(self):
        self.main_window.exit_group_by()

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.reject()

class groupby_Func(QDialog):
    def __init__(self, main_window, header):
        super().__init__()
        self.setWindowTitle('집계 팝업')
        self.main_window = main_window

        layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        layout4 = QHBoxLayout()
        self.label1 = QLabel('기준 Column')
        self.cmb1 = QComboBox()
        self.label2 = QLabel('집계 Column')
        self.cmb2 = QComboBox()
        self.label3 = QLabel('집계 함수 선택 : ')
        self.radio_btn1 = QRadioButton('개수', self)
        self.radio_btn2 = QRadioButton('합계', self)
        self.radio_btn3 = QRadioButton('최소 값', self)
        self.radio_btn4 = QRadioButton('최대 값', self)
        self.radio_btn5 = QRadioButton('평균 값', self)
        self.accept_btn = QPushButton('집계 실행')
        self.exit_dialog_btn = QPushButton('취소')

        layout.addLayout(layout2)
        layout2.addWidget(self.label1)
        layout2.addWidget(self.cmb1)
        layout2.addWidget(self.label2)
        layout2.addWidget(self.cmb2)

        layout.addLayout(layout3)
        layout3.addWidget(self.label3)
        layout3.addWidget(self.radio_btn1)
        layout3.addWidget(self.radio_btn2)
        layout3.addWidget(self.radio_btn3)
        layout3.addWidget(self.radio_btn4)
        layout3.addWidget(self.radio_btn5)

        layout.addLayout(layout4)
        layout4.addWidget(self.accept_btn)
        layout4.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)

        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

        self.header = header
        self.cmb_add_item()

        self.selected_combo_item = None
        self.selected_combo_item2 = None
        self.selected_radio_button = None

    def cmb_add_item(self):
        # 콤보박스에 테이블의 헤더 목록 가져오기
        for col in range(len(self.header)):
            header_item = self.header[col]
            if header_item:
                self.cmb1.addItem(header_item)
                self.cmb2.addItem(header_item)

    def accept_func(self):
        # 다이얼 로그 값 전달
        self.selected_combo_item = self.cmb1.currentText()
        self.selected_combo_item2 = self.cmb2.currentText()
        if self.radio_btn1.isChecked():
            self.selected_radio_button = "count"
        elif self.radio_btn2.isChecked():
            self.selected_radio_button = "sum"
        elif self.radio_btn3.isChecked():
            self.selected_radio_button = "min"
        elif self.radio_btn4.isChecked():
            self.selected_radio_button = "max"
        elif self.radio_btn5.isChecked():
            self.selected_radio_button = "mean"
        else:
            QMessageBox.warning(self, '경고', '집계 함수를 선택해 주세요.')
            return
        self.accept()

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.reject()

class insert_col_Func(QDialog):
    def __init__(self, main_window, header):
        super().__init__()
        self.setWindowTitle('데이터 삽입')
        self.main_window = main_window
        self.header = header

        layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        layout4 = QHBoxLayout()
        layout5 = QHBoxLayout()
        layout6 = QHBoxLayout()
        self.label1 = QLabel('기준 Column')
        self.cmb1 = QComboBox()
        self.label2 = QLabel('헤더 지정 : ')
        self.lineedit1 = QLineEdit()
        self.label3 = QLabel('삽입 데이터 : ')
        self.lineedit2 = QLineEdit()
        self.label4 = QLabel('삽입 위치 지정 : ')
        self.radio_btn1 = QRadioButton('기준 칼럼 앞', self)
        self.radio_btn2 = QRadioButton('기준 칼럼 뒤', self)
        self.accept_btn = QPushButton('삽입')
        self.exit_dialog_btn = QPushButton('취소')

        layout.addLayout(layout2)
        layout2.addWidget(self.label1)
        layout2.addWidget(self.cmb1)

        layout.addLayout(layout3)
        layout3.addWidget(self.label2)
        layout3.addWidget(self.lineedit1)

        layout.addLayout(layout4)
        layout4.addWidget(self.label3)
        layout4.addWidget(self.lineedit2)

        layout.addLayout(layout5)
        layout5.addWidget(self.label4)
        layout5.addWidget(self.radio_btn1)
        layout5.addWidget(self.radio_btn2)

        layout.addLayout(layout6)
        layout6.addWidget(self.accept_btn)
        layout6.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)

        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

        self.selected_combo_index = None
        self.selected_radio_button = None
        self.inserted_header_val = None
        self.inserted_data_val = None
        self.cmb1.addItems(header)

    def accept_func(self):
        try:
            self.selected_combo_index = self.cmb1.currentIndex()
            self.inserted_header_val = self.lineedit1.text()
            self.inserted_data_val = self.lineedit2.text()
            if self.radio_btn1.isChecked():
                self.selected_radio_button = 0
            elif self.radio_btn2.isChecked():
                self.selected_radio_button = 1
            else:
                QMessageBox.warning(self, '경고', '데이터 삽입 위치를 선택해 주세요.')
                return
            self.accept()
        except Exception as e:
            print(e)

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.reject()