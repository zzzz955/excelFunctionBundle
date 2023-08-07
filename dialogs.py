from PyQt5.QtWidgets import QComboBox, QRadioButton, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QMessageBox

class func_Bundle(QDialog):
    def __init__(self, main_window):
        super().__init__()
        self.setWindowTitle('기능 모음')
        self.main_window = main_window
        layout = QVBoxLayout(self)

        layout2 = QHBoxLayout()
        layout.addLayout(layout2)

        self.do_group_by_btn = QPushButton()
        self.do_group_by_btn.setText('집계 함수 실행')
        layout2.addWidget(self.do_group_by_btn)
        self.do_group_by_btn.clicked.connect(self.connect_dialog)

        self.exit_group_by_btn = QPushButton()
        self.exit_group_by_btn.setText('집계 함수 해제')
        layout2.addWidget(self.exit_group_by_btn)
        self.exit_group_by_btn.clicked.connect(self.exit_group_by)

        self.exit_dialog_btn = QPushButton()
        self.exit_dialog_btn.setText('종료')
        layout.addWidget(self.exit_dialog_btn)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)
        if hasattr(self.main_window.tab_widget.currentWidget(), 'reserve_table_widget'):
            self.table_data = self.main_window.tab_widget.currentWidget().reserve_table_widget
        else:
            self.table_data=None
        self.button_toggle()

    def button_toggle(self):
        try:
            if hasattr(self.main_window.tab_widget.currentWidget(), 'reserve_table_widget'):
                self.do_group_by_btn.setEnabled(True)
                self.exit_group_by_btn.setEnabled(True)
            else:
                self.do_group_by_btn.setEnabled(False)
                self.exit_group_by_btn.setEnabled(False)
        except Exception as e:
            QMessageBox(self, '경고', f'{e}')

    def connect_dialog(self):
        # 집계 관련 다이얼로그 호출 함수
        try:
            if self.table_data.rowCount() > 0:
                self.accept()
                self.main_window.group_by_dialog()
        except Exception as e:
            QMessageBox(self, '경고', f'{e}')

    def exit_group_by(self):
        self.main_window.exit_group_by()

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.reject()

class groupby_Fucn(QDialog):
    # 집계 기능 다이얼 로그
    def __init__(self, main_window, colCount, header_col):
        super().__init__()
        self.setWindowTitle('집계 팝업')
        self.main_window = main_window
        layout = QVBoxLayout(self)
        layout2 = QHBoxLayout(self)

        layout.addLayout(layout2)

        self.label1 = QLabel('기준 Column')
        layout2.addWidget(self.label1)
        self.cmb1 = QComboBox()
        layout2.addWidget(self.cmb1)
        self.label2 = QLabel('집계 Column')
        layout2.addWidget(self.label2)
        self.cmb2 = QComboBox()
        layout2.addWidget(self.cmb2)

        layout3 = QHBoxLayout(self)
        layout.addLayout(layout3)

        self.label3 = QLabel('집계 함수 선택 : ')
        layout3.addWidget(self.label3)
        self.radio_btn1 = QRadioButton('개수', self)
        layout3.addWidget(self.radio_btn1)
        self.radio_btn2 = QRadioButton('합계', self)
        layout3.addWidget(self.radio_btn2)
        self.radio_btn3 = QRadioButton('최소 값', self)
        layout3.addWidget(self.radio_btn3)
        self.radio_btn4 = QRadioButton('최대 값', self)
        layout3.addWidget(self.radio_btn4)
        self.radio_btn5 = QRadioButton('평균 값', self)
        layout3.addWidget(self.radio_btn5)

        layout4 = QHBoxLayout(self)
        layout.addLayout(layout4)

        self.accept_btn = QPushButton()
        self.accept_btn.setText('집계 실행')
        layout4.addWidget(self.accept_btn)
        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn = QPushButton()
        self.exit_dialog_btn.setText('취소')
        layout4.addWidget(self.exit_dialog_btn)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

        self.colCount = colCount
        self.header_col = header_col
        self.cmb_add_item()

        self.selected_combo_item = None
        self.selected_combo_item2 = None
        self.selected_radio_button = None

    def cmb_add_item(self):
        # 콤보박스에 테이블의 헤더 목록 가져오기
        for col in range(self.colCount):
            header_item = self.header_col[col]
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