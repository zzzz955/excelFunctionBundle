from PyQt5.QtWidgets import QComboBox, QRadioButton, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QMessageBox

class PopupDialog(QDialog):
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
        self.close_btn = QPushButton()
        self.close_btn.setText('취소')
        layout4.addWidget(self.close_btn)
        self.close_btn.clicked.connect(self.cancle_func)

        self.colCount = colCount
        self.header_col = header_col
        self.cmb_add_item()

        self.selected_combo_item = None
        self.selected_combo_item2 = None
        self.selected_radio_button = None

    def cmb_add_item(self):
        for col in range(self.colCount):
            header_item = self.header_col[col]
            if header_item:
                self.cmb1.addItem(header_item)
                self.cmb2.addItem(header_item)

    def accept_func(self):
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

    def cancle_func(self):
        self.reject()