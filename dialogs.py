from PyQt5.QtWidgets import QComboBox, QRadioButton, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, \
    QMessageBox, QLineEdit, QTableWidget, QListWidget, QGridLayout, QCheckBox
from PyQt5.QtGui import QValidator, QIntValidator
from PyQt5.QtCore import Qt


class func_Bundle(QDialog):
    def __init__(self, main_window):
        super().__init__()
        self.setWindowTitle('기능 모음')
        self.main_window = main_window
        
        # 위젯 추가
        layout = QVBoxLayout()
        layout2 = QGridLayout()
        self.do_group_by_btn = QPushButton('집계 함수 실행')
        self.do_duplicate_btn = QPushButton('중복 제거')
        self.do_insert_btn_h = QPushButton('행 일괄 삽입')
        self.do_delete_btn_h = QPushButton('행 일괄 삭제')
        self.do_insert_btn_v = QPushButton('열 일괄 삽입')
        self.do_delete_btn_v = QPushButton('열 일괄 삭제')
        self.do_replace_btn = QPushButton('문자열 변환')
        self.do_clear_black_row_btn = QPushButton('빈 행 삭제')
        self.do_text_filter_btn = QPushButton('텍스트 필터 적용')
        self.exit_dialog_btn = QPushButton('종료')

        # 레이아웃 지정
        layout.addLayout(layout2)
        layout2.addWidget(self.do_group_by_btn, 0, 0)
        layout2.addWidget(self.do_duplicate_btn, 0, 1)
        layout2.addWidget(self.do_insert_btn_h, 1, 0)
        layout2.addWidget(self.do_delete_btn_h, 1, 1)
        layout2.addWidget(self.do_insert_btn_v, 2, 0)
        layout2.addWidget(self.do_delete_btn_v, 2, 1)
        layout2.addWidget(self.do_replace_btn, 3, 0)
        layout2.addWidget(self.do_clear_black_row_btn, 3, 1)
        layout2.addWidget(self.do_text_filter_btn, 4, 0)
        layout.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)
        
        # 시그널 추가
        self.do_group_by_btn.clicked.connect(self.con_group_by_dialog)
        self.do_duplicate_btn.clicked.connect(self.con_duplicate_dialog)
        self.do_insert_btn_h.clicked.connect(self.con_row_insert_dialog)
        self.do_delete_btn_h.clicked.connect(self.con_row_delete_dialog)
        self.do_insert_btn_v.clicked.connect(self.con_col_insert_dialog)
        self.do_delete_btn_v.clicked.connect(self.con_col_delete_dialog)
        self.do_replace_btn.clicked.connect(self.con_replace_dialog)
        self.do_clear_black_row_btn.clicked.connect(self.con_clear_black_row)
        self.do_text_filter_btn.clicked.connect(self.con_text_filter_dialog)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)
        
        # 데이터 확인
        if hasattr(self.main_window.tab_widget.currentWidget(), 'table_widget'):
            self.table_data = self.main_window.tab_widget.currentWidget().table_widget
        else:
            self.table_data = None

    def con_group_by_dialog(self):
        # 집계 관련 다이얼 로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.group_by_dialog()

    def con_duplicate_dialog(self):
        # 중복 제거 다이얼 로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.duplicate_dialog()

    def con_col_insert_dialog(self):
        # 열 삽입 관련 다이얼 로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.insert_col_dialog()

    def con_row_insert_dialog(self):
        # 행 삽입 관련 다이얼 로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.insert_row_dialog()

    def con_col_delete_dialog(self):
        # 열 삭제 관련 다이얼 로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.delete_col_dialog()

    def con_row_delete_dialog(self):
        # 행 삭제 관련 다이얼 로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.delete_row_dialog()

    def con_replace_dialog(self):
        # 찾아 바꾸기 관련 다이얼 로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.replace_dialog()

    def con_clear_black_row(self):
        # 빈 셀 삭제 관련 다이얼 로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.clear_black_row()

    def con_text_filter_dialog(self):
        # 필터 관련 다이얼 로그 호출 함수
        if self.table_data.rowCount() > 0:
            self.accept()
            self.main_window.text_filter_dialog()

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.reject()


class groupby_Func(QDialog):
    def __init__(self, main_window, header):
        super().__init__()
        self.setWindowTitle('집계 팝업')
        self.main_window = main_window

        # 위젯 추가
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

        # 레이아웃 지정
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

        # 시그널 추가
        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

        # 변수 초기화
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


class duplicate_Func(QDialog):
    def __init__(self, main_window, header):
        super().__init__()
        self.setWindowTitle('중복 제거 팝업')
        self.main_window = main_window

        # 위젯 추가
        layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        layout4 = QHBoxLayout()
        self.label1 = QLabel('기준 Column')
        self.cmb1 = QComboBox()
        self.label2 = QLabel('중복 제거 후 남길 Row : ')
        self.radio_btn1 = QRadioButton('맨 위', self)
        self.radio_btn2 = QRadioButton('맨 아래', self)
        self.accept_btn = QPushButton('중복 제거')
        self.exit_dialog_btn = QPushButton('취소')

        # 레이아웃 지정
        layout.addLayout(layout2)
        layout2.addWidget(self.label1)
        layout2.addWidget(self.cmb1)

        layout.addLayout(layout3)
        layout3.addWidget(self.label2)
        layout3.addWidget(self.radio_btn1)
        layout3.addWidget(self.radio_btn2)

        layout.addLayout(layout4)
        layout4.addWidget(self.accept_btn)
        layout4.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)

        # 시그널 추가
        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

        # 변수 초기화
        self.header = header
        self.cmb1.addItems(self.header)
        self.selected_combo_item = None
        self.selected_radio_button = None

    def accept_func(self):
        # 다이얼 로그 값 전달
        self.selected_combo_item = self.cmb1.currentText()
        if self.radio_btn1.isChecked():
            self.selected_radio_button = 'first'
        elif self.radio_btn2.isChecked():
            self.selected_radio_button = 'last'
        else:
            QMessageBox.warning(self, '경고', '중복 제거 후 남길 Row 기준을 선택해 주세요.')
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

        # 위젯 추가
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

        # 레이아웃 지정
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

        # 시그널 추가
        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

        # 변수 초기화
        self.selected_combo_index = None
        self.selected_radio_button = None
        self.inserted_header_val = None
        self.inserted_data_val = None
        self.cmb1.addItems(header)

    def accept_func(self):
        # 다이얼 로그 값 전달
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

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.reject()


class insert_row_Func(QDialog):
    def __init__(self, main_window, header, rows):
        super().__init__()
        self.setWindowTitle('데이터 삽입')
        self.main_window = main_window
        self.header = header
        self.rows = rows
        self.setGeometry(850, 50, 800, self.height())

        # 위젯 추가
        layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        layout4 = QHBoxLayout()
        self.label1 = QLabel('기준 Row : ')
        self.lineedit1 = QLineEdit()
        self.label2 = QLabel('삽입 위치 지정 : ')
        self.radio_btn1 = QRadioButton('기준 Row 위', self)
        self.radio_btn2 = QRadioButton('기준 Row 아래', self)
        self.row_add_btn = QPushButton('행 추가')
        self.table_widget = QTableWidget()
        self.accept_btn = QPushButton('삽입')
        self.exit_dialog_btn = QPushButton('취소')

        # 레이아웃 지정
        layout.addLayout(layout2)
        layout2.addWidget(self.label1)
        layout2.addWidget(self.lineedit1)

        layout.addLayout(layout3)
        layout2.addWidget(self.label2)
        layout2.addWidget(self.radio_btn1)
        layout2.addWidget(self.radio_btn2)

        layout2.addWidget(self.row_add_btn)
        layout.addWidget(self.table_widget)

        layout.addLayout(layout4)
        layout4.addWidget(self.accept_btn)
        layout4.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)

        # lineedit 입력 값 제한
        self.lineedit1.setValidator(QIntValidator())

        # 시그널 추가
        self.row_add_btn.clicked.connect(self.add_table_row)
        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

        # 변수 초기화
        self.selected_row_index = None
        self.selected_radio_button = None
        self.insert_list = []
        self.set_table()

    def set_table(self):
        # row 추가용 임시 테이블 초기화
        new_header = self.header.copy()
        new_header.insert(0, '삭제')
        self.table_widget.setColumnCount(len(new_header))
        self.table_widget.setHorizontalHeaderLabels(new_header)
        self.table_widget.resizeColumnsToContents()

    def add_table_row(self):
        # 임시 테이블 row 추가
        del_btn = QPushButton('삭제')
        del_btn.clicked.connect(self.del_table_row)
        row_index = self.table_widget.rowCount()
        self.table_widget.insertRow(row_index)
        self.table_widget.setCellWidget(row_index, 0, del_btn)

    def del_table_row(self):
        # 임시 테이블 row 삭제
        sender = self.sender()
        if isinstance(sender, QPushButton):
            index = self.table_widget.indexAt(sender.pos())
            self.table_widget.removeRow(index.row())

    def accept_func(self):
        # 다이얼 로그 값 전달
        self.insert_list = []
        row_index = int(self.lineedit1.text())
        if self.table_widget.rowCount() > 0:
            if row_index and 0 < row_index <= self.rows:
                self.selected_row_index = self.lineedit1.text()
            else:
                QMessageBox.warning(self, '경고', f'데이터 삽입 기준 Row 값을 확인해 주세요. 입력 가능 범위 : 1~{self.rows}')
                return
            if self.radio_btn1.isChecked():
                self.selected_radio_button = 0
            elif self.radio_btn2.isChecked():
                self.selected_radio_button = 1
            else:
                QMessageBox.warning(self, '경고', '데이터 삽입 위치를 선택해 주세요.')
                return
            rows = self.table_widget.rowCount()
            cols = self.table_widget.columnCount()-1
            for row in range(rows):
                insert_row = []
                for col in range(cols):
                    item = self.table_widget.item(row, col+1)
                    if item is not None:
                        insert_row.append(item.text())
                    else:
                        insert_row.append("")
                self.insert_list.append(insert_row)
            self.accept()

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.reject()


class delete_col_Func(QDialog):
    def __init__(self, main_window, header):
        super().__init__()
        self.setGeometry(850, 50, 300, 100)
        self.setWindowTitle('데이터 삭제')
        self.main_window = main_window
        self.header = header

        # 위젯 추가
        layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        self.label1 = QLabel('<b>시작 기준 Column : </b>')
        self.combobox1 = QComboBox()
        self.label2 = QLabel('<b>종료 기준 Column : </b>')
        self.combobox2 = QComboBox()
        self.accept_btn = QPushButton('삭제')
        self.exit_dialog_btn = QPushButton('취소')

        # 레이아웃 지정
        layout.addLayout(layout2)
        layout2.addWidget(self.label1)
        layout2.addWidget(self.combobox1)
        layout2.addWidget(self.label2)
        layout2.addWidget(self.combobox2)

        layout.addLayout(layout3)
        layout3.addWidget(self.accept_btn)
        layout3.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)

        # 시그널 추가
        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

        self.add_comboboxes_items()

    def add_comboboxes_items(self):
        self.combobox1.addItems(self.header)
        self.combobox2.addItems(self.header)

    def accept_func(self):
        # 다이얼 로그 값 전달
        first_col_index = self.combobox1.currentIndex()
        last_col_index = self.combobox2.currentIndex()
        if first_col_index > len(self.header) or last_col_index > len(self.header):
            QMessageBox.warning(self, '경고', f'데이터 삽입 기준 Column 값을 확인해 주세요. '
                                            f'입력 가능 범위 : 0~{len(self.header)}')
            return
        elif first_col_index - last_col_index > 0:
            QMessageBox.warning(self, '경고', f'데이터 삽입 기준 Column 값을 확인해 주세요. '
                                            f'시작 기준 칼럼은 종료 기준 칼럼의 앞에 위치 해야 합니다.')
            return
        self.accept()
        self.main_window.do_delete_col(first_col_index, last_col_index)

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.reject()


class delete_row_Func(QDialog):
    def __init__(self, main_window, rows):
        super().__init__()
        self.setGeometry(850, 50, 300, 100)
        self.setWindowTitle('데이터 삭제')
        self.main_window = main_window
        self.rows = rows

        # 위젯 추가
        layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        self.label1 = QLabel('<b>시작 기준 Row : </b>')
        self.lineedit1 = QLineEdit()
        self.label2 = QLabel('<b>종료 기준 Row : </b>')
        self.lineedit2 = QLineEdit()
        self.accept_btn = QPushButton('삭제')
        self.exit_dialog_btn = QPushButton('취소')

        # 레이아웃 지정
        layout.addLayout(layout2)
        layout2.addWidget(self.label1)
        layout2.addWidget(self.lineedit1)
        layout2.addWidget(self.label2)
        layout2.addWidget(self.lineedit2)

        layout.addLayout(layout3)
        layout3.addWidget(self.accept_btn)
        layout3.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)

        # 시그널 추가
        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

    def accept_func(self):
        # 다이얼 로그 값 전달
        if self.lineedit1.text() and self.lineedit2.text():
            first_row_index = int(self.lineedit1.text())-1
            last_row_index = int(self.lineedit2.text())-1
            if 0 <= first_row_index <= self.rows:
                if 0 <= last_row_index <= self.rows:
                    self.accept()
                    self.main_window.do_delete_row(first_row_index, last_row_index)
                else:
                    QMessageBox.warning(self, '경고', f'종료 기준 Row의 데이터 값을 확인해 주세요. '
                                                    f'입력 가능 범위 : 1~{self.rows}')
                    return
            else:
                QMessageBox.warning(self, '경고', f'시작 기준 Row의 데이터 값을 확인해 주세요. '
                                                f'입력 가능 범위 : 1~{self.rows}')
                return

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.reject()


class show_listwidget(QDialog):
    def __init__(self, main_window):
        super().__init__(main_window, flags=Qt.Window)
        self.setGeometry(850,50,300,600)
        self.setWindowTitle('업로드 파일 목록')
        self.main_window = main_window

        layout = QVBoxLayout()
        self.list_widget = QListWidget()
        self.exit_dialog_btn = QPushButton('종료')

        layout.addWidget(self.list_widget)
        layout.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)

        self.list_widget.doubleClicked.connect(self.change_table)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

    def change_table(self):
        # 리스트 위젯 더블 클릭 시 실행 함수
        file_path = self.list_widget.currentItem()
        self.main_window.list_widget_exec(file_path.text())

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.close()


class replace_Func(QDialog):
    def __init__(self, main_window):
        super().__init__(main_window, flags=Qt.Window)
        self.setGeometry(850, 50, self.width(), self.height())
        self.setWindowTitle('문자열 변환')
        self.main_window = main_window

        # 위젯 추가
        layout = QVBoxLayout()
        glayout = QGridLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        layout4 = QHBoxLayout()
        layout5 = QHBoxLayout()

        self.label1 = QLabel('기준 Column : ')
        self.lineedit1 = QLineEdit()
        self.label2 = QLabel(' ~ ')
        self.lineedit2 = QLineEdit()
        self.label3 = QLabel('전체 Column : ')
        self.checkbox1 = QCheckBox()
        self.label4 = QLabel('기준 Row : ')
        self.lineedit3 = QLineEdit()
        self.label5 = QLabel(' ~ ')
        self.lineedit4 = QLineEdit()
        self.label6 = QLabel('전체 Row : ')
        self.checkbox2 = QCheckBox()
        self.label7 = QLabel('찾을 문구 : ')
        self.lineedit5 = QLineEdit()
        self.label8 = QLabel('변환할 문구 : ')
        self.lineedit6 = QLineEdit()
        self.label9 = QLabel('정확히 일치 여부 : ')
        self.checkbox3 = QCheckBox()
        self.accept_btn = QPushButton('변환')
        self.exit_dialog_btn = QPushButton('취소')

        # 레이아웃 지정
        layout.addLayout(glayout)
        glayout.addWidget(self.label1, 0, 0)
        glayout.addWidget(self.lineedit1, 0, 1)
        glayout.addWidget(self.label2, 0, 2)
        glayout.addWidget(self.lineedit2, 0, 3)
        glayout.addWidget(self.label3, 0, 4)
        glayout.addWidget(self.checkbox1, 0, 5)
        glayout.addWidget(self.label4, 1, 0)
        glayout.addWidget(self.lineedit3, 1, 1)
        glayout.addWidget(self.label5, 1, 2)
        glayout.addWidget(self.lineedit4, 1, 3)
        glayout.addWidget(self.label6, 1, 4)
        glayout.addWidget(self.checkbox2, 1, 5)

        layout.addLayout(layout2)
        layout2.addWidget(self.label7)
        layout2.addWidget(self.lineedit5)

        layout.addLayout(layout3)
        layout3.addWidget(self.label8)
        layout3.addWidget(self.lineedit6)

        layout.addLayout(layout4)
        layout4.addWidget(self.label9)
        layout4.addWidget(self.checkbox3)
        layout4.addStretch(1)

        layout.addLayout(layout5)
        layout5.addWidget(self.accept_btn)
        layout5.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)

        # 시그널 추가
        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)
        self.checkbox1.stateChanged.connect(self.checkbox1_signal)
        self.checkbox2.stateChanged.connect(self.checkbox2_signal)

        self.checkbox1.setChecked(True)
        self.checkbox2.setChecked(True)
        self.lineedit1.setValidator(QIntValidator())
        self.lineedit2.setValidator(QIntValidator())
        self.lineedit3.setValidator(QIntValidator())
        self.lineedit4.setValidator(QIntValidator())
        self.label2.setAlignment(Qt.AlignCenter)
        self.label5.setAlignment(Qt.AlignCenter)

    def accept_func(self):
        # 다이얼 로그 값 전달
        if self.lineedit1.text() and self.lineedit2.text() and \
                self.lineedit3.text() and self.lineedit4.text():
            first_col_index = int(self.lineedit1.text())
            last_col_index = int(self.lineedit2.text())
            first_row_index = int(self.lineedit3.text())
            last_row_index = int(self.lineedit4.text())
            if not (0 <= first_col_index <= len(self.main_window.header)) or \
                    not (0 <= last_col_index <= len(self.main_window.header)) or \
                    not (0 <= first_row_index <= self.main_window.rows) or \
                    not (0 <= last_row_index <= self.main_window.rows):
                QMessageBox.warning(self, '경고', '범위 밖의 데이터가 입력 되었습니다.')
                return
            find_text = self.lineedit5.text()
            replace_text = self.lineedit6.text()
            if self.checkbox3.isChecked():
                checkbox_value = True
            else:
                checkbox_value = False
            self.main_window.do_replace(first_col_index-1, last_col_index-1, first_row_index-1, last_row_index-1,
                                        find_text, replace_text, checkbox_value)

    def checkbox1_signal(self, state):
        if state == 2:
            self.lineedit1.setText("1")
            self.lineedit2.setText(str(len(self.main_window.header)))
            self.lineedit1.setReadOnly(True)
            self.lineedit2.setReadOnly(True)
        elif state == 0:
            self.lineedit1.setReadOnly(False)
            self.lineedit2.setReadOnly(False)

    def checkbox2_signal(self, state):
        if state == 2:
            self.lineedit3.setText("1")
            self.lineedit4.setText(str(self.main_window.rows))
            self.lineedit3.setReadOnly(True)
            self.lineedit4.setReadOnly(True)
        elif state == 0:
            self.lineedit3.setReadOnly(False)
            self.lineedit4.setReadOnly(False)

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.close()


class text_filter_Func(QDialog):
    def __init__(self, main_window, header):
        super().__init__()
        self.setGeometry(850, 50, self.width(), 200)
        self.setWindowTitle('텍스트 필터링')
        self.main_window = main_window
        self.header = header

        # 위젯 추가
        layout = QVBoxLayout()
        layout2 = QHBoxLayout()
        layout3 = QHBoxLayout()
        layout4 = QHBoxLayout()
        layout5 = QHBoxLayout()
        self.label1 = QLabel('<b>기준 Column : </b>')
        self.combobox1 = QComboBox()
        self.label2 = QLabel('<b>필터링 텍스트 : </b>')
        self.lineedit1 = QLineEdit()
        self.label3 = QLabel('<b>정확히 일치 여부 : </b>')
        self.checkbox1 = QCheckBox()
        self.accept_btn = QPushButton('적용')
        self.exit_dialog_btn = QPushButton('취소')

        # 레이아웃 지정
        layout.addLayout(layout2)
        layout2.addWidget(self.label1)
        layout2.addWidget(self.combobox1)

        layout.addLayout(layout3)
        layout3.addWidget(self.label2)
        layout3.addWidget(self.lineedit1)

        layout.addLayout(layout4)
        layout4.addWidget(self.label3)
        layout4.addWidget(self.checkbox1)
        layout4.addStretch(1)

        layout.addLayout(layout5)
        layout5.addWidget(self.accept_btn)
        layout5.addWidget(self.exit_dialog_btn)
        self.setLayout(layout)

        # 시그널 추가
        self.accept_btn.clicked.connect(self.accept_func)
        self.exit_dialog_btn.clicked.connect(self.exit_dialog)

        self.add_comboboxes_items()

    def add_comboboxes_items(self):
        self.combobox1.addItems(self.header)

    def accept_func(self):
        # 다이얼 로그 값 전달
        criteria_col = self.combobox1.currentText()
        criteria_text = self.lineedit1.text()
        if self.checkbox1.isChecked():
            checkbox_value = True
        else:
            checkbox_value = False
        self.accept()
        self.main_window.do_text_filter(criteria_col, criteria_text, checkbox_value)

    def exit_dialog(self):
        # 다이얼 로그 종료
        self.reject()

