from PyQt5.QtWidgets import (
    QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QDateEdit, QMessageBox, QHBoxLayout, QTableWidget, QTableWidgetItem, QSplitter, QComboBox
)
from PyQt5.QtCore import QDate, Qt

class InputFormView(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(100, 100, 800, 600)  # 창 크기 조정
        
        main_layout = QVBoxLayout()
        
        # 목표 달 선택 UI
        date_layout = QHBoxLayout()
        self.year_label = QLabel("몇 년 몇 월 출근부인가요?: ")
        self.year_combo = QComboBox()
        self.month_combo = QComboBox()
        
        current_year = QDate.currentDate().year()
        current_month = QDate.currentDate().month()
        
        for year in range(current_year - 5, current_year + 6):
            self.year_combo.addItem(f"{year}년", year)
        
        for month in range(1, 13):
            self.month_combo.addItem(f"{month}월", month)
        
        self.year_combo.setCurrentText(f"{current_year}년")
        self.month_combo.setCurrentText(f"{current_month}월")
        
        date_layout.addWidget(self.year_label)
        date_layout.addWidget(self.year_combo)
        date_layout.addWidget(self.month_combo)
        
        main_layout.addLayout(date_layout)
        
        splitter = QSplitter(Qt.Horizontal)
        left_splitter = QSplitter(Qt.Vertical)
        right_splitter = QSplitter(Qt.Vertical)

        form_layout = QVBoxLayout()
        self.name_label = QLabel('이름:')
        self.name_input = QLineEdit(self)
        form_layout.addWidget(self.name_label)
        form_layout.addWidget(self.name_input)

        self.alternative_label = QLabel('대체 휴무 날짜:')
        self.alternative_input = QDateEdit(self)
        self.alternative_input.setCalendarPopup(True)
        self.alternative_input.setDate(QDate.currentDate())
        form_layout.addWidget(self.alternative_label)
        form_layout.addWidget(self.alternative_input)

        self.saturday_label = QLabel('토요일 근무 날짜:')
        self.saturday_input = QDateEdit(self)
        self.saturday_input.setCalendarPopup(True)
        self.saturday_input.setDate(QDate.currentDate())
        form_layout.addWidget(self.saturday_label)
        form_layout.addWidget(self.saturday_input)
        # 다음달 버튼 추가
        self.next_month_button = QPushButton('다음달')
        form_layout.addWidget(self.next_month_button)
        # 이전달 버튼 추가
        self.prev_month_button = QPushButton('이전달')
        form_layout.addWidget(self.prev_month_button)

        self.submit_button = QPushButton('입력')
        form_layout.addWidget(self.submit_button)

        self.delete_button = QPushButton('삭제')
        form_layout.addWidget(self.delete_button)

        self.print_button = QPushButton("한글 파일 출력")
        form_layout.addWidget(self.print_button)

        form_widget = QWidget()
        form_widget.setLayout(form_layout)
        left_splitter.addWidget(form_widget)

        # 추가: 공휴일 입력 폼
        holiday_layout = QVBoxLayout()
        self.holiday_label = QLabel('공휴일 날짜:')
        self.holiday_input = QDateEdit(self)
        self.holiday_input.setCalendarPopup(True)
        self.holiday_input.setDate(QDate.currentDate())
        holiday_layout.addWidget(self.holiday_label)
        holiday_layout.addWidget(self.holiday_input)

        self.holiday_submit_button = QPushButton('공휴일 추가')
        holiday_layout.addWidget(self.holiday_submit_button)

        self.holiday_delete_button = QPushButton('공휴일 삭제')
        holiday_layout.addWidget(self.holiday_delete_button)

        holiday_widget = QWidget()
        holiday_widget.setLayout(holiday_layout)
        left_splitter.addWidget(holiday_widget)

        splitter.addWidget(left_splitter)

        # 오른쪽: 테이블 출력
        self.table = QTableWidget()
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["이름", "대체 휴무 날짜", "토요일 근무 날짜"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)  # 수정 불가 설정
        right_splitter.addWidget(self.table)

        # 공휴일 테이블 출력
        self.holiday_table = QTableWidget()
        self.holiday_table.setColumnCount(1)
        self.holiday_table.setHorizontalHeaderLabels(["공휴일 날짜"])
        self.holiday_table.setEditTriggers(QTableWidget.NoEditTriggers)  # 수정 불가 설정
        right_splitter.addWidget(self.holiday_table)

        splitter.addWidget(right_splitter)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)

        main_layout.addWidget(splitter)
        self.setLayout(main_layout)

        self.setWindowTitle('사용자 입력 및 데이터 보기')