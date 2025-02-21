import sys
import csv
import os
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QDateEdit, QMessageBox, QHBoxLayout, QTableWidget, QTableWidgetItem, QSplitter, QComboBox
)
from PyQt5.QtCore import QDate, Qt

from eroom import EroomManagerSchedule, MetaData
from write_hwp import modify_hwp_file

class InputForm(QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = "user_data.txt"
        self.holiday_file_path = "holiday_data.txt"
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
        self.next_month_button.clicked.connect(self.move_all_to_next_month)
        form_layout.addWidget(self.next_month_button)
        # 이전달 버튼 추가
        self.prev_month_button = QPushButton('이전달')
        self.prev_month_button.clicked.connect(self.move_all_to_prev_month)
        form_layout.addWidget(self.prev_month_button)

        self.submit_button = QPushButton('입력')
        self.submit_button.clicked.connect(self.save_data)
        form_layout.addWidget(self.submit_button)

        self.delete_button = QPushButton('삭제')
        self.delete_button.clicked.connect(self.delete_data)
        form_layout.addWidget(self.delete_button)

        self.print_button = QPushButton("한글 파일 출력")
        self.print_button.clicked.connect(self.print_to_hwp)
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
        self.holiday_submit_button.clicked.connect(self.save_holiday)
        holiday_layout.addWidget(self.holiday_submit_button)

        self.holiday_delete_button = QPushButton('공휴일 삭제')
        self.holiday_delete_button.clicked.connect(self.delete_holiday)
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
        self.table.itemSelectionChanged.connect(self.fill_form_from_selection)
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

        self.load_data()
        self.load_holiday_data()

    def fill_form_from_selection(self):
        selected_row = self.table.currentRow()
        if selected_row == -1:
            return
        self.name_input.setText(self.table.item(selected_row, 0).text())
        self.alternative_input.setDate(QDate.fromString(self.table.item(selected_row, 1).text(), "yyyy-MM-dd"))
        self.saturday_input.setDate(QDate.fromString(self.table.item(selected_row, 2).text(), "yyyy-MM-dd"))

    def save_data(self):
        name = self.name_input.text()
        alternative_leave = self.alternative_input.date().toString("yyyy-MM-dd")
        saturday_work = self.saturday_input.date().toString("yyyy-MM-dd")

        if not name:
            QMessageBox.warning(self, "입력 오류", "이름을 입력하세요!")
            return

        updated_data = {}
        if os.path.exists(self.file_path):
            with open(self.file_path, mode='r', encoding='utf-8') as file:
                reader = csv.reader(file)
                headers = next(reader, None)
                for row in reader:
                    updated_data[row[0]] = row[1:]

        updated_data[name] = [alternative_leave, saturday_work]
        
        with open(self.file_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["이름", "대체 휴무 날짜", "토요일 근무 날짜"])
            for key, value in updated_data.items():
                writer.writerow([key] + value)

        QMessageBox.information(self, "저장 완료", "데이터가 성공적으로 저장되었습니다.")
        self.load_data()

    def load_data(self):
        if not os.path.exists(self.file_path):
            return
        with open(self.file_path, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            next(reader, None)
            data = list(reader)
        self.table.setRowCount(len(data))
        for row_idx, row_data in enumerate(data):
            for col_idx, col_data in enumerate(row_data):
                self.table.setItem(row_idx, col_idx, QTableWidgetItem(col_data))

    def save_holiday(self):
        holiday_date = self.holiday_input.date().toString("yyyy-MM-dd")
        with open(self.holiday_file_path, mode='a', newline='', encoding='utf-8') as file:
            file.write(holiday_date + "\n")
        self.load_holiday_data()

    def load_holiday_data(self):
        if not os.path.exists(self.holiday_file_path):
            return
        with open(self.holiday_file_path, mode='r', encoding='utf-8') as file:
            holidays = file.readlines()
        self.holiday_table.setRowCount(len(holidays))
        for i, date in enumerate(holidays):
            self.holiday_table.setItem(i, 0, QTableWidgetItem(date.strip()))
    def delete_data(self):
        selected_row = self.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "삭제 오류", "삭제할 행을 선택하세요!")
            return

        self.table.removeRow(selected_row)

        # 데이터 파일에서 해당 행 삭제
        updated_data = []
        for row in range(self.table.rowCount()):
            updated_data.append([
                self.table.item(row, 0).text(),
                self.table.item(row, 1).text(),
                self.table.item(row, 2).text()
            ])

        with open(self.file_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(["이름", "대체 휴무 날짜", "토요일 근무 날짜"])
            writer.writerows(updated_data)

        QMessageBox.information(self, "삭제 완료", "데이터가 성공적으로 삭제되었습니다.")
        self.load_data()

    def move_all_to_next_month(self):
        """ CSV 파일의 모든 데이터를 4주(28일) 후로 이동 """
        if not os.path.exists(self.file_path):
            return

        updated_data = []
        with open(self.file_path, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            headers = next(reader, None)
            for row in reader:
                row[1] = QDate.fromString(row[1], "yyyy-MM-dd").addDays(28).toString("yyyy-MM-dd")
                row[2] = QDate.fromString(row[2], "yyyy-MM-dd").addDays(28).toString("yyyy-MM-dd")
                updated_data.append(row)

        with open(self.file_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(headers)
            writer.writerows(updated_data)

        self.load_data()

    def move_all_to_prev_month(self):
        """ CSV 파일의 모든 데이터를 4주(28일) 전으로 이동 """
        if not os.path.exists(self.file_path):
            return

        updated_data = []
        with open(self.file_path, mode='r', encoding='utf-8') as file:
            reader = csv.reader(file)
            headers = next(reader, None)
            for row in reader:
                row[1] = QDate.fromString(row[1], "yyyy-MM-dd").addDays(-28).toString("yyyy-MM-dd")
                row[2] = QDate.fromString(row[2], "yyyy-MM-dd").addDays(-28).toString("yyyy-MM-dd")
                updated_data.append(row)

        with open(self.file_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(headers)
            writer.writerows(updated_data)

        self.load_data()

    def delete_holiday(self):
        selected_row = self.holiday_table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self, "삭제 오류", "삭제할 공휴일을 선택하세요!")
            return

        self.holiday_table.removeRow(selected_row)
        
        updated_holidays = []
        for row in range(self.holiday_table.rowCount()):
            updated_holidays.append(self.holiday_table.item(row, 0).text())
        
        with open(self.holiday_file_path, mode='w', encoding='utf-8') as file:
            file.writelines([date + "\n" for date in updated_holidays])
        
        QMessageBox.information(self, "삭제 완료", "공휴일이 성공적으로 삭제되었습니다.")
        self.load_holiday_data()
    
    def print_to_hwp(self):
        try:
            current_year = self.year_combo.currentData()
            current_month = self.month_combo.currentData()
            
            if not os.path.exists(self.file_path):
                QMessageBox.warning(self, "오류", "출력할 데이터가 없습니다.")
                return
            
            with open(self.file_path, mode='r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                data = list(reader)
            
            for row in data:
                ems = EroomManagerSchedule(row["이름"],row["대체 휴무 날짜"],row["토요일 근무 날짜"])
                
                meta_data = MetaData(
                    default_file_path=os.getcwd(),
                    input_file="청년이룸출근부.hwp",
                    output_file_name=f"청년이룸출근부_{current_year}년_{current_month}월_{ems.name}.hwp",
                    target_date=f"{current_year}-{str(current_month).zfill(2)}"
                )
                
                modify_hwp_file(meta_data, ems)
            
            QMessageBox.information(self, "출력 완료", "한글 파일 출력이 완료되었습니다.")
        except Exception as e:
            QMessageBox.warning(self, "오류", f"오류 발생: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    form = InputForm()
    form.show()
    sys.exit(app.exec_())
