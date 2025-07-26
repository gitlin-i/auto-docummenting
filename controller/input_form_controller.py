import csv
import os
from PyQt5.QtCore import QDate
from PyQt5.QtWidgets import QMessageBox ,QTableWidgetItem
from model.input_form_model import InputFormModel
from model.eroom import EroomManagerSchedule, MetaData, PublicHoliday
from controller.hwp_controller import modify_hwp_file
from view.input_form_view import InputFormView
class InputFormController:
    def __init__(self, view : InputFormView ):
        self.view = view
        self.model = InputFormModel()
        self.connect_signals()

    def connect_signals(self):
        self.view.submit_button.clicked.connect(self.save_data)
        self.view.delete_button.clicked.connect(self.delete_data)
        self.view.print_button.clicked.connect(self.print_to_hwp)
        self.view.next_month_button.clicked.connect(self.move_all_to_next_month)
        self.view.prev_month_button.clicked.connect(self.move_all_to_prev_month)
        self.view.holiday_submit_button.clicked.connect(self.save_holiday)
        self.view.holiday_delete_button.clicked.connect(self.delete_holiday)
        self.view.table.itemSelectionChanged.connect(self.fill_form_from_selection)

    def fill_form_from_selection(self):
        selected_row = self.view.table.currentRow()
        if selected_row == -1:
            return
        self.view.name_input.setText(self.view.table.item(selected_row, 0).text())
        self.view.alternative_input.setDate(QDate.fromString(self.view.table.item(selected_row, 1).text(), "yyyy-MM-dd"))
        self.view.saturday_input.setDate(QDate.fromString(self.view.table.item(selected_row, 2).text(), "yyyy-MM-dd"))

    def save_data(self):
        name = self.view.name_input.text()
        alternative_leave = self.view.alternative_input.date().toString("yyyy-MM-dd")
        saturday_work = self.view.saturday_input.date().toString("yyyy-MM-dd")

        if not name:
            QMessageBox.warning(self.view, "입력 오류", "이름을 입력하세요!")
            return

        self.model.save_data(name, alternative_leave, saturday_work)
        QMessageBox.information(self.view, "저장 완료", "데이터가 성공적으로 저장되었습니다.")
        self.load_data()

    def load_data(self):
        data = self.model.load_data()
        self.view.table.setRowCount(len(data))
        for row_idx, row_data in enumerate(data):
            for col_idx, col_data in enumerate(row_data):
                self.view.table.setItem(row_idx, col_idx, QTableWidgetItem(col_data))

    def save_holiday(self):
        holiday_date = self.view.holiday_input.date().toString("yyyy-MM-dd")
        self.model.save_holiday(holiday_date)
        self.load_holiday_data()

    def load_holiday_data(self):
        holidays = self.model.load_holiday_data()
        self.view.holiday_table.setRowCount(len(holidays))
        for i, date in enumerate(holidays):
            self.view.holiday_table.setItem(i, 0, QTableWidgetItem(date))

    def delete_data(self):
        selected_row = self.view.table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self.view, "삭제 오류", "삭제할 행을 선택하세요!")
            return

        name = self.view.table.item(selected_row, 0).text()
        self.model.delete_data(name)
        QMessageBox.information(self.view, "삭제 완료", "데이터가 성공적으로 삭제되었습니다.")
        self.load_data()

    def move_all_to_next_month(self):
        self.model.move_all_to_next_month()
        self.load_data()

    def move_all_to_prev_month(self):
        self.model.move_all_to_prev_month()
        self.load_data()

    def delete_holiday(self):
        selected_row = self.view.holiday_table.currentRow()
        if selected_row == -1:
            QMessageBox.warning(self.view, "삭제 오류", "삭제할 공휴일을 선택하세요!")
            return

        holiday_date = self.view.holiday_table.item(selected_row, 0).text()
        self.model.delete_holiday(holiday_date)
        QMessageBox.information(self.view, "삭제 완료", "공휴일이 성공적으로 삭제되었습니다.")
        self.load_holiday_data()

    def print_to_hwp(self):
        try:
            current_year = self.view.year_combo.currentData()
            current_month = self.view.month_combo.currentData()
            holiday_days = self.model.load_holiday_data()
            if not os.path.exists(self.model.file_path):
                QMessageBox.warning(self.view, "오류", "출력할 데이터가 없습니다.")
                return
            
            with open(self.model.file_path, mode='r', encoding='utf-8') as file:
                reader = csv.DictReader(file)
                data = list(reader)
            
            for row in data:
                ems = EroomManagerSchedule(row["이름"], row["대체 휴무 날짜"], row["토요일 근무 날짜"])
                
                meta_data = MetaData(
                    default_file_path=os.getcwd(),
                    input_file="청년이룸출근부.hwp",
                    output_file_name=f"청년이룸출근부_{current_year}년_{current_month}월_{ems.name}.hwp",
                    target_date=f"{current_year}-{str(current_month).zfill(2)}"
                )
                
                modify_hwp_file(meta_data, ems, holiday_days)
            
            QMessageBox.information(self.view, "출력 완료", "한글 파일 출력이 완료되었습니다.")
        except Exception as e:
            QMessageBox.warning(self.view, "오류", f"오류 발생: {str(e)}")