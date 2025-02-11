import sys
import csv
import os
import win32com.client
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout,
    QDateEdit, QMessageBox, QHBoxLayout, QTableWidget, QTableWidgetItem, QSplitter
)
from PyQt5.QtCore import QDate, Qt

class HwpProcessor:
    def __init__(self, template_path, output_path):
        self.template_path = template_path
        self.output_path = output_path
        self.hwp = self._initialize_hwp()

    def _initialize_hwp(self):
        try:
            hwp = win32com.client.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
            return hwp
        except Exception as e:
            raise Exception(f"HWP 초기화 실패: {e}")

    def create_hwp_file(self, name):
        try:
            self.hwp.Open(self.template_path)
            self.hwp.HAction.Run("MoveTop")
            
            self.hwp.HParameterSet.HFindReplace.FindString = "%Name"
            self.hwp.HParameterSet.HFindReplace.ReplaceString = name
            self.hwp.HParameterSet.HFindReplace.ReplaceMode = 1
            self.hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
            self.hwp.HAction.Execute("AllReplace", self.hwp.HParameterSet.HFindReplace.HSet)
            
            output_file = os.path.join(self.output_path, f"출근부_{name}.hwp")
            self.hwp.SaveAs(output_file)
            self.hwp.Quit()
            return output_file
        except Exception as e:
            self.hwp.Quit()
            raise Exception(f"파일 생성 실패: {e}")

class InputForm(QWidget):
    def __init__(self):
        super().__init__()
        self.file_path = "user_data.txt"
        self.holiday_file_path = "holiday_data.txt"
        self.hwp_template = "C:/Users/pc/Desktop/project/청년이룸출근부.hwp"
        self.hwp_output_path = "C:/Users/pc/Desktop/project/"
        self.initUI()

    def initUI(self):
        self.setGeometry(100, 100, 800, 600)
        
        main_layout = QHBoxLayout()
        splitter = QSplitter(Qt.Horizontal)
        left_splitter = QSplitter(Qt.Vertical)
        right_splitter = QSplitter(Qt.Vertical)

        form_layout = QVBoxLayout()
        self.name_label = QLabel('이름:')
        self.name_input = QLineEdit(self)
        form_layout.addWidget(self.name_label)
        form_layout.addWidget(self.name_input)

        self.submit_button = QPushButton('입력')
        self.submit_button.clicked.connect(self.save_data)
        form_layout.addWidget(self.submit_button)

        self.generate_hwp_button = QPushButton('HWP 파일 생성')
        self.generate_hwp_button.clicked.connect(self.generate_hwp)
        form_layout.addWidget(self.generate_hwp_button)

        form_widget = QWidget()
        form_widget.setLayout(form_layout)
        left_splitter.addWidget(form_widget)
        splitter.addWidget(left_splitter)

        main_layout.addWidget(splitter)
        self.setLayout(main_layout)
        self.setWindowTitle('사용자 입력 및 HWP 생성')

    def generate_hwp(self):
        name = self.name_input.text()
        if not name:
            QMessageBox.warning(self, "입력 오류", "이름을 입력하세요!")
            return
        
        try:
            processor = HwpProcessor(self.hwp_template, self.hwp_output_path)
            output_file = processor.create_hwp_file(name)
            QMessageBox.information(self, "완료", f"HWP 파일 생성 완료: {output_file}")
        except Exception as e:
            QMessageBox.critical(self, "오류", str(e))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    form = InputForm()
    form.show()
    sys.exit(app.exec_())