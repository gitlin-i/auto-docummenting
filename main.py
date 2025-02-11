import sys
from input_form import InputForm
from PyQt5.QtWidgets import (
    QApplication
)

app = QApplication(sys.argv)
form = InputForm()
form.show()
sys.exit(app.exec_())