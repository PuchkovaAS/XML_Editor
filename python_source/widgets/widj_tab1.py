from PyQt5.QtCore import pyqtSignal
from PyQt5.QtWidgets import QFileDialog, QPushButton, QLabel, QGridLayout, QWidget


class line_btn_xml_path(QWidget):
    my_signal = pyqtSignal('PyQt_PyObject')
    text_label = 'XML path:'

    def __init__(self, mainwin):
        super().__init__()
        self.mainwin = mainwin
        layout: QGridLayout = QGridLayout()
        self.btn = QPushButton("...")
        self.line_text = QLabel(self)
        self.line_text.setText('')

        self.btn.clicked.connect(self.open_file)

        self.btn.setFixedSize(30, 30)

        self.line_text.setText('')
        self.line_text.setMinimumWidth(150)
        layout.addWidget(QLabel(self.text_label), 0, 0, 1, 3)
        layout.addWidget(self.line_text, 1, 0, 1, 3)
        layout.addWidget(self.btn, 1, 3, 1, 1)
        self.setLayout(layout)

    def open_file(self):
        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        path, _ = QFileDialog.getOpenFileName(self, 'Открыть файл', "", "xml (*.xml)", options=options)

        if path == '':
            return
        self.line_text.setText(path)
        self.my_signal.emit(path)


class line_btn_xls_path(line_btn_xml_path):
    text_label = 'XLSX path:'

    def open_file(self):
        options = QFileDialog.Options()
        # options |= QFileDialog.DontUseNativeDialog
        path, _ = QFileDialog.getOpenFileName(self, 'Открыть файл', "", "xlsx (*.xlsx)", options=options)

        if path == '':
            return
        self.line_text.setText(path)
        self.my_signal.emit(path)
