from PyQt5.QtWidgets import (
	QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog, QLabel, QPlainTextEdit, 
	QTableWidget, QTableWidgetItem)
import pandas as pd
import os


class StartWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        #self.setFixedSize(640, 480)
        self.central_widget = QWidget()
        self.button_folder_select = QPushButton('1. Select Folder...', self.central_widget)
        self.button_output_loc = QPushButton('2. Save As...', self.central_widget)
        self.button_convert = QPushButton('4. Convert', self.central_widget)
        self.button_column_labels = QPushButton('3. Column Labels...', self.central_widget)
        self.l1 = QLabel("Select Parent Directory.", self.central_widget)
        self.l2 = QLabel("Select output file name and location.", self.central_widget)
        #self.l2.setStyleSheet('color: red')
        self.l3 = QLabel("Enter column labels (don't forget units!).", self.central_widget)
        self.l4 = QLabel("Convert to XLSX file.", self.central_widget)
        # self.tableWidget = QTableWidget()
        # self.tableWidget.setRowCount(10)
        # self.tableWidget.setColumnCount(1)
        # self.tableWidget.setItem(0,0, QTableWidgetItem("Time (s)"))
        # self.tableWidget.setItem(1,0, QTableWidgetItem("Displacement (in)"))
        # self.tableWidget.setItem(2,0, QTableWidgetItem("Load (lbf)"))


        self.layout = QVBoxLayout(self.central_widget)
        self.layout.addWidget(self.l1)
        self.layout.addWidget(self.button_folder_select)
        self.layout.addWidget(self.l2)
        self.layout.addWidget(self.button_output_loc)
        self.layout.addWidget(self.l3)
        self.layout.addWidget(self.button_column_labels)
        #self.layout.addWidget(self.tableWidget)
        self.layout.addWidget(self.l4)
        self.layout.addWidget(self.button_convert)

        self.button_output_loc.setDisabled(True)
        self.button_convert.setDisabled(True)
        self.button_column_labels.setDisabled(True)

        self.setCentralWidget(self.central_widget)
        self.button_folder_select.clicked.connect(self.browse_folder)
        self.button_output_loc.clicked.connect(self.file_save_loc)
        self.button_convert.clicked.connect(self.convert_dat)
        self.button_column_labels.clicked(self.column_labels_textbox)

        self.dat_file_names = []  # file_names in scratch1

    def browse_folder(self):
        self.button_folder_select.folder_loc = QFileDialog.getExistingDirectory(self.central_widget,
                                                                                "Select Parent Folder")
        self.button_output_loc.setDisabled(False)
        self.button_folder_select.setDisabled(True)
       

    def file_save_loc(self):
        self.button_output_loc.saveXLSX, filter = QFileDialog.getSaveFileName(self, "Output File",  filter= '.xlsx')
        self.button_convert.setDisabled(False)
        self.button_output_loc.setDisabled(True)

    def column_labels_textbox(self):
    	self.button_column_labels = QPlainTextEdit(self)



#Save location works however only if .xlsx extension is explicitly stated.

    def convert_dat(self):
        i = 0
        w = pd.ExcelWriter(self.button_output_loc.saveXLSX)
        for root, dirs, files in os.walk(self.button_folder_select.folder_loc):
            for name in dirs:
                self.dat_file_names.append(name)

            for subdir in files:
                if subdir.endswith('.dat'):
                    with open(os.path.join(root, subdir)) as f1:
                        # change skiprow if needed
                        df = pd.read_csv(f1, sep='\s+', skiprows=16, error_bad_lines=False)
                        df = df.apply(pd.to_numeric, errors='coerce')
                        df = df.dropna()  # drop blank rows from coerce process
                        # manually change column names per test data acquisition
                        df.columns = ['Time (s)', 'Axial Displacement (in)', 'Axial Strain (in/in)',
                                      'Force (lbf)']
                        df.to_excel(w, sheet_name="%s" % self.dat_file_names[i], index=False)
                        i += 1
        w.save()
        
        self.button_convert.setStyleSheet("color: green")
        self.button_convert.setDisabled(True)
        self.button_convert = QPushButton('3. ', self.central_widget) #Doesn't change button text



if __name__ == '__main__':
    app = QApplication([])
    # app.setStyle('Windows')
    window = StartWindow()
    window.show()
    app.exit(app.exec_())

