from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QFileDialog
import pandas as pd
import os


class StartWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.central_widget = QWidget()
        self.button_folder_select = QPushButton('1. Select Folder...', self.central_widget)
        self.button_output_loc = QPushButton('2. Save As...', self.central_widget)
        self.button_convert = QPushButton('3. Convert', self.central_widget)

        self.layout = QVBoxLayout(self.central_widget)
        self.layout.addWidget(self.button_folder_select)
        self.layout.addWidget(self.button_output_loc)
        self.layout.addWidget(self.button_convert)

        self.button_output_loc.setDisabled(True)
        self.button_convert.setDisabled(True)

        self.setCentralWidget(self.central_widget)
        self.button_folder_select.clicked.connect(self.browse_folder)
        self.button_output_loc.clicked.connect(self.file_save_loc)
        self.button_convert.clicked.connect(self.convert_dat)

        self.dat_file_names = []  # file_names in scratch1

    def browse_folder(self):
        self.button_folder_select.folder_loc = QFileDialog.getExistingDirectory(self.central_widget,
                                                                                "Select Parent Folder")
        self.button_output_loc.setDisabled(False)
        print(self.button_folder_select.folder_loc)

    def file_save_loc(self):
        self.button_output_loc.saveXLSX, filter = QFileDialog.getSaveFileName(self, "Output File")
        self.button_convert.setDisabled(False)
        print(self.button_output_loc.saveXLSX)

    def convert_dat(self):
        i = 0
        w = pd.ExcelWriter('C:/Users/abh85/Desktop/Hilti Shearwall 2019/Hilti Shear Wall.xlsx')
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






if __name__ == '__main__':
    app = QApplication([])
    # app.setStyle('Windows')
    window = StartWindow()
    window.show()
    app.exit(app.exec_())

