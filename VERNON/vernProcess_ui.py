import os
import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout, QWidget, QFileDialog
import pandas as pd
import time



class FileCaptureWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("File Capture")

        self.file_picker_button1 = QPushButton("Select Book file (FP REC) ")
        self.file_picker_button2 = QPushButton("Select Bank DLR")
        self.capture_button = QPushButton("Process")
        self.file_path_label1 = QLabel("")
        self.file_path_label2 = QLabel("")

        self.file_picker_button1.clicked.connect(self.select_file1)
        self.file_picker_button2.clicked.connect(self.select_file2)
        self.capture_button.clicked.connect(self.capture_files)

        central_widget = QWidget()
        layout = QVBoxLayout(central_widget)
        layout.addWidget(self.file_picker_button1)
        layout.addWidget(self.file_picker_button2)
        layout.addWidget(self.capture_button)
        layout.addWidget(self.file_path_label1)
        layout.addWidget(self.file_path_label2)

        self.setCentralWidget(central_widget)

    def select_file1(self):
        current_dir = os.path.dirname(os.path.realpath(__file__))
        file_dialog = QFileDialog()
        file_dialog.setDirectory(current_dir)
        file_path, _ = file_dialog.getOpenFileName(self, "Select book file (FP REC) ")
        if file_path:
            self.file_path_label1.setText(file_path)

    def select_file2(self):
        current_dir = os.path.dirname(os.path.realpath(__file__))
        file_dialog = QFileDialog()
        file_dialog.setDirectory(current_dir)
        file_path, _ = file_dialog.getOpenFileName(self, "Select bank DLR")
        if file_path:
            self.file_path_label2.setText(file_path)

    def capture_files(self):
        file_path1 = self.file_path_label1.text()
        file_path2 = self.file_path_label2.text()
        if file_path1 and file_path2:
            # Perform capture logic here
            #Books
            fl_plan = pd.read_excel(file_path1)

            #DLR
            timestr = time.strftime("%Y_%m_%d")

            bank_account = pd.read_excel(file_path2,sheet_name='sheet1')
            file_name = r'C:\Users\lsend\dev\py\sam\VERNON\vernon_rec_{}.xlsx'.format(timestr)

            writer = pd.ExcelWriter(file_name, engine='xlsxwriter')

            bank_account = bank_account.rename(
                columns={'\nFull Serial Number': 'VIN'})


            bank_account = bank_account.rename(columns=lambda x: x.strip())
            fl_plan =fl_plan.rename(columns=lambda x: x.strip())
            fl_plan.dropna(subset=['Control'], inplace=True)

            bank_account['VIN'] = bank_account['VIN'].str.replace(' ', '')

            bank_account = bank_account[~bank_account['VIN'].astype(
                str).str.startswith('total')]

            bank_account = bank_account[~bank_account['VIN'].astype(
                str).str.startswith('Customer')]

            bank_account = bank_account[~bank_account['VIN'].astype(
                str).str.startswith('Allamount')]
            bank_account = bank_account[~bank_account['VIN'].astype(
                str).str.startswith('Total')]


            if 'Description' not in fl_plan.columns:
                print("Error: No Description in Floor Plan Report")


            fl_plan['VIN'] = fl_plan['Description'].str[-17:]

            if 'VIN' not in fl_plan.columns:
                print("Error: 'vin' column not found in report")
                exit()
            if 'VIN' not in bank_account.columns:
                print("Error: 'vin' column not found in bank_account")
                exit()
            # Merge the two dataframes based on the vin number

            print("bank:")
            print(bank_account.columns.tolist())
            print("fp:")
            print(fl_plan.columns.tolist())

            #bank_account = bank_account.drop(0)

            merged = pd.merge(fl_plan, bank_account,  right_on='VIN',
                            left_on='VIN', how='outer', indicator=True)

            in_file1 = merged['_merge'] == 'left_only'
            in_file2 = merged['_merge'] == 'right_only'
            matching = merged['_merge'] == 'both'

            df1 = merged[in_file1].drop('_merge', axis=1)
            df2 = merged[in_file2].drop('_merge', axis=1)
            df2 = df2[df2['Current Principal'] != 0]

            print("\nMatching items in both report and bank_account:")
            print(merged[matching].drop('_merge', axis=1))

            df3 = merged[matching].drop('_merge', axis=1)

            column_add= pd.Series([float('nan')] * len(df3))
            column_add_bool= pd.Series([bool()] * len(df3))



            df3.insert(df3.columns.get_loc('VIN') + 1, 'HAS_DIFF', column_add_bool)
            df3.insert(df3.columns.get_loc('VIN') + 1, 'DIFFER', column_add)
            df3.insert(df3.columns.get_loc('VIN') + 1, 'BOOKS_T_AMOUNT', column_add)


            columns_to_sum = ['23100','2G3100','23120','33100','33120','1310','1311']

            existing_columns = [col for col in columns_to_sum if col in df3.columns]
            df3['BOOKS_T_AMOUNT'] = df3[existing_columns].sum(axis=1)
            df3['DIFFER'] = df3['BOOKS_T_AMOUNT'] + df3['Current Principal']
            df3['HAS_DIFF'] = df3['DIFFER'] != 0

            df3['HAS_DIFF'] = df3['HAS_DIFF'].astype(bool)
            recon_df = df3[df3['HAS_DIFF'] == True].copy()

            recon_df.to_excel(writer, sheet_name='Recon')
            df1.to_excel(writer, sheet_name='On Books')
            df2.to_excel(writer, sheet_name='Key Bank')
            df3.to_excel(writer, sheet_name='In both')
            writer.save()
        else:    
            print("Please Add your 2 Files ")
app = QApplication(sys.argv)
window = FileCaptureWindow()
window.show()

sys.exit(app.exec())
