import sys
from PySide2.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QComboBox, \
    QMainWindow, QMessageBox
from PySide2.QtGui import QPixmap
import qrcode
import pandas as pd
import cv2
import logging
import os
import glob

constant_position = 15
constant_size = 1
file_name = ''
image_path = ''
folder_path = ''
total_name_Written = []
log_file = '../dummy/gatepass.log'
if not os.path.exists(log_file):
    with open(log_file, 'w'):
        pass

# Configure the logging settings
logging.basicConfig(filename=log_file, level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')


def create_folder(name):
    for i in name:
        os.makedirs(i, exist_ok=True)


class AnotherWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Preview Window")
        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)
        layout = QVBoxLayout(self.central_widget)
        self.Excel_combo = QComboBox(self)
        df = pd.read_excel(file_name)
        value = df['MAC ID']
        self.Excel_combo.addItems(value)
        self.Excel_combo.currentIndexChanged.connect(self.Preview_details_button)
        layout.addWidget(self.Excel_combo)
        self.left_side_button = QPushButton("<", self)
        self.left_side_button.clicked.connect(self.on_left_button_clicked)
        layout.addWidget(self.left_side_button)
        self.right_side_button = QPushButton(">", self)
        self.right_side_button.clicked.connect(self.on_right_button_clicked)
        layout.addWidget(self.right_side_button)
        self.add_size_button = QPushButton("+", self)
        self.add_size_button.clicked.connect(self.on_add_button_clicked)
        layout.addWidget(self.add_size_button)
        self.minus_size_button = QPushButton("-", self)
        self.minus_size_button.clicked.connect(self.on_minus_button_clicked)
        layout.addWidget(self.minus_size_button)

        self.image_label = QLabel(self.central_widget)
        layout.addWidget(self.image_label)
        self.Save_button = QPushButton("Save")
        self.Save_button.clicked.connect(self.Save_window)
        layout.addWidget(self.Save_button)
        self.Cancel_button = QPushButton("Cancel")
        self.Cancel_button.clicked.connect(self.Cancel_window)
        layout.addWidget(self.Cancel_button)

    def Save_window(self):
        self.close()

    def Cancel_window(self):
        global constant_position, constant_size
        constant_position, constant_size = 20, 1
        if self.Excel_combo:
            self.Preview_details_button()

    def on_left_button_clicked(self):
        global constant_position
        constant_position -= 1
        if self.Excel_combo:
            self.Preview_details_button()

    def on_right_button_clicked(self):
        global constant_position
        constant_position += 1
        if self.Excel_combo:
            self.Preview_details_button()

    def on_minus_button_clicked(self):
        global constant_size
        constant_size -= 0.1
        if self.Excel_combo:
            self.Preview_details_button()

    def on_add_button_clicked(self):
        global constant_size
        constant_size += 0.1
        if self.Excel_combo:
            self.Preview_details_button()

    def Preview_details_button(self):
        try:
            pre_img_save = self.Excel_combo.currentText()
            pre_img_loc = './pre_img/pre_img_save.png'
            qr_img = qrcode.make(pre_img_save)
            qr_img.save(pre_img_loc)
            img = cv2.imread(pre_img_loc)
            font = cv2.FONT_HERSHEY_SIMPLEX
            bottom_left_corner_of_text = (constant_position, img.shape[0] - 10)
            font_scale = constant_size
            font_color = (0, 0, 0)
            line_type = 2
            cv2.putText(img, pre_img_save, bottom_left_corner_of_text, font, font_scale, font_color, line_type)
            cv2.imwrite(pre_img_loc, img)
            pixmap = QPixmap(pre_img_loc)
            self.image_label.setPixmap(pixmap)
            self.image_label.setFixedSize(pixmap.size())
        except Exception as e:
            logging.error(f"Error: {str(e)}")
            print(e)


class ExcelSheetSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.Preview_button = None
        self.another_window = None
        self.success_label = None
        self.print_qr_button = None
        self.Generateqr_button = None
        self.setWindowTitle("Excel Sheet Selector")
        self.layout = QVBoxLayout()
        self.file_label = QLabel("No file selected")
        self.layout.addWidget(self.file_label)
        self.folder_label = QLabel("No folder selected")
        self.layout.addWidget(self.folder_label)
        self.select_button = QPushButton("Select Excel File")
        self.select_button.clicked.connect(self.select_excel_file)
        self.select_folder_button = QPushButton("Select Folder ")
        self.select_folder_button.clicked.connect(self.select_folder)
        self.layout.addWidget(self.select_button)
        self.layout.addWidget(self.select_folder_button)
        self.msg_box = QMessageBox()
        self.msg_box.setIcon(QMessageBox.Warning)
        self.msg_box.setText(folder_path)
        self.msg_box.setWindowTitle("Warning")
        self.msg_box.setStandardButtons(QMessageBox.Ok)
        self.setLayout(self.layout)

    def open_another_window(self):
        self.another_window = AnotherWindow()
        self.another_window.show()

    def select_folder(self):
        global folder_path
        options = QFileDialog.Options()
        folder_path = QFileDialog.getExistingDirectory(self, "Select Folder", options=options)
        if folder_path:
            self.folder_label.setText(folder_path)

    def delete_images(self):
        try:
            img_folder_path = '../dummy/img'
            image_files = (glob.glob(os.path.join(img_folder_path, '*.jpg')) + glob.glob(
                os.path.join(img_folder_path, '*.jpeg')) + glob.glob(os.path.join(img_folder_path, '*.png'))
                           + glob.glob(os.path.join(img_folder_path, '*.gif')) + glob.glob(
                        os.path.join(img_folder_path, '*.bmp')))
            # Iterate over the list and delete each file
            for image_file in image_files:
                os.remove(image_file)
                print(f"Deleted: {image_file}")
        except Exception as e:
            logging.error(f"Error: {str(e)}")
            print(e)

    def write_name_on_image(self, name, total_value):
        try:
            img = cv2.imread(image_path)
            font = cv2.FONT_HERSHEY_SIMPLEX
            bottom_left_corner_of_text = (constant_position, img.shape[0] - 10)
            font_scale = constant_size
            font_color = (0, 0, 0)
            line_type = 2
            cv2.putText(img, name, bottom_left_corner_of_text, font, font_scale, font_color, line_type)
            cv2.imwrite("%s/%s.jpg" % (folder_path, name), img)
            total_name_Written.append(name)
            if total_value == len(total_name_Written):
                self.msg_box.setText("Image saved successfully in this location\n%s" % folder_path)
                self.msg_box.exec_()
                self.delete_images()
        except Exception as e:
            logging.error(f"Error: {str(e)}")
            print(e)

    def Generate_qr(self):
        try:
            df = pd.read_excel(file_name)
            total_value = len(df['MAC ID'])
            global image_path
            if folder_path:
                for i in df['MAC ID']:
                    qr_img = qrcode.make(i)
                    qr_img.save('./img/%s.png' % i)
                    image_path = './img/%s.png' % i
                    self.write_name_on_image(i, total_value)
            else:
                self.folder_label.setText("Please Select folder to Save Qr Code")
        except Exception as e:
            logging.error(f"Error: {str(e)}")
            print(e)

    def select_excel_file(self):
        global file_name
        self.delete_images()
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)",
                                                   options=options)
        if file_name:
            self.file_label.setText(file_name)
            self.Generateqr_button = QPushButton("Generate QR")
            self.Generateqr_button.clicked.connect(self.Generate_qr)
            self.layout.addWidget(self.Generateqr_button)
            self.Preview_button = QPushButton("Preview", self)
            self.Preview_button.clicked.connect(self.open_another_window)
            self.layout.addWidget(self.Preview_button)
            self.success_label = QLabel()
            self.layout.addWidget(self.success_label)


if __name__ == "__main__":
    folder_name = ["img", 'pre_img']
    create_folder(folder_name)
    app = QApplication(sys.argv)
    window = ExcelSheetSelector()
    window.show()
    sys.exit(app.exec_())
