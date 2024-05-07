import os, re, datetime
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QScrollArea

def read_files(path):
   outdated_files = []
   all_files = get_files(path)
   current_year = str(datetime.date.today().year)[-2:]
   for file_path in all_files:
       if file_path.endswith(".pdf"):
           filename, _ = os.path.splitext(file_path)
           match = re.search(r"\d{2}$", filename)

           if match:
               file_number = int(match.group())
               if file_number < int(current_year):
                   outdated_files.append(os.path.join(path, filename))
   return outdated_files

def get_files(path):
   all_files = []
   for root, _, files in os.walk(path):
       for f in files:
           full_path = os.path.join(root, f)
           all_files.append(full_path)
   return all_files

def select_folder():
    folder_path = QFileDialog.getExistingDirectory(None, "Seleccione la carpeta")
    return folder_path

def check_files(window):
    folder_path = select_folder()
    if folder_path:
        all_outdated_files = read_files(folder_path)
        text_area = window.findChild(QLabel)
        if all_outdated_files:
            message = f"ADVERTENCIA: Los siguients archivos ({len(all_outdated_files)}) estan desactualizados.\n"
            outdated_list = ""

            for file_path in all_outdated_files:
                outdated_list += f"\n - {file_path}"

            message += outdated_list
            data_frame = pd.DataFrame({"Desactualizados": all_outdated_files})
            file_name = os.path.join(folder_path, "Estatus de legajos.csv")
            data_frame.to_csv(file_name, index=False)
            message += f"\n\nLista de archivos desactualizados escrita en: {folder_path} - Estatus de legajos.csv"

            text_area.setText(message)
        else:
          text_area.setText("\nNo se encontraron archivos desactualizados.")

def main():
    app = QApplication([])
    window = QWidget()
    window.setWindowTitle("Verificador de archivos")

    window.resize(600, 400)

    layout = QVBoxLayout()
    window.setLayout(layout)

    select_folder_button = QPushButton("Seleccionar carpeta")
    select_folder_button.clicked.connect(lambda: check_files(window))
    layout.addWidget(select_folder_button)

    text_area = QLabel("")
    text_area.setWordWrap(True)
    text_area.setStyleSheet("border: 1px solid black")

    scroll_area = QScrollArea()
    scroll_area.setWidgetResizable(True)
    scroll_area.setWidget(text_area)

    layout.addWidget(scroll_area)

    window.show()
    app.exec_()


if __name__ == "__main__":
    main()
