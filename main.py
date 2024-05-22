import os, re, datetime
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLabel, QScrollArea

# Diccionario de meses para usar sus valores numericos.
months = {
    "ENE": 1,
    "FEB": 2,
    "MAR": 3,
    "ABR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AGO": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DIC": 12,
}

# Estos archivos solo se presentan una vez y no hace falta actualizarlos
non_updating_files = [
    "DNI",
    "AFIP",
    "ESTATUTO",
    "DDJJ PJ 1014",
    "DDJJ SO 1013",
    "1026",
    "Autoriza. Apod. 1015",
    "RF 1016",
    "Solic. Datos 1004",
    "BF 1023",
    "DNI Autoriza",
    "PODERES",
]

# Estos archivos se actualizan cada 3 meses
quarterly_updating_files = [
    "VENTAS",
    "DEUDAS",
]

current_year = str(datetime.datetime.now().year)[-2:]
current_month = datetime.datetime.now().month

def find_outdated_files(path):
   """
    Lee los archivos buscando aquellos que tengan meses y años declarados,
    e invocando metodos para comprobar si estan desactualizados.

    Argumentos:
        path: La ruta del archivo, provista por el SO

    Retorna:
        outdated_files: Un array que contiene la ruta/nombre de cada archivo desactualizado.
   """

   outdated_files = []
   all_files = get_files(path)

   for file_path in all_files:
        filename, _ = os.path.splitext(file_path)

        filename_parts = filename.split(None, 1)

        if filename_parts[0] in non_updating_files:
            continue

        match = re.search(r".*(\w{3}) (\d{2})", filename)
        if match and filename_parts[0] not in non_updating_files:
            file_month_str = match.group(1)
            if file_month_str in months:
                file_month = months[file_month_str]
                file_year = match.group(2)

                reference_date = int(file_year)

                if is_outdated_yearly(int(file_year), file_month):
                    outdated_files.append(file_path)
                elif filename in quarterly_updating_files:
                    if is_outdated_quarterly(file_month, reference_date):
                        outdated_files.append(file_path)

   return outdated_files 

def get_files(path):

    """
    Consigue todos los archivos en una ruta dada.

    Argumentos:
        path: La ruta de una carpeta.

    Retorna:
        all_files: Un array que contiene los archivos de la carpeta.
    """

    all_files = []
    for root, _, files in os.walk(path):
           for f in files:
               full_path = os.path.join(root, f)
               all_files.append(full_path)
    return all_files

def send_mail(data_frame, file_path):
    '''   
    Envia un mail conteniendo un excel formado por
    los nombres de los archivos que hay que notificar

    Argumentos:
        data_frame: el archivo excel a enviar
    '''
    
    sender_email = "automation.bst@gmail.com"
    receiver_email = "federicoresano1@gmail.com"
    sender_pass = "mrxb dnbt ojcm kata"#"AsQwasqw123"

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = "Legajos"

    msg_body = """
    <html>
    <head></head>
    <body>
    <p>Hola,</p>
    <p>Encuentre debajo el estatus de los archivos.</p>
    <p>Adjunto puede encontrar un archivo con los resultados.</p>
    <p>Desde ya, muchas gracias.</p>
    <p>Saludos.</p>
    <br>
    <table border="1">
        <tr>
            <th>Desactualizados</th>
            <th>Faltantes</th>
        </tr>
    """

    for i in range(len(data_frame)):
        msg_body += "<tr>"
        msg_body += f"<td>{data_frame.iloc[i, 0]}</td>"
        msg_body += f"<td>{data_frame.iloc[i, 1]}</td>"
        msg_body += "</tr>"

    msg_body += """
    </table>
    </body>
    </html>
    """

    msg.attach(MIMEText(msg_body, 'html'))

    with open(file_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
    msg.attach(part)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        #server.starttls()
        server.login(sender_email, sender_pass)
        server.sendmail(sender_email, receiver_email, msg.as_string())

def check_files(window):
    '''
    Recorre todos los archivos de una carpeta recursivamente,
    guarda en un array los que esten marcados como desactualizados 
    o que falten y los muestra en una ventana si los hay o no.

    Argumentos:
        window: usado mas adelante con la libreria PyQt5 para crear
        la interfaz
    '''
    folder_path = select_folder()
    if folder_path:
        all_files = get_files(folder_path)
        all_outdated_files = find_outdated_files(folder_path)

        important_files = [
            "BCE",
            "ASAMBLEA",
            "DIRECTORIO",
            "FLUJOS",
            "FLUJOS PREMISAS",
            "GRUPO ECO. 2005",
            "CLIEN. NO VINC 2006"
        ]
        
        missing_files = []

        for file_name in important_files:
            if not any(file_name.lower() in file_path.lower() for file_path in all_files):
                missing_files.append(file_name)

        text_area = window.findChild(QLabel)
        message = ""

        if missing_files:
                missing_files_str = "\n".join(missing_files)
                message =  "======================================================================================="
                message += f"\nIMPORTANTE: Faltan los siguientes ({len(missing_files)}) archivos importantes:"
                message += "\n=======================================================================================\n"

                message += missing_files_str

                bce_files = [file for file in all_files if "BCE".lower() in file.lower()]
                if len(bce_files) < 3:
                    if len(bce_files) == 2:
                        message += "\n\nFalta 1 archivo de balances"
                    elif len(bce_files) == 1:
                        message += "\n\nFaltan 2 archivos de balances"
                    elif len(bce_files) == 0:
                        message += "\n\nFaltan los archivos de balances"


        if all_outdated_files:
                message += "\n======================================================================================="
                message += f"\nADVERTENCIA: Los siguients archivos ({len(all_outdated_files)}) estan desactualizados:\n"
                message += "======================================================================================="
                outdated_list = ""

                for file_path in all_outdated_files:
                    outdated_list += f"\n - {os.path.basename(file_path)}"

                message += outdated_list

        text_area.setText(message if message else "\nNo se encontraron archivos desactualizados o faltantes")

        max_length = max(len(all_outdated_files), len(missing_files))
        
        pad_outdated_df = [os.path.basename(file) for file in all_outdated_files] + [''] * (max_length - len(all_outdated_files))
        pad_missing_df = missing_files + [''] * (max_length - len(missing_files))

        df = pd.DataFrame({
                "Desactualizados": pad_outdated_df,
                "Faltantes": pad_missing_df
            })


        current_time = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
        folder_name = os.path.basename(folder_path)
        excel_name = os.path.join(folder_path, f"Estatus {folder_name} {current_time}.csv")
        message += f"\nArchivos desactualizados/faltantes escritos en {excel_name}"
        df.to_csv(excel_name, index=False)

        send_mail(df, excel_name)

def is_outdated_yearly(file_year, file_month):

    """
    Comprueba si un archivo esta desactualizado basandonse en el año y mes extraidos del nombre del archivo
    y el mes y año actual.

    Argumentos:
        current_year
        current_month
        file_year : Año extraido del archivo.
        file_month : Mes extraido del archivo.

    Retorna:
        boolean: Verdadero si esta desactualizado.
    """
    year_diff = int(current_year) - file_year
    current_date = datetime.date(int(current_year), current_month, 1)
    file_date = datetime.date(file_year, file_month, 1)
    month_diff = (current_date.month - file_date.month) % 12

    if year_diff == 1:
        return month_diff >= file_month
    else:
        return year_diff > 0

def is_outdated_quarterly(file_month, reference_date):
    '''
    Comprubea si un archivo que se actualiza cada 3 meses necesita
    actualización

    Argumentos:
        file_month: el mes escrito en el archivo
        reference_date: este argumento es el año del archivo,
                        es aumentado por 1 cada vez que se 
                        calcula un archivo que tiene que ser actualizado
                        en un mes del siguiente año, es usado para wrappear
                        alrededor de un año
    Retorna:
        boolean: Verdadero si esta desactualizado
    '''
    file_month_int = months[file_month]
    month_diff = (current_month - file_month_int) % 12

    if file_month_int == 12 and current_month < file_month_int:
        reference_date +=1 
    return month_diff > 3 and reference_date == int(current_year)

def select_folder():
    '''
    Selecciona la carpeta clickeada 
    '''
    folder_path = QFileDialog.getExistingDirectory(None, "Seleccione la carpeta")
    return folder_path

def main():
    '''
    Inicializa la interfaz grafica
    '''
    app = QApplication([])
    window = QWidget()
    window.setWindowTitle("Verificador de archivos")

    window.resize(900, 600)

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
