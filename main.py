import os, re, datetime
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QMessageBox, QDialog, QLabel, QScrollArea
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
    "   1026",
    "AUTORIZA. APOD. 1015",
    "RF 1016",
    "SOLIC. DATOS 1004",
    "BF 1023",
    "DNI AUTORIZA.",
    "PODERES",  
]

# Estos archivos se actualizan cada 3 meses
quarterly_updating_files = [
    #"VENTAS",
    "VENT",
    "DEUD",
    #"DEUDAS",
]

#Archivos requeridos por cada tipo de cliente
required_files = {
    'persona': {
        "AFIP" : "AFIP", "BALANCE" : "BCE INDIV", "CERTIFICADO CUMPLIMIENTO CENSAL AGRO" : "CERTIFI. CUMPLIMIENTO CENSAL AGRO", "CERTIFICADO CUMPLIMIENTO CENSAL" : "CERTIFI. CUMPLIMIENTO CENSAL", "DECLARACION JURADA BIENES PERSONALES" : "DDJJ BIENES PERS","DECLARACION JURADA INGRESOS" : "DDJJ IIGG",
        "DECLARACION JURADA VINCULACION ENTIDAD" : "DDJJ VINCULACION ENTIDAD","DNI" : "DNI", "FOND" : "FONDOS","INFORMACION LEGAL" : "INFO LEGAL", "DNI AUTORIZACION" : "DNI AUTORIZA",  "POD" : "PODERES"
},
    'pyme': {
        "BCE" : "BALANCE", "VENT": "VENTAS", "DEUD": "DEUDAS", "FLUJOS": "FLUJ", "ACTA DE DIRECTORIO" : "ACTA DIR", "ESTATUTO" : "EST", "REFORMA DE ESTATUTO": "REFORM EST", "PODERES" : "POD", "FORMULARIO 1025" : "FORM 1025", "FORMULARIO 2006" : "FORM 2006", "FORMULARIO 2005" : "FORM 2005", "DECLARACION JURADA PJ 1014" : "DDJJ PJ 1014", "DECLARACION JURADA SO 1013" : "DDJJ SO 1013",
        "AUTORIZACION APODERADO 1015" : "AUTORIZA. APOD. 1015", "RF. 1016" :"RF 1016", "SOLICITUD DE DATOS 1004" : "SOLIC DATOS 1004", "BF. 1023" : "BF 1023", "AFIP" : "AFIP", "DNI AUTORIZACION" : "DNI AUTORIZA", "DECLARACION JURADA VINCULACION ENTIDAD" : "DDJJ VINCULACION ENTIDAD", "CERT MIPYME" : "CERTIFICADO MIPYME"
},                      
    'empresa': {
        "BCE" : "BALANCE","VENT": "VENTAS", "DEUD": "DEUDAS", "ACTA DIRECTORIO": "ACTA DIR", "FLUJOS": "FLUJ", "FLUJOS PREMISAS": "FLUJ PREM", "ACTA ASAMBLEA": "ACTA AS", "GRUPO ECO. 2005" : "GRUPO ECONOMICO", "CLIEN. NO VINC 2006": "CLIENTE NO VINCULADO", "FORM 1025" : "FORMULARIO 1025"
    }
}

current_year = str(datetime.datetime.now().year)[-2:]
current_month = datetime.datetime.now().month

def get_client_type(folder_name):
    """
    Identifica el tipo de cliente basado en el nombre de la carpeta.                    

    Arguments:
        folder_name: Nombre de la carpeta.

    Returns:
        client_type: Un string que representa el tipo de cliente ('persona', 'pyme', 'empresa').
    """

    folder_basename = os.path.basename(folder_name)
    logger.info(f"Checking folder: {folder_basename}")  

    if re.search(r"\b\d{11}\b", folder_basename):
        logger.info(f"Identified as 'empresa': {folder_basename}")
        return 'empresa'
    elif 'PYME' in folder_basename.upper():
        logger.info(f"Identified as 'pyme': {folder_basename}")
        return 'pyme'
    elif re.search(r"\b\d{2}-\d{8}-\d\b", folder_basename):
        logger.info(f"Identified as 'persona': {folder_basename}")
        return 'persona'
    else:
        logger.info("Identified as 'desconocido'")
        return 'desconocido'
    
def get_current_year_bce_files():
    """
    Obtiene los nombres esperados de los archivos BCE para el año actual, el año anterior y dos años antes.

    Retorna:
        expected_bce_files: Lista con los nombres de los archivos BCE esperados.
    """
    current_year = datetime.datetime.now().year
    return [
        f"BCE DIC {str(current_year)[-2:]}",
        f"BCE DIC {str(current_year - 1)[-2:]}",
        f"BCE DIC {str(current_year - 2)[-2:]}"
    ]

def find_latest_bce_file(files):
    """
    Encuentra el archivo más reciente de los archivos BCE.

    Argumentos:
        files: Lista de archivos BCE.

    Retorna:
        El archivo BCE más reciente.
    """
    latest_bce_file = None
    latest_date = None

    for file in files:
        match = re.search(r".*BCE (\w{3}) (\d{2})", file, re.IGNORECASE)
        if match:
            file_month_str = match.group(1)
            file_year = match.group(2)
            if file_month_str in months:
                file_month = months[file_month_str]
                file_date = datetime.date(int("20" + file_year), file_month, 1)
                if not latest_date or file_date > latest_date:
                    latest_date = file_date
                    latest_bce_file = file

    return latest_bce_file

def extract_year_month(filename):
    """
    Extrae el año y mes de un nombre de archivo.

    Argumentos:
        filename: El nombre del archivo.

    Retorna:
        Una tupla (año, mes).
    """
    match = re.search(r".*(\w{3}) (\d{2})", filename)
    if match:
        file_month_str = match.group(1)
        file_year = match.group(2)
        if file_month_str in months:
            file_month = months[file_month_str]
            return int("20" + file_year), file_month
    return None, None

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

    bce_files = [file for file in all_files if re.search(r".*BCE (\w{3}) (\d{2})", file, re.IGNORECASE)]
    latest_bce_file = find_latest_bce_file(bce_files)

    if latest_bce_file:
        latest_year, latest_month = extract_year_month(latest_bce_file)
        if not is_outdated_yearly(latest_year, latest_month):
            bce_files.remove(latest_bce_file)

    for file_path in all_files:
        if "BCE" in file_path:
            continue 

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
                    outdated_files.append((file_path, os.path.dirname(path)))
                elif filename in quarterly_updating_files:
                    if is_outdated_quarterly(file_month, reference_date):
                        outdated_files.append((file_path, os.path.dirname(path)))

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
    receiver_email = "federicoresano1@gmail.com"#"vracauchi@grupost.com.ar"#"arojas@bst.com.ar"
    sender_pass = "mrxb dnbt ojcm kata"

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

    clients = set(data_frame.iloc[:, 0])

    for client in clients:
        msg_body += f"<tr><td>{client}</td></tr>"


    msg_body += """
    </table>
    </body>
    </html>
    """

    msg.attach(MIMEText(msg_body, 'html'))
    try:
        with open(file_path, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())

        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file_path)}')
        msg.attach(part)

        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            #server.ttls()
            server.login(sender_email, sender_pass)
            server.sendmail(sender_email, receiver_email, msg.as_string())

              
        logger.info("Email enviado a %s", receiver_email)
    except smtplib.SMTPAuthenticationError:
        logger.error("Fallo la autenticacion con el servidor SMTP. Usuario y contraseña incorrectos.")
    except smtplib.SMTPException as e:
        logger.error("Error de STMP: %s", e)
    except Exception as e:
        logger.error("Error: %s", e)

def write_excel_report(outdated_data, missing_data, output_path):
    """
    Escribe un excel en el directorio que contiene los clientes.

    Arguments:
        report_data: Diccionario que contiene la información para el reporte.
        output_path: Directorio donde el excel se escribe.
    """
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

    df_outdated = pd.DataFrame(outdated_data, columns=['Cliente', 'Tipo', 'Archivo', 'Carpeta'])
    df_outdated.to_excel(writer, sheet_name='Desactualizados', index=False)

    df_missing = pd.DataFrame(missing_data, columns=['Cliente', 'Tipo', 'Archivo', 'Carpeta'])
    df_missing.to_excel(writer, sheet_name='Faltantes', index=False)

    writer._save()
    return df_outdated, df_missing


def check_files(window):
    '''
    Recorre todos los archivos de una carpeta recursivamente,
    guarda en un array los que esten marcados como desactualizados 
    o que falten y los muestra en una ventana si los hay o no.

    Argumentos:
        window: Usado mas adelante con la libreria PyQt5 para crear
        la interfaz.
    '''
    folder_path = select_folder()
    if not folder_path:
        return
    
    client_folders = next(os.walk(folder_path))[1]
    if not client_folders:
        QMessageBox.warning(window, "Error", "No se encontraron carpetas de clientes.")
        return

    all_outdated_files = []
    all_missing_files = []

    for client_folder in client_folders:
        client_path = os.path.join(folder_path, client_folder)
        client_type = get_client_type(client_path)

        if client_type == 'unknown':
            logger.warning("No se pudo determinar el tipo de cliente para la carpeta %s.", client_folder)
            continue

        all_files = get_files(client_path)
        outdated_files = find_outdated_files(client_path)
        important_files = required_files.get(client_type, {})

        missing_files = []
        file_count = {key: 0 for key in important_files}

        for file_path in all_files:
            lower_path = file_path.lower()
            for key in important_files:
                if key.lower() in lower_path:
                    file_count[key] += 1

        for key, count in file_count.items():
            if count < 1:
                 missing_files.append((client_folder, client_type, key, client_path))

        non_updating_missing_files = []
        for non_updating_file in important_files:
            if not any(non_updating_file.lower() in file.lower() for file in all_files):
                non_updating_missing_files.append((client_folder, client_type, non_updating_file, client_path))

        for file in outdated_files:
            all_outdated_files.append((client_folder, client_type, os.path.basename(file[0]).replace('.', ''), file[1]))

        all_missing_files.extend(missing_files + non_updating_missing_files)
        
    message = ""

    if all_missing_files:
        message += (
            "=======================================================================================\n"
            "IMPORTANTE: Se encontraron archivos importantes faltantes:\n"
            "=======================================================================================\n"
        )
        for file in all_missing_files:
            message += f"{file[0]} ({file[2]}) en {file[3]}\n"

    if all_outdated_files:
        message += (
            "=======================================================================================\n"
            "ADVERTENCIA: Se encontraron archivos desactualizados:\n"
            "=======================================================================================\n"
        )
        for file in all_outdated_files:
            message += f"{file[0]} ({file[2]}) en {file[3]}\n"

    if not message:
        message = "\nNo se encontraron archivos desactualizados o faltantes para ninguno de los clientes."

    if all_outdated_files or all_missing_files:
        parent_folder = os.path.dirname(folder_path)
        report_filename = os.path.join(parent_folder, 'Resultado_de_verificacion.xlsx')
        df_outdated, df_missing = write_excel_report(all_outdated_files, all_missing_files, report_filename)
        QMessageBox.information(window, "Reporte Creado", f"Se ha creado el reporte en: {report_filename}")
        send_mail(pd.concat([df_outdated, df_missing]), report_filename)

    dialog = QDialog(window)
    dialog.setWindowTitle("Resultados de la verificación")
    dialog.setMinimumWidth(400)

    scroll_area = QScrollArea(dialog)
    scroll_area.setWidgetResizable(True)

    scroll_content = QWidget()
    scroll_layout = QVBoxLayout(scroll_content)

    label = QLabel(message)
    label.setWordWrap(True)

    scroll_layout.addWidget(label)
    scroll_area.setWidget(scroll_content)

    main_layout = QVBoxLayout()
    main_layout.addWidget(scroll_area)

    ok_button = QPushButton("OK")
    ok_button.clicked.connect(dialog.accept)
    main_layout.addWidget(ok_button)

    dialog.setLayout(main_layout)
    dialog.exec_()

    #popup = QMessageBox(window)
    #popup.setWindowTitle("Resultados de la Verificación")
    #popup.setText(message)
    #popup.setStandardButtons(QMessageBox.Ok)
    #popup.exec_()

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

def is_outdated_quarterly(file_month, file_year):
    """
    Comprueba si un archivo trimestral está desactualizado basándose en el año y mes extraídos del nombre del archivo
    y el mes y año actual.

    Argumentos:
        file_month: Mes extraído del archivo.
        file_year: Año extraído del archivo.

    Retorna:
        boolean: Verdadero si está desactualizado.
    """
    current_year_full = datetime.datetime.now().year
    current_month = datetime.datetime.now().month
    current_quarter = (current_month - 1) // 3 + 1
    file_quarter = (file_month - 1) // 3 + 1

    if file_year < current_year_full:
        return True
    elif file_year == current_year_full:
        return file_quarter < current_quarter
    return False

def select_folder():
    """
    Selecciona la carpeta clickeada 
    """
    folder_path = QFileDialog.getExistingDirectory(None, "Seleccione la carpeta")
    return folder_path

def main():
    """
    Inicializa la interfaz grafica
    """
    app = QApplication([])
    window = QWidget()
    window.setWindowTitle("Verificador de archivos")

    window.resize(900, 600)

    layout = QVBoxLayout()
    window.setLayout(layout)

    select_folder_button = QPushButton("Seleccionar carpeta")
    select_folder_button.clicked.connect(lambda: check_files(window))
    layout.addWidget(select_folder_button)

    window.show()
    app.exec_()


if __name__ == "__main__":
    main()
