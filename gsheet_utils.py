import os
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials


def get_gsheet_client():
    import tempfile

    creds_path = os.environ.get(
        "GOOGLE_SHEETS_CREDS",
        r"C:/Users/lucas/OneDrive/Documentos/SITE-RIOELAS-TESTE/identificador-488615-c1ab55e9b31b.json",
    )
    creds_content = os.environ.get("GOOGLE_SHEETS_CREDS_CONTENT")

    if creds_content:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".json")
        tmp.write(creds_content.encode("utf-8"))
        tmp.close()
        creds_path = tmp.name

    if not os.path.exists(creds_path):
        raise FileNotFoundError(
            f"Arquivo de credencial não encontrado: {creds_path}\n"
            "Verifique o caminho ou a variável de ambiente GOOGLE_SHEETS_CREDS ou GOOGLE_SHEETS_CREDS_CONTENT."
        )

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    creds = Credentials.from_service_account_file(creds_path, scopes=scopes)
    return gspread.authorize(creds)


def append_to_sheet(data):
    client = get_gsheet_client()
    sheet_id = os.environ.get("GOOGLE_SHEETS_ID", "1fgYehlrnCDEClPOSao2VOA7tnwNCGy3l9qj5dtVMwAg")
    sheet_name = os.environ.get("GOOGLE_SHEETS_TAB", "INFORMAÇÕES")
    sheet = client.open_by_key(sheet_id).worksheet(sheet_name)

    header = [
        "Data Envio",
        "Nome do Local",
        "Região",
        "Endereço Completo",
        "CEP",
        "Cursos",
        "Local da Turma",
        "Horário",
        "Vagas",
        "Turma",
        "Dias de Aula",
        "Data de Início",
        "Encerramento",
        "Cor da Ficha",
    ]

    existing_header = sheet.row_values(1)
    if existing_header != header:
        if existing_header:
            sheet.delete_rows(1)
        sheet.insert_row(header, 1)

    now = datetime.now().strftime("%d/%m/%Y")
    data_to_save = [now] + data
    next_row = len(sheet.get_all_values()) + 1
    sheet.insert_row(data_to_save, next_row)
