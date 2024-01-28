from datetime import datetime
import logging
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def calculate_average(p1, p2, p3):
    return (p1 + p2 + p3) / 3

def determine_naf(media_notas):
    if 50 <= media_notas < 70:
        grade = round(70 - media_notas)
        return grade

def determine_situation(media_notas, faltas, total_aulas=60):
    if faltas > total_aulas * 0.25:
        return "Reprovado por Falta", 0

    if media_notas >= 70:
        return "Aprovado", 0

    if 50 <= media_notas < 70:
        naf = determine_naf(media_notas)
        return "Exame Final", naf

    if media_notas < 50:
        return "Reprovado por Nota", 0

    return "Situação não encontrada", 0

def calculate_situation(values):
    cabecalho = values[0]
    cabecalho.append("Situação")
    cabecalho.append("Nota para Aprovação Final")

    for linha in values[1:]:
        if len(linha) < 6:
            continue

        p1 = float(linha[3])
        p2 = float(linha[4])
        p3 = float(linha[5])
        faltas = int(linha[2])

        media_notas = calculate_average(p1, p2, p3)

        situacao, naf = determine_situation(media_notas, faltas)

        linha.append(situacao)
        linha.append(naf)

        logger.info(f"Student: {linha[1]}, Situation: {situacao}, Final Exam Grade: {naf}")

    return values

def formatar_horario_brasileiro(dt):
    return dt.strftime("%d/%m/%Y %H:%M:%S")     
    
def main():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'client_secret.json', SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)

        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId='1PQGV6OwyPd8dJQYijFzRptzn5mwm-j5L9LZIjeuu_9A', range='engenharia_de_software!A3:H27')
            .execute()
        )
        values = result.get("values", [])

        values_with_situation = calculate_situation(values)

        bodyValue = {
            "values": [[row[-2], row[-1]] for row in values_with_situation]
        }
        result = sheet.values().update(
            spreadsheetId='1PQGV6OwyPd8dJQYijFzRptzn5mwm-j5L9LZIjeuu_9A',
            range='engenharia_de_software!G3:H27',
            valueInputOption="USER_ENTERED",
            body=bodyValue
        ).execute()

        if not os.path.exists("last_update.txt"):
            with open("last_update.txt", "w", encoding="utf-8") as txt_file:
                horario_atual = datetime.now()
                horario_formatado = formatar_horario_brasileiro(horario_atual)
                txt_file.write(f"Última atualização às {horario_formatado}")
        else:
            with open("last_update.txt", "r+", encoding="utf-8") as txt_file:
                conteudo = txt_file.read()

                posicao_ultima_atualizacao = conteudo.rfind("Última atualização às")

                atualizacoes_anteriores = conteudo[:posicao_ultima_atualizacao].strip()
                ultima_atualizacao = conteudo[posicao_ultima_atualizacao:].strip()

                horario_atual = datetime.now()
                horario_formatado = formatar_horario_brasileiro(horario_atual)
                novo_conteudo = f"Última atualização às {horario_formatado}\n{atualizacoes_anteriores}"

                txt_file.seek(0)
                txt_file.truncate()
                txt_file.write(novo_conteudo)

                txt_file.write("\nAtualização anterior, " + ultima_atualizacao[24:])

        service = build("sheets", "v4", credentials=creds)

        sheet = service.spreadsheets()
        result = (
            sheet.values().get(
                spreadsheetId='1PQGV6OwyPd8dJQYijFzRptzn5mwm-j5L9LZIjeuu_9A', 
                range='engenharia_de_software!A3:H27')
            .execute()
        )
        values = result.get("values", [])

        values_with_situation = calculate_situation(values)

        bodyValue = {
            "values": [[row[-2], row[-1]] for row in values_with_situation]
        }
        result = sheet.values().update(
            spreadsheetId='1PQGV6OwyPd8dJQYijFzRptzn5mwm-j5L9LZIjeuu_9A',
            range='engenharia_de_software!G3:H27',
            valueInputOption="USER_ENTERED",
            body=bodyValue
        ).execute()

        requests = []
        for index, row in enumerate(values_with_situation):
            if len(row) >= 7:
                color = None
                if row[-2] == "Aprovado":
                    color = {"red": 0.9, "green": 1, "blue": 0.9}
                elif row[-2] == "Exame Final":
                    color = {"red": 1, "green": 0.8, "blue": 0.8}
                elif row[-2] in ["Reprovado por Falta", "Reprovado por Nota"]:
                    color = {"red": 1, "green": 0.8, "blue": 0.8}

                if color:
                    requests.append({
                        "updateCells": {
                            "range": {"sheetId": 0, "startRowIndex": index + 2, "endRowIndex": index + 3},
                            "rows": [{"values": [{"userEnteredFormat": {"backgroundColor": color}}]}],
                            "fields": "userEnteredFormat.backgroundColor"
                        }
                    })

                    for i, _ in enumerate(row[:7]):
                        requests[-1]["updateCells"]["rows"][0]["values"].append({"userEnteredFormat": {"backgroundColor": color}})

        if requests:
            batch_update_request = {"requests": requests}
            service.spreadsheets().batchUpdate(
                spreadsheetId='1PQGV6OwyPd8dJQYijFzRptzn5mwm-j5L9LZIjeuu_9A',
                body=batch_update_request
            ).execute()
    
    except HttpError as err:
        print(err)

if __name__ == "__main__":
    main()