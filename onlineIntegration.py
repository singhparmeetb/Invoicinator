import os
import requests
import datetime
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import gspread
from google.oauth2 import service_account
from oauth2client.service_account import ServiceAccountCredentials

def authenticate_google_drive():
    # gauth = GoogleAuth()
    # gauth.LocalWebserverAuth()
    # drive = GoogleDrive(gauth)
    # return drive
    credentials = service_account.Credentials.from_service_account_file(
        'invoicing-410113-key.json', scopes=['https://www.googleapis.com/auth/drive']
    )
    gc = gspread.authorize(credentials)
    return gc


def authenticate_pdfco():
    api_key = 'jixekeb127@ubinert.com_Tm9HSFa92nI9d59enc6682FlD7Ja2SuSI56bnMRX2ukZGpgCH1p7wZ5F49PD9mN0n02KN16xZ7LXOJFmm2Ol3A4a149S0L23hl0a42neM3X1emuQA5Ts82gLhiN3YP892045zNotUW3xnlq07UO5Mplx6z'
    return api_key

def open_google_sheet(spreadsheet_id, sheet_name):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('invoicing-410113-key.json', scope)
    gc = gspread.authorize(credentials)
    spreadsheet = gc.open_by_key(spreadsheet_id)
    sheet = spreadsheet.worksheet(sheet_name)
    return sheet


def extract_text_from_pdfco(pdf_url, api_key):
    endpoint = 'https://api.pdf.co/v1/pdf/extract/text'

    headers = {'content-type': 'application/json', 'x-api-key': api_key}

    payload = {'url': pdf_url}
    response = requests.post(endpoint, json=payload, headers=headers)

    if response.status_code == 200:
        result = response.json()
        return result['text']

    return None

def parse_invoice_data(data):

    invoice_number = next(obj['value'] for obj in data['objects'] if obj['name'] == 'invoiceId')

    date_issued_str = next(obj['value'] for obj in data['objects'] if obj['name'] == 'dateIssued')

    date_issued = datetime.datetime.strptime(date_issued_str, "%Y-%m-%dT%H:%M:%S")
    gross_amount = next(obj['value'] for obj in data['objects'] if obj['name'] == 'total')
    sub_total = next(obj['value'] for obj in data['objects'] if obj['name'] == 'subTotal')
    tax_amount = next(obj['value'] for obj in data['objects'] if obj['name'] == 'tax')

    net_amount = float(gross_amount) - float(tax_amount)

    return {
        "Invoice Number": invoice_number,
        "Date Issued": date_issued.strftime("%Y-%m-%d"),
        "Gross Amount": gross_amount,
        "Tax Amount": tax_amount,
        "Net Amount": net_amount
    }

def write_data_to_sheet(data, sheet):
    sheet.append_row([data["Invoice Number"], data["Date Issued"],data["Gross Amount"], data["Tax Amount"], data["Net Amount"]])

def move_to_processed_folder(drive, file_id, processed_folder_id):
    pdf_file = drive.CreateFile({'id': file_id})
    pdf_file.Upload({'parents': [{'id': processed_folder_id}]})

    print(f"Moved to 'processed' folder: {pdf_file['title']}")

def process_pdf(drive, pdf_url, pdfco_api_key, sheet):
    pdf_text = extract_text_from_pdfco(pdf_url, pdfco_api_key)

    if pdf_text:
        invoice_data = parse_invoice_data(pdf_text)
        write_data_to_sheet(invoice_data, sheet)
        move_to_processed_folder(drive, pdf_file['id'], processed_folder_id)

def main():
    drive = authenticate_google_drive()
    pdfco_api_key = authenticate_pdfco()

    spreadsheet_id = '1omHGcuLSbvWtXkGCykSUpGCLHFRCiovHH5hFCQ2Xyos'
    sheet_name = 'Sheet1'
    sheet = open_google_sheet(spreadsheet_id, sheet_name)

    folder_id = '1ZrAs1EFZZerf2z-aiQCajgYGI68MXWgA'

    processed_folder_id = '1tOFNomY0zkB0rhwE5WOxcCJMLr471td_'
    file_list = drive.ListFile({'q': f"'{folder_id}' in parents and trashed=false"}).GetList()


    for pdf_file in file_list:
        if pdf_file['title'].endswith('.pdf'):
            pdf_url = pdf_file['exportLinks']['application/pdf']
            process_pdf(drive, pdf_url, pdfco_api_key, sheet)

if __name__ == "__main__":
    main()
