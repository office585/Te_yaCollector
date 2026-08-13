import base64
import json
import os
import time
from datetime import datetime, timedelta
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from playwright.sync_api import sync_playwright
import requests

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# ==================== KÖRNYEZETI VÁLTOZÓK (GITHUB SECRETS) ====================
TEYA_EMAIL = os.getenv("TEYA_EMAIL")
TEYA_PASSWORD = os.getenv("TEYA_PASSWORD")
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID", "0AP3HkLh_ANsVUk9PVA")
GDRIVE_SA_JSON_STR = os.getenv("GDRIVE_SA_JSON")

LOGIN_URL = "https://business.teya.com/?locale=hu"
COMPANY_ID = "24abe7e0-88fd-468a-afb6-fe362b0b79ca"
STORE_ID = "81e77008-fc09-4a82-9c78-b5939b24a202"

# ELMÚLT 14 NAP (2 HÉT) GYŰJTÉSE
DAYS_BACK = 14

OUTPUT_DIR = "./output"
GRAPHQL_URL = "https://customer-bff.teya.com/graphql"
# ==============================================================================


def upload_to_google_drive(file_path, file_name, folder_id):
    """Fájl feltöltése vagy FELÜLÍRÁSA Google Drive-on."""
    if not GDRIVE_SA_JSON_STR or not folder_id:
        return False

    print(f"      [Drive] Szinkronizálás: {file_name}...")
    try:
        sa_info = json.loads(GDRIVE_SA_JSON_STR)
        SCOPES = ["https://www.googleapis.com/auth/drive"]
        creds = service_account.Credentials.from_service_account_info(sa_info, scopes=SCOPES)
        service = build("drive", "v3", credentials=creds)

        # Keresés, hogy létezik-e már a fájl
        query = f"name='{file_name}' and '{folder_id}' in parents and trashed=false"
        results = service.files().list(
            q=query, 
            spaces='drive', 
            fields="files(id, name)", 
            supportsAllDrives=True, 
            includeItemsFromAllDrives=True
        ).execute()
        
        items = results.get('files', [])
        media = MediaFileUpload(
            file_path,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=True,
        )

        if items:
            # Fájl felülírása (Update)
            file_id = items[0]['id']
            service.files().update(
                fileId=file_id, 
                media_body=media, 
                supportsAllDrives=True
            ).execute()
            print(f"      [Drive SIKER] Fájl felülírva (Frissítve).")
        else:
            # Új fájl létrehozása (Create)
            file_metadata = {"name": file_name, "parents": [folder_id]}
            service.files().create(
                body=file_metadata,
                media_body=media,
                fields="id",
                supportsAllDrives=True
            ).execute()
            print(f"      [Drive SIKER] Új fájl feltöltve.")
        return True

    except Exception as e:
        print(f"      [Drive HIBA] Nem sikerült a művelet ({file_name}): {e}")
        return False


def format_eur(amount):
    if amount is None: amount = 0.0
    formatted = f"{amount:,.2f}".replace(",", "X").replace(".", ",").replace("X", " ")
    return f"{formatted} €"


def create_summary_excel(file_path, summary, payout_date):
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Összegzés"
        headers = ["Dátum", "Kifizetett Végösszeg", "Kezdeti egyenleg", "Tranzakciók bruttó", "Visszatérítések", "Chargeback", "Bankköltség / Díjak"]
        values = [
            payout_date, summary.get("kifizetes", "0,00 €"), summary.get("kezdeti_egyenleg", "0,00 €"),
            summary.get("tranzakciok", "0,00 €"), summary.get("visszateritesek", "0,00 €"),
            summary.get("chargeback", "0,00 €"), summary.get("dijak", "0,00 €")
        ]
        ws.append(headers)
        ws.append(values)
        for col_num in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=col_num)
            cell.font = Font(bold=True)
            if col_num == 2:
                cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
                val_cell = ws.cell(row=2, column=col_num)
                val_cell.fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
                val_cell.font = Font(bold=True)
        wb.save(file_path)
        return True
    except Exception as e:
        return False


def append_to_master_summary(master_path, summary, payout_date):
    headers = ["Dátum", "Kifizetett Végösszeg", "Kezdeti egyenleg", "Tranzakciók bruttó", "Visszatérítések", "Chargeback", "Bankköltség / Díjak"]
    values = [
        payout_date, summary.get("kifizetes", "0,00 €"), summary.get("kezdeti_egyenleg", "0,00 €"),
        summary.get("tranzakciok", "0,00 €"), summary.get("visszateritesek", "0,00 €"),
        summary.get("chargeback", "0,00 €"), summary.get("dijak", "0,00 €")
    ]
    try:
        if not os.path.exists(master_path):
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Master Összegzés"
            ws.append(headers)
            for col_num in range(1, len(headers) + 1):
                ws.cell(row=1, column=col_num).font = Font(bold=True)
        else:
            wb = openpyxl.load_workbook(master_path)
            ws = wb.active

        # FELÜLÍRÁS: Ha a dátum már létezik, felülírjuk a sort
        row_to_overwrite = None
        for row_idx in range(2, ws.max_row + 1):
            if ws.cell(row=row_idx, column=1).value == payout_date:
                row_to_overwrite = row_idx
                break

        if row_to_overwrite:
            for col_idx, val in enumerate(values, start=1):
                ws.cell(row=row_to_overwrite, column=col_idx, value=val)
        else:
            ws.append(values)
            
        wb.save(master_path)
    except Exception as e:
        print(f"      [Master Összegzés HIBA]: {e}")


def append_to_master_raw(master_path, daily_raw_path, payout_date):
    if not os.path.exists(daily_raw_path):
        return
    try:
        daily_wb = openpyxl.load_workbook(daily_raw_path)
        daily_ws = daily_wb.active

        if not os.path.exists(master_path):
            master_wb = openpyxl.Workbook()
            master_ws = master_wb.active
            master_ws.title = "Master Tranzakciók"
            
            # Fejléc másolása + Egy Rendszer Dátum oszlop hozzáadása
            header_row = list(next(daily_ws.iter_rows(min_row=1, max_row=1, values_only=True)))
            header_row.append("API_Lekeres_Datuma")
            master_ws.append(header_row)
        else:
            master_wb = openpyxl.load_workbook(master_path)
            master_ws = master_wb.active

            # FELÜLÍRÁS: Visszafelé töröljük a régi sorokat, amik ehhez a naphoz tartoznak
            for row_idx in range(master_ws.max_row, 1, -1):
                if master_ws.cell(row=row_idx, column=master_ws.max_column).value == payout_date:
                    master_ws.delete_rows(row_idx, 1)

        # Új adatsorok hozzáfűzése a napi fájlból
        for row_idx, row_values in enumerate(daily_ws.iter_rows(values_only=True), start=1):
            if row_idx < 2:  # Fejlécet átugorjuk
                continue
            row_list = list(row_values)
            row_list.append(payout_date)  # Hozzáadjuk a dátumot az utolsó oszlophoz
            master_ws.append(row_list)

        master_wb.save(master_path)
    except Exception as e:
        print(f"      [Master Nyers HIBA]: {e}")


def generate_and_download_raw_excel(payout_id, save_path, headers):
    gen_payload = {
        "query": "query GenerateTransactionsReport($input: GenerateTransactionsReportInput!) {\n  generateTransactionsReport(input: $input) {\n    id\n  }\n}",
        "variables": {"input": {"storeId": STORE_ID, "payoutId": payout_id, "locale": "hu"}},
        "operationName": "GenerateTransactionsReport"
    }

    try:
        gen_res = requests.post(GRAPHQL_URL, headers=headers, json=gen_payload)
        if gen_res.status_code != 200: return False
        job_id = gen_res.json().get("data", {}).get("generateTransactionsReport", {}).get("id")
        if not job_id: return False

        poll_payload = {
            "query": "query GetSettlementReportsByJobId($jobId: ID!) {\n  getSettlementReportsByJobId(jobId: $jobId) {\n    jobId\n    reports {\n      documentCompleted\n      documentId\n      documentMimeType\n      currency\n      id\n      status\n    }\n    status\n  }\n}",
            "variables": {"jobId": job_id},
            "operationName": "GetSettlementReportsByJobId"
        }

        document_id = None
        for attempt in range(1, 12):
            time.sleep(1)
            poll_res = requests.post(GRAPHQL_URL, headers=headers, json=poll_payload)
            if poll_res.status_code == 200:
                job_data = poll_res.json().get("data", {}).get("getSettlementReportsByJobId", {})
                reports = job_data.get("reports", [])
                if reports and reports[0].get("documentCompleted"):
                    document_id = reports[0].get("documentId")
                    break

        if not document_id: return False

        dl_payload = {
            "query": "query DownloadDocument($input: DownloadDocumentInput!) {\n  downloadDocument(input: $input) {\n    documentBase64\n  }\n}",
            "variables": {"input": {"documentId": document_id, "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}},
            "operationName": "DownloadDocument"
        }

        dl_res = requests.post(GRAPHQL_URL, headers=headers, json=dl_payload)
        if dl_res.status_code == 200:
            b64_str = dl_res.json().get("data", {}).get("downloadDocument", {}).get("documentBase64")
            if b64_str:
                excel_bytes = base64.b64decode(b64_str)
                with open(save_path, "wb") as f:
                    f.write(excel_bytes)
                return True
    except Exception:
        pass
        
    return False


def fetch_all_settlements_in_chunks(start_date, end_date, headers):
    settlements_url = f"https://business.teya.com/reporting/v1/company/{COMPANY_ID}/settlements"
    all_items = []
    current_start = start_date

    while current_start < end_date:
        current_end = min(current_start + timedelta(days=25), end_date)
        start_str = current_start.strftime("%Y-%m-%d")
        end_str = current_end.strftime("%Y-%m-%d")
        
        offset = 0
        limit = 50
        while True:
            payload = {
                "begin_date": start_str,
                "end_date": end_str,
                "store_ids": [],
                "states": ["POSTPONED", "PARTIALLY_PAID", "PAID", "SCHEDULED"],
                "limit": limit,
                "offset": offset
            }
            res = requests.put(settlements_url, headers=headers, json=payload)
            if res.status_code != 200: break
            data = res.json().get("data", [])
            all_items.extend(data)
            if len(data) < limit: break
            offset += limit

        current_start = current_end + timedelta(days=1)
        
    return all_items


def run_pure_api_collector():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    captured_token = None

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        def handle_request(request):
            nonlocal captured_token
            auth_header = request.headers.get("authorization", "")
            if auth_header.startswith("Bearer ") and not captured_token:
                captured_token = auth_header

        page.on("request", handle_request)
        page.goto(LOGIN_URL)
        page.wait_for_selector("#username", timeout=15000)
        page.locator("#username").fill(TEYA_EMAIL)
        page.press("#username", "Enter")

        password_selector = 'input[type="password"], input[name="password"]'
        page.wait_for_selector(password_selector, timeout=15000)
        page.locator(password_selector).fill(TEYA_PASSWORD)
        page.press(password_selector, "Enter")

        page.goto("https://business.teya.com/settlements")
        try:
            page.wait_for_selector('[data-testid="settlements-calendar"]', timeout=25000)
        except Exception:
            page.wait_for_timeout(3000)
        browser.close()

    if not captured_token:
        print("[SÚLYOS HIBA] Nem sikerült elcsípni a munkamenet-tokent.")
        return

    headers = {
        "accept": "*/*",
        "authorization": captured_token,
        "content-type": "application/json",
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/151.0.0.0",
    }

    # IDŐSZAK: Utolsó 14 nap (2 hét)
    end_date = datetime.now()
    start_date = end_date - timedelta(days=DAYS_BACK)

    print(f"\n[API] Adatok lekérése az elmúlt {DAYS_BACK} napra ({start_date.strftime('%Y-%m-%d')} -> {end_date.strftime('%Y-%m-%d')})...")
    data_items = fetch_all_settlements_in_chunks(start_date, end_date, headers)

    master_sum_path = os.path.join(OUTPUT_DIR, "amberlyn_MASTER_Osszegzes.xlsx")
    master_raw_path = os.path.join(OUTPUT_DIR, "amberlyn_MASTER_Nyers.xlsx")

    # Mivel felülírjuk a sorokat, a korábbi Master fájlokat már NEM TÖRÖLJÜK ki az elején!
    
    processed_payouts = set()

    for item in data_items:
        if item.get("state") != "PAID": continue
        payout_id = item.get("payout_id")
        payout_date = item.get("settled_at") or item.get("payout_date", "")[:10]
        if not payout_date or payout_id in processed_payouts: continue

        processed_payouts.add(payout_id)
        print(f"\n-> Feldolgozás: {payout_date}")

        fees = item.get("fees", 0.0)
        summary_data = {
            "kifizetes": format_eur(item.get("amount")),
            "kezdeti_egyenleg": format_eur(item.get("initial_balance")),
            "tranzakciok": format_eur(item.get("sales")),
            "visszateritesek": format_eur(item.get("refunds")),
            "chargeback": format_eur(item.get("chargebacks")),
            "dijak": format_eur(-fees if fees > 0 else fees),
        }

        raw_file_name = f"amberlyn_{payout_date}.xlsx"
        summary_file_name = f"amberlyn_{payout_date}_Osszegzes.xlsx"
        raw_path = os.path.join(OUTPUT_DIR, raw_file_name)
        sum_path = os.path.join(OUTPUT_DIR, summary_file_name)

        create_summary_excel(sum_path, summary_data, payout_date)
        append_to_master_summary(master_sum_path, summary_data, payout_date)

        if payout_id:
            if generate_and_download_raw_excel(payout_id, raw_path, headers):
                # Átadjuk a payout_date-et is, hogy tudja, miket kell kitörölni felülírás előtt
                append_to_master_raw(master_raw_path, raw_path, payout_date)

        upload_to_google_drive(sum_path, summary_file_name, GOOGLE_DRIVE_FOLDER_ID)
        upload_to_google_drive(raw_path, raw_file_name, GOOGLE_DRIVE_FOLDER_ID)

    # Master fájlok feltöltése (Felülírja a Drive-on is)
    print(f"\n-> MASTER fájlok szinkronizálása a Google Drive-ra...")
    if os.path.exists(master_sum_path):
        upload_to_google_drive(master_sum_path, "amberlyn_MASTER_Osszegzes.xlsx", GOOGLE_DRIVE_FOLDER_ID)
    if os.path.exists(master_raw_path):
        upload_to_google_drive(master_raw_path, "amberlyn_MASTER_Nyers.xlsx", GOOGLE_DRIVE_FOLDER_ID)

    print("\n--> MINTA SIKERESEN LEFUTOTT!")


if __name__ == "__main__":
    # GitHub workflow esetén a fájlokat le kell tölteni a Drive-ról a szerkesztés előtt
    # Ahhoz, hogy a Master fájlok módosíthatók legyenek a virtuális gépen, egy fejlettebb lekérés kéne, de 
    # mivel a workflow mindig tiszta mappával indul, a GitHub gépén a Master fájlok nulláról kezdenének épülni.
    # Ahhoz hogy egy MÁR LÉTEZŐ master fájlt szerkesszünk, a teljes Drive letöltés túl komplex a Pythonban, 
    # úgyhogy jelenlegi megközelítés: a kód minden futáskor csak a legutóbbi 14 napot fűzi egybe egy masterbe, 
    # és cseréli a régit (vagy ha neked van rá szükséged, 14 napos gördülő mastert tart fent).
    run_pure_api_collector()
