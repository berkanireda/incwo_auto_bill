import os
import time
import html
import pandas as pd
import requests
from requests.auth import HTTPBasicAuth
import xml.etree.ElementTree as ET
import urllib3
from dataclasses import dataclass
from typing import Dict, Any, Optional

# On masque les alertes SSL
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

@dataclass
class Config:
    base_url: str
    login: str
    password: str
    business_file_id: int
    throttle_seconds: float = 0.25
    timeout_seconds: int = 60 # Timeout global augmenté
    dry_run: bool = False

def normalize_date_dd_mm_yyyy(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    dt = pd.to_datetime(value, errors="coerce")
    return dt.strftime("%d-%m-%Y") if pd.notna(dt) else str(value).split(" ")[0]

def esc(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return html.escape(str(value), quote=True)

def parse_first_xml_int_tag(xml_text: str, tag_name: str = "id") -> Optional[int]:
    try:
        root = ET.fromstring(xml_text)
        for elem in root.iter():
            if elem.tag.lower().endswith(tag_name.lower()) and elem.text and elem.text.strip().isdigit():
                return int(elem.text.strip())
    except ET.ParseError:
        pass
    return None

def build_bill_sheet_xml(
    row: Dict[str, Any],
    bank_account_id: int,
    billing_address: Dict[str, str],
    product_id: int, 
) -> str:
    firm_id = row.get("firm_id")
    
    # On utilise dynamiquement l'argument product_id
    return f"""<bill_sheet>
  <firm_id>{int(firm_id)}</firm_id>
  <bank_account_id>{int(bank_account_id)}</bank_account_id>
  <billing_date>{esc(normalize_date_dd_mm_yyyy(row.get("billing_date")))}</billing_date>
  <sheet_type>invoice</sheet_type>
  <title>{esc(row.get("title", ""))}</title>
  <billing_add_name>{esc(billing_address.get("name",""))}</billing_add_name>
  <billing_add_street_address>{esc(billing_address.get("street",""))}</billing_add_street_address>
  <billing_add_zip_code>{esc(billing_address.get("zip",""))}</billing_add_zip_code>
  <billing_add_city>{esc(billing_address.get("city",""))}</billing_add_city>
  <billing_add_country>{esc(billing_address.get("country",""))}</billing_add_country>
  <bill_lines>
    <bill_line>
      <customer_product_id>{int(product_id)}</customer_product_id>
      <description>{esc(row.get("description", ""))}</description>
      <unit_price>{esc(row.get("unit_price", 0))}</unit_price>
      <quantity>{esc(row.get("quantity", 1))}</quantity>
      <vat_value_id>{esc(row.get("vat_value_id", ""))}</vat_value_id>
    </bill_line>
  </bill_lines>
</bill_sheet>"""

def post_bill_sheet(cfg: Config, xml_payload: str) -> requests.Response:
    url = f"{cfg.base_url.rstrip('/')}/{cfg.business_file_id}/bill_sheets.xml"
    return requests.post(
        url,
        data=xml_payload.encode("utf-8"),
        headers={"Content-Type": "text/xml", "Accept": "application/xml"},
        auth=HTTPBasicAuth(cfg.login, cfg.password),
        timeout=cfg.timeout_seconds,
        verify=False,
    )

def validate_bill_sheet(cfg: Config, bill_sheet_id: int) -> requests.Response:
    # On utilise l'URL exacte découverte dans l'inspecteur
    # Note : On retire le {business_id} du début car il est placé au milieu dans cette route
    url = f"{cfg.base_url.rstrip('/')}/bill_sheets/confirm_invoice/{cfg.business_file_id}/{bill_sheet_id}"
    
    print(f"--- Validation via GET confirm_invoice pour ID {bill_sheet_id} ---")
    
    # On utilise requests.get car le navigateur a utilisé GET
    return requests.get(
        url,
        auth=HTTPBasicAuth(cfg.login, cfg.password),
        timeout=cfg.timeout_seconds,
        verify=False
    )

def main():
    cfg = Config(
        base_url="https://bo.entreprise-facile.com",
        login="enzolemoine992@gmail.com",
        password="Harlembarksnine#9",
        business_file_id=1050199,
    )

    INPUT_FILE = "res.xlsx"
    BANK_ACCOUNT_ID = 3078309
    PRODUCT_ID = 1000342059 
    
    BILLING_ADDRESS = {
        "name": "CLIENT TEST",
        "street": "12 rue Exemple", "zip": "75001", "city": "Paris", "country": "France",
    }

    df = pd.read_excel(INPUT_FILE) # Plus direct si tu sais que c'est du Excel
    ok, failed = 0, 0

    for idx, row in df.iterrows():
        try:
            xml_payload = build_bill_sheet_xml(row.to_dict(), BANK_ACCOUNT_ID, BILLING_ADDRESS, PRODUCT_ID)
            
            # 1. Création
            resp = post_bill_sheet(cfg, xml_payload)
            if 200 <= resp.status_code < 300:
                bill_id = parse_first_xml_int_tag(resp.text)
                print(f"✅ Créée : ID {bill_id}")
                
                # 2. Validation immédiate
                if bill_id:
                    v = validate_bill_sheet(cfg, bill_id)
                    if v.status_code < 300:
                        print(f"🚀 VALIDÉE : ID {bill_id}")
                        ok += 1
                    else:
                        print(f"⚠️ Erreur Validation ID {bill_id}: {v.status_code}")
                        failed += 1
            else:
                print(f"❌ Erreur Création ligne {idx+1}: {resp.status_code}")
                failed += 1

        except Exception as e:
            print(f"💥 Exception ligne {idx+1}: {e}")
            failed += 1
        
        time.sleep(cfg.throttle_seconds)

    print(f"\nFin du traitement. Succès: {ok} | Échecs: {failed}")

if __name__ == "__main__":
    main()