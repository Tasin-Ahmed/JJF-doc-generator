"""
invoice_extractor.py — Extracts structured data from commercial invoice Excel files.

This module is interface-agnostic: it only reads Excel, returns Python dicts.
It does not know about Telegram, Web, or file paths outside of the Excel file itself.
"""

import re
import pandas as pd
from typing import Any, Dict, List, Optional, Tuple


class CommercialInvoiceExtractor:
    """Extract data from commercial invoice Excel files with dynamic column detection."""

    BANKING_FIELDS_CONFIG = {
        'negotiated_with': {
            'label_variations': ['Negotiated With', 'NEGOTIATED WITH', 'Negotiated with'],
            'parse_details': True
        },
        'drawn_on': {
            'label_variations': ['Drawn On', 'DRAWN ON', 'Drawn on'],
            'parse_details': True
        },
        'charges': {
            'label_variations': ['Charges', 'CHARGES', 'Bank Charges'],
            'parse_details': False
        }
    }

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.data = None
        self.pl_details_data = None
        self.invoice_data = {}

    # ── File Loading ──────────────────────────────────────────────────────────

    def load_file(self, sheet_name: str = 'INVOICE') -> pd.DataFrame:
        try:
            if self.file_path.lower().endswith('.xls'):
                engine = 'xlrd'
            elif self.file_path.lower().endswith('.xlsx'):
                engine = 'openpyxl'
            else:
                raise ValueError("Unsupported file format. Please provide an .xls or .xlsx file.")

            df = pd.read_excel(
                self.file_path,
                sheet_name=sheet_name,
                header=None,
                engine=engine
            )

            if sheet_name == 'INVOICE':
                self.data = df
            elif sheet_name == 'PL DETAILS':
                self.pl_details_data = df

            # print(f"✓ Loaded sheet: {sheet_name} ({df.shape[0]} rows x {df.shape[1]} cols)")
            return df

        except Exception as e:
            # print(f"✗ Error loading file: {e}")
            raise

    # ── Cell / Search Helpers ─────────────────────────────────────────────────

    def extract_cell_value(self, row_idx: int, col_idx: int,
                           default: str = "", df: pd.DataFrame = None) -> str:
        if df is None:
            df = self.data
        try:
            value = df.iloc[row_idx, col_idx]
            if pd.isna(value):
                return default
            return str(value).strip()
        except Exception:
            return default

    def find_text_in_sheet(self, search_text: str, column: Optional[int] = None,
                           df: pd.DataFrame = None) -> Optional[Tuple[int, int]]:
        if df is None:
            df = self.data
        for row_idx in range(len(df)):
            cols = [column] if column is not None else range(len(df.columns))
            for col_idx in cols:
                if search_text.lower() in self.extract_cell_value(row_idx, col_idx, df=df).lower():
                    return (row_idx, col_idx)
        return None

    def find_column_by_header(self, header_text: str, header_row: int,
                              search_variations: List[str] = None,
                              df: pd.DataFrame = None) -> Optional[int]:
        if df is None:
            df = self.data
        search_terms = [header_text] + (search_variations or [])
        for col_idx in range(len(df.columns)):
            cell_value = self.extract_cell_value(header_row, col_idx, df=df).lower()
            for term in search_terms:
                if term.lower() in cell_value:
                    return col_idx
        return None

    def extract_text_after_label(self, label: str, row_idx: int, col_idx: int) -> str:
        cell_value = self.extract_cell_value(row_idx, col_idx)
        if label.lower() in cell_value.lower():
            parts = cell_value.split(':', 1)
            if len(parts) > 1:
                return parts[1].strip()
        return cell_value

    # ── Packing List ──────────────────────────────────────────────────────────

    def find_all_packing_list_headers(self, header_row: int,
                                      df: pd.DataFrame) -> List[Dict[str, int]]:
        header_mappings = {
            'pallet_no':    ['Pallet No.', 'Pallet No', 'PALLET NO', 'Pallet'],
            'no_of_spool':  ['No. of Spool', 'No of Spool', 'Spool', 'NO. OF SPOOL'],
            'gross_weight': ['Gross Wt.', 'Gross Weight', 'GROSS WT', 'Gross Wt. (Kgs)'],
            'net_weight':   ['Net Wt.', 'Net Weight', 'NET WT', 'Net Wt. (Kgs)']
        }

        pallet_positions = []
        for col_idx in range(len(df.columns)):
            cell = self.extract_cell_value(header_row, col_idx, df=df).lower()
            if any(v.lower() in cell for v in header_mappings['pallet_no']):
                pallet_positions.append(col_idx)

        # print(f"  ✓ Found {len(pallet_positions)} table instance(s) at columns: {pallet_positions}")

        table_groups = []
        for pallet_col in pallet_positions:
            group = {'pallet_no': pallet_col}
            end = min(pallet_col + 7, len(df.columns))
            for field, variations in header_mappings.items():
                if field == 'pallet_no':
                    continue
                for col_idx in range(pallet_col, end):
                    cell = self.extract_cell_value(header_row, col_idx, df=df).lower()
                    if any(v.lower() in cell for v in variations):
                        group[field] = col_idx
                        break
            if len(group) == 4:
                table_groups.append(group)
                # print(f"    ✓ Table at column {pallet_col}: {group}")
            else:
                print(f"    ⚠ Incomplete table at column {pallet_col}: {group}")

        return table_groups

    def extract_packing_list_details(self) -> Dict[str, Any]:
        # print("\n--- Extracting Packing List Details ---")

        if self.pl_details_data is None:
            self.load_file('PL DETAILS')

        df = self.pl_details_data
        result = {
            'packing_list': [],
            'summary': {
                'total_gross_weight_kgs': 0,
                'total_net_weight_kgs': 0,
                'total_pallets': 0,
                'total_spools': 0
            }
        }

        header_pos = self.find_text_in_sheet("Pallet No", df=df)
        if not header_pos:
            # print("⚠ Could not find packing list table header")
            return result

        header_row = header_pos[0]
        # print(f"✓ Found packing list header at row: {header_row}")

        table_groups = self.find_all_packing_list_headers(header_row, df)
        if not table_groups:
            # print("⚠ No complete table groups found")
            return result

        data_start = header_row + 1
        for table_idx, cols in enumerate(table_groups):
            # print(f"\n  Extracting from table {table_idx + 1}...")
            for offset in range(40):
                row = data_start + offset
                if row >= len(df):
                    break
                pallet_val = self.extract_cell_value(row, cols['pallet_no'], df=df)
                if not pallet_val or "total" in pallet_val.lower():
                    break
                try:
                    record = {
                        'pallet_no': int(float(pallet_val)),
                        'no_of_spool': 0,
                        'gross_weight_kgs': 0.0,
                        'net_weight_kgs': 0.0
                    }
                    if 'no_of_spool' in cols:
                        v = self.extract_cell_value(row, cols['no_of_spool'], df=df)
                        if v:
                            record['no_of_spool'] = int(float(v))
                    if 'gross_weight' in cols:
                        v = self.extract_cell_value(row, cols['gross_weight'], df=df)
                        if v:
                            record['gross_weight_kgs'] = float(v)
                    if 'net_weight' in cols:
                        v = self.extract_cell_value(row, cols['net_weight'], df=df)
                        if v:
                            record['net_weight_kgs'] = float(v)
                    result['packing_list'].append(record)
                    # print(f"    ✓ Pallet {record['pallet_no']}: {record['no_of_spool']} spools")
                except Exception as e:
                    print(f"    ⚠ Row {row}: {e}")

        for label, key in [("Total Gross Weight", "total_gross_weight_kgs"),
                            ("Total Net Weight",  "total_net_weight_kgs")]:
            pos = self.find_text_in_sheet(label, df=df)
            if pos:
                r, c = pos
                for off in range(1, 5):
                    v = self.extract_cell_value(r, c + off, df=df)
                    try:
                        if v and v.lower() != 'kgs':
                            result['summary'][key] = float(v)
                            # print(f"    ✓ {label}: {v} kgs")
                            break
                    except Exception:
                        continue

        result['summary']['total_pallets'] = len(result['packing_list'])
        result['summary']['total_spools']  = sum(i['no_of_spool'] for i in result['packing_list'])

        # print(f"\n✓ Packing list done — {result['summary']['total_pallets']} pallets, "
            #   f"{result['summary']['total_spools']} spools")
        return result

    # ── Banking Details ───────────────────────────────────────────────────────

    def parse_bank_details(self, bank_text: str) -> Dict[str, str]:
        details = {
            'full_text': bank_text, 'bank_name': '', 'branch': '',
            'address': '', 'account_number': '', 'swift_code': '', 'routing_number': ''
        }
        if not bank_text:
            return details

        m = re.search(r'^([^,]+?)(?:,|\s+Branch)', bank_text)
        if m:
            details['bank_name'] = m.group(1).strip()

        m = re.search(r'([^,]*Branch[^,]*)', bank_text, re.IGNORECASE)
        if m:
            details['branch'] = m.group(1).strip()

        m = re.search(r'Branch,\s*(.+?)(?:A/C|Account)', bank_text, re.IGNORECASE)
        if m:
            details['address'] = m.group(1).strip()

        for pattern in [r'A/C\.?\s*No\.?\s*[:\s]*([0-9]+)', r'A/C[:\s]+([0-9]+)',
                        r'Account\s*No\.?\s*[:\s]*([0-9]+)', r'Account[:\s]+([0-9]+)']:
            m = re.search(pattern, bank_text, re.IGNORECASE)
            if m:
                details['account_number'] = m.group(1).strip()
                break

        m = re.search(r'Swift\s*Code[:\s]*([A-Z0-9]+)', bank_text, re.IGNORECASE)
        if m:
            details['swift_code'] = m.group(1).strip()

        m = re.search(r'Routing\s*[Nn]o\.?\s*[:\s]*([0-9]+)', bank_text, re.IGNORECASE)
        if m:
            details['routing_number'] = m.group(1).strip()

        return details

    def extract_banking_details(self, label_column: int = 1,
                                value_column: int = 3) -> Dict[str, Any]:
        banking = {}
        # print("\n--- Extracting Banking Details ---")

        first_field  = list(self.BANKING_FIELDS_CONFIG.keys())[0]
        first_labels = self.BANKING_FIELDS_CONFIG[first_field]['label_variations']
        start_pos    = None
        for label in first_labels:
            start_pos = self.find_text_in_sheet(label, column=label_column)
            if start_pos:
                break

        if not start_pos:
            # print("⚠ Banking details section not found")
            return banking

        current_row = start_pos[0]
        # print(f"✓ Found banking section at row: {current_row}")

        for field_name, cfg in self.BANKING_FIELDS_CONFIG.items():
            found = False
            for offset in range(20):
                row = current_row + offset
                if row >= len(self.data):
                    break
                label = self.extract_cell_value(row, label_column)
                for variation in cfg['label_variations']:
                    if variation.lower() in label.lower():
                        value = self.extract_cell_value(row, value_column)
                        if cfg['parse_details']:
                            parsed = self.parse_bank_details(value)
                            banking[field_name] = parsed
                            # print(f"✓ {field_name}: bank={parsed['bank_name']}, "
                                #   f"account={parsed['account_number']}")
                        else:
                            banking[field_name] = {'full_text': value}
                            # print(f"✓ {field_name}: {value[:50]}")
                        found = True
                        current_row = row + 1
                        break
                if found:
                    break
            if not found:
                # print(f"⚠ Field '{field_name}' not found")
                banking[field_name] = {}

        return banking

    # ── Invoice Sections ──────────────────────────────────────────────────────

    def find_all_column_headers(self, header_row: int) -> Dict[str, int]:
        columns = {}
        mappings = {
            'marks':        ['MARKS & NOS', 'MARKS', 'MARKS AND NOS', 'MARK'],
            'description':  ['DESCRIPTION OF GOODS', 'DESCRIPTION', 'COMMODITY'],
            'quantity':     ['QTY./MT.', 'QTY/MT', 'QUANTITY', 'QTY', 'MT'],
            'unit_price':   ['UNIT PRICE', 'PRICE', 'UNIT PRICE IN USD', 'RATE'],
            'total_amount': ['TOTAL AMOUNT', 'AMOUNT', 'TOTAL', 'TOTAL AMOUNT IN USD']
        }
        for field, variations in mappings.items():
            col = self.find_column_by_header(variations[0], header_row, variations[1:])
            if col is not None:
                columns[field] = col
                # print(f"  ✓ '{field}' at column {col}")
            else:
                print(f"  ⚠ Header '{field}' not found")
        return columns

    def extract_shipper_info(self) -> Dict[str, str]:
        info = {'label': '', 'name': '', 'address_line1': '', 'address_line2': ''}
        pos = self.find_text_in_sheet("SHIPPER/EXPORTER")
        if pos:
            r, c = pos
            info['label']        = self.extract_cell_value(r,     c)
            info['name']         = self.extract_cell_value(r + 1, c)
            info['address_line1']= self.extract_cell_value(r + 2, c)
            info['address_line2']= self.extract_cell_value(r + 3, c)
        return info

    def extract_for_account_info(self) -> Dict[str, str]:
        info = {'label': '', 'name': '', 'address_line1': '', 'address_line2': '', 'license_no': ''}
        pos = self.find_text_in_sheet("For Account and Risk")
        if pos:
            r, c = pos
            info['label']        = self.extract_cell_value(r,     c)
            info['name']         = self.extract_cell_value(r + 1, c)
            info['address_line1']= self.extract_cell_value(r + 2, c)
            info['address_line2']= self.extract_cell_value(r + 3, c)
            lic = self.find_text_in_sheet("LICENSE NO")
            if lic:
                info['license_no'] = self.extract_text_after_label("LICENSE NO", lic[0], lic[1])
        return info

    def extract_notify_party(self) -> Dict[str, str]:
        info = {'label': '', 'name': '', 'address_line1': '', 'address_line2': ''}
        pos = self.find_text_in_sheet("NOTIFY PARTY")
        if pos:
            r, c = pos
            info['label']        = self.extract_cell_value(r,     c)
            info['name']         = self.extract_cell_value(r + 1, c)
            info['address_line1']= self.extract_cell_value(r + 2, c)
            info['address_line2']= self.extract_cell_value(r + 3, c)
        return info

    def extract_invoice_details(self) -> Dict[str, str]:
        details = {
            'invoice_no': '', 'invoice_date': '', 'exp_no': '', 'exp_date': '',
            'lc_sc_no': '', 'lc_sc_date': '', 'bl_no': '', 'bl_date': '',
            'pi_no': '', 'pi_date': '', 'terms_of_delivery': '',
            'hs_code': '', 'country_of_origin': ''
        }

        pos = self.find_text_in_sheet("INVOICE NO")
        if pos:
            text = self.extract_cell_value(pos[0], pos[1] + 2)
            m = re.search(r'([A-Z\-0-9]+)\s+DATED\s+(\d{2}/\d{2}/\d{4})', text)
            if m:
                details['invoice_no'], details['invoice_date'] = m.group(1), m.group(2)

        pos = self.find_text_in_sheet("EXP NO")
        if pos:
            text = self.extract_cell_value(pos[0], pos[1] + 2)
            m = re.search(r'([0-9\-]+)\s+DATED\s+(\d{2}/\d{2}/\d{4})', text)
            if m:
                details['exp_no'], details['exp_date'] = m.group(1), m.group(2)

        pos = self.find_text_in_sheet("LC/SC")
        if pos:
            text = self.extract_cell_value(pos[0], pos[1] + 2)
            m = re.search(r'([A-Z0-9\-]+)\s+DATED[:\s]+(\d{2}/\d{2}/\d{4})', text)
            if m:
                details['lc_sc_no'], details['lc_sc_date'] = m.group(1), m.group(2)

        pos = self.find_text_in_sheet("BILL OF LADING")
        if pos:
            text = self.extract_cell_value(pos[0], pos[1] + 2)
            m = re.search(r'([A-Z0-9]+)', text)
            if m:
                details['bl_no'] = m.group(1)
                dm = re.search(r'DATED[:\s]+(\d{2}/\d{2}/\d{4})', text)
                if dm:
                    details['bl_date'] = dm.group(1)

        pos = self.find_text_in_sheet("PI NO")
        if pos:
            text = self.extract_cell_value(pos[0], pos[1] + 2)
            m = re.search(r'([A-Z0-9\/\-]+)\s+DATED[:\s]+(\d{2}/\d{2}/\d{4})', text, re.IGNORECASE)
            if m:
                details['pi_no'], details['pi_date'] = m.group(1), m.group(2)
            else:
                details['pi_no'] = text

        for label, key in [("TERMS OF DELIVERY", "terms_of_delivery"),
                           ("H.S. CODE NO",       "hs_code"),
                           ("COUNTRY OF ORIGIN",  "country_of_origin")]:
            pos = self.find_text_in_sheet(label)
            if pos:
                details[key] = self.extract_cell_value(pos[0], pos[1] + 2)

        return details

    def extract_port_info(self) -> Dict[str, str]:
        ports = {'port_of_loading': '', 'port_of_discharge': ''}
        for label, key in [("PORT OF LOADING",   "port_of_loading"),
                           ("PORT OF DISCHARGE",  "port_of_discharge")]:
            pos = self.find_text_in_sheet(label)
            if pos:
                ports[key] = self.extract_cell_value(pos[0], pos[1] + 2)
        return ports

    def extract_goods_description(self) -> Dict[str, Any]:
        goods = {
            'marks_and_nos': '', 'description': '',
            'quantity_mt': 0, 'unit_price_usd': 0, 'total_amount_usd': 0
        }
        # print("\n--- Extracting Goods Description ---")

        pos = self.find_text_in_sheet("MARKS & NOS")
        if not pos:
            # print("⚠ Could not find 'MARKS & NOS' header")
            return goods

        header_row = pos[0]
        columns    = self.find_all_column_headers(header_row)
        data_start = header_row + 2

        for offset in range(10):
            row = data_start + offset
            if 'marks' in columns:
                v = self.extract_cell_value(row, columns['marks'])
                if v:
                    goods['marks_and_nos'] = v
            if 'description' in columns:
                v = self.extract_cell_value(row, columns['description'])
                if v:
                    goods['description'] = v
            try:
                qty   = float(self.extract_cell_value(row, columns.get('quantity', -1)))
                price = float(self.extract_cell_value(row, columns.get('unit_price', -1)))
                total = float(self.extract_cell_value(row, columns.get('total_amount', -1)))
                if qty > 0 and total > 0:
                    goods['quantity_mt']      = qty
                    goods['unit_price_usd']   = price
                    goods['total_amount_usd'] = total
                    # print(f"✓ Qty: {qty} MT | Price: ${price} | Total: ${total}")
                    break
            except Exception:
                continue

        return goods

    def extract_amount_in_words(self) -> str:
        pos = self.find_text_in_sheet("USD IN WORD")
        if pos:
            text = self.extract_cell_value(pos[0], pos[1])
            return text.replace("USD IN WORD :", "").replace("USD IN WORD:", "").strip()
        return ""

    def extract_container_info(self) -> List[Dict[str, Any]]:
        containers = []
        # print("\n--- Extracting Container Information ---")

        pos = self.find_text_in_sheet("CONTAINER NO")
        if not pos:
            # print("⚠ Could not find 'CONTAINER NO' header")
            return containers

        header_row = pos[0]
        mappings   = {
            'container_no':   ['CONTAINER NO', 'CONTAINER', 'CONT NO'],
            'seal_no':        ['SEAL NO', 'SEAL', 'SEAL NUMBER'],
            'container_size': ['CONTAINER SIZE', 'SIZE', 'CONT SIZE'],
            'pallets':        ['PALLETS/TRUSS', 'PALLETS', 'TRUSS'],
            'gross_weight':   ['GROSS WT', 'GROSS WEIGHT', 'G.WT'],
            'net_weight':     ['NET WT', 'NET WEIGHT', 'N.WT']
        }
        cols = {}
        for field, variations in mappings.items():
            col = self.find_column_by_header(variations[0], header_row, variations[1:])
            if col is not None:
                cols[field] = col

        row = header_row + 1
        while row < len(self.data):
            if 'container_no' not in cols:
                break
            container_no = self.extract_cell_value(row, cols['container_no'])
            if not container_no or "TOTAL" in container_no.upper():
                break

            c = {
                'container_no':    container_no,
                'seal_no':         self.extract_cell_value(row, cols.get('seal_no', 0)),
                'container_size':  self.extract_cell_value(row, cols.get('container_size', 0)),
                'pallets_truss':   self.extract_cell_value(row, cols.get('pallets', 0)),
                'gross_weight_kg': self.extract_cell_value(row, cols.get('gross_weight', 0)),
                'net_weight_kg':   self.extract_cell_value(row, cols.get('net_weight', 0)),
            }
            for key in ('pallets_truss',):
                try:
                    c[key] = int(float(c[key]))
                except Exception:
                    pass
            for key in ('gross_weight_kg', 'net_weight_kg'):
                try:
                    c[key] = float(c[key])
                except Exception:
                    pass

            containers.append(c)
            # print(f"✓ Container: {container_no}")
            row += 1

        # print(f"✓ Total containers: {len(containers)}")
        return containers

    # ── Main Extract ──────────────────────────────────────────────────────────

    def extract_all_data(self, include_packing_list: bool = True) -> Dict[str, Any]:
        if self.data is None:
            raise ValueError("No data loaded. Call load_file() first.")

        # print("\n" + "=" * 60)
        # print("EXTRACTING INVOICE DATA")
        # print("=" * 60)

        self.invoice_data = {
            'shipper':          self.extract_shipper_info(),
            'for_account_of':   self.extract_for_account_info(),
            'notify_party':     self.extract_notify_party(),
            'invoice_details':  self.extract_invoice_details(),
            'port_info':        self.extract_port_info(),
            'goods':            self.extract_goods_description(),
            'amount_in_words':  self.extract_amount_in_words(),
            'containers':       self.extract_container_info(),
            'banking_details':  self.extract_banking_details(),
        }

        if include_packing_list:
            self.invoice_data['packing_list_details'] = self.extract_packing_list_details()

        # print("\n✓ Extraction complete")
        return self.invoice_data
