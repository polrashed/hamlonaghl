# -*- coding: utf-8 -*-
"""
دوستان حمل و نقل | سامانه جامع ممیزی ناوگان، تحلیل داپ و سوخت مسیر
Doustan Transport - Fleet Audit & Route Fuel Analysis System
نسخه ۳.۰ - پیشرفته و بدون باگ
"""

import os
import sys
import re
import math
from datetime import date, timedelta
from collections import defaultdict
from pathlib import Path

# Prevent native crashes on older OS / CPUs
for _m in ('scipy', 'sklearn', 'torch', 'tensorflow', 'cv2'):
    sys.modules[_m] = None

# GUI Imports
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    TK_AVAILABLE = True
except Exception as _tk_err:
    tk = ttk = filedialog = messagebox = None
    TK_AVAILABLE = False

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except Exception:
    HAS_PIL = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

APP_NAME = "دوستان حمل و نقل"
APP_SUBTITLE = "سامانه هوشمند ممیزی بارنامه، صحت‌سنجی داپ و تحلیل سوخت‌گیری مسیر"
APP_VERSION = "3.0.0"

# ----------------- Canonical Jalali Calendar Engine -----------------
def _g2j(gy, gm, gd):
    gdm = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    gy2 = gy + 1 if gm > 2 else gy
    days = 355666 + 365 * gy + ((gy2 + 3) // 4) - ((gy2 + 99) // 100) + ((gy2 + 399) // 400) + gd + gdm[gm - 1]
    jy = -1595 + 33 * (days // 12053)
    days %= 12053
    jy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        jy += (days - 1) // 365
        days = (days - 1) % 365
    if days < 186:
        jm = 1 + days // 31
        jd = 1 + days % 31
    else:
        days -= 186
        jm = 7 + days // 30
        jd = 1 + days % 30
    return jy, jm, jd

def _j2g(jy, jm, jd):
    jy += 1595
    days = -355668 + 365 * jy + (jy // 33) * 8 + ((jy % 33 + 3) // 4) + jd + (jm * 31 if jm < 7 else (jm - 1) * 30 + 6)
    gy = 400 * (days // 146097)
    days %= 146097
    if days > 36524:
        days -= 1
        gy += 100 * (days // 36524)
        days %= 36524
        if days >= 365:
            days += 1
    gy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        gy += (days - 1) // 365
        days = (days - 1) % 365
    gd = days + 1
    sal_a = [0, 31, (29 if (gy % 4 == 0 and (gy % 100 != 0 or gy % 400 == 0)) else 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    gm = 0
    while gm < 13 and gd > sal_a[gm]:
        gd -= sal_a[gm]
        gm += 1
    return gy, gm, gd

def jalali_to_date(jstr):
    if not jstr:
        return None
    s = str(jstr).strip().replace('-', '/').split()[0]
    m = re.match(r'^(\d{4})/(\d{1,2})/(\d{1,2})$', s)
    if not m:
        return None
    try:
        jy, jm, jd = int(m.group(1)), int(m.group(2)), int(m.group(3))
        gy, gm, gd = _j2g(jy, jm, jd)
        return date(gy, gm, gd)
    except Exception:
        return None

def date_to_jalali(d):
    if not d:
        return ""
    jy, jm, jd = _g2j(d.year, d.month, d.day)
    return f"{jy:04d}/{jm:02d}/{jd:02d}"

def normalize_text(s):
    if s is None:
        return ""
    s = str(s).strip()
    s = s.replace('ي', 'ی').replace('ك', 'ک').replace('‌', ' ').replace('ـ', '')
    # Arabic/Persian digits to English digits
    trans = str.maketrans('۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩', '01234567890123456789')
    return s.translate(trans).strip()

def normalize_plate(s):
    if s is None:
        return ""
    s = normalize_text(s)
    s = re.sub(r'[\s\-_/]+', '', s)
    return s

def clean_float(val, default=0.0):
    if val is None:
        return default
    try:
        s = str(val).strip().replace(',', '')
        return float(s)
    except Exception:
        return default

# ----------------- Core Data Engine -----------------
class AuditEngine:
    def __init__(self):
        self.vin_records = []       # list of dict: {vin, pan, plate, plate_series, owner, plate_full}
        self.dap_records = []       # list of dict: {row_id, bill_no, bill_series, trip_date, plate, plate_series, plate_full, origin_city, origin_prov, dest_city, dest_prov, status, route_prov1, route_prov2, distance, driver, mobile, goods_name, company, vehicle_type}
        self.fuel_records = []      # list of dict: {dt_str, date_str, time_str, vin, pan, liters, station, area, zone}
        
        # Lookups
        self.vin_by_plate = {}      # normalized plate_full -> vin
        self.pan_by_vin = {}        # vin -> pan
        self.owner_by_vin = {}      # vin -> owner
        self.plate_by_vin = {}      # vin -> plate_full
        
        # Results
        self.audit_results = []
        self.stats = {
            'total_waybills': 0,
            'total_fuel_tx': 0,
            'total_vehicles': 0,
            'dap_verified_count': 0,
            'fuel_in_route_count': 0,
            'fuel_outside_count': 0,
            'no_fuel_count': 0,
            'total_liters': 0.0
        }

    def load_vin_file(self, filepath):
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        ws = wb.active
        records = []
        
        # Find header row
        header_row_idx = None
        col_map = {}
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
            row_norm = [normalize_text(c) for c in row if c is not None]
            if any('VIN' in str(c).upper() for c in row_norm) or any('پلاک' in str(c) for c in row_norm):
                header_row_idx = row_idx
                for col_idx, cell in enumerate(row):
                    cn = normalize_text(cell).upper()
                    if 'VIN' in cn or 'وین' in cn:
                        col_map['vin'] = col_idx
                    elif 'PAN' in cn or 'کارت' in cn:
                        col_map['pan'] = col_idx
                    elif 'سریال' in cn:
                        col_map['series'] = col_idx
                    elif 'پلاک' in cn and 'سریال' not in cn:
                        col_map['plate'] = col_idx
                    elif any(k in cn for k in ['راننده', 'مالک', 'نام']):
                        col_map['owner'] = col_idx
                break
        
        if header_row_idx is None:
            # Fallback default columns
            col_map = {'owner': 2, 'plate': 3, 'series': 4, 'vin': 5, 'pan': 6}
            start_r = 4
        else:
            start_r = header_row_idx + 1

        for row in ws.iter_rows(min_row=start_r, values_only=True):
            if not row or all(c is None for c in row):
                continue
            vin = normalize_text(row[col_map['vin']]) if 'vin' in col_map and col_map['vin'] < len(row) else ""
            pan = normalize_text(row[col_map['pan']]) if 'pan' in col_map and col_map['pan'] < len(row) else ""
            plate = normalize_text(row[col_map['plate']]) if 'plate' in col_map and col_map['plate'] < len(row) else ""
            series = normalize_text(row[col_map['series']]) if 'series' in col_map and col_map['series'] < len(row) else ""
            owner = normalize_text(row[col_map['owner']]) if 'owner' in col_map and col_map['owner'] < len(row) else ""
            
            if not vin and not plate:
                continue
            
            plate_full = normalize_plate(plate + series)
            rec = {
                'vin': vin,
                'pan': pan,
                'plate': plate,
                'series': series,
                'plate_full': plate_full,
                'owner': owner
            }
            records.append(rec)
            if plate_full and vin:
                self.vin_by_plate[plate_full] = vin
                self.plate_by_vin[vin] = plate_full
            if vin and pan:
                self.pan_by_vin[vin] = pan
            if vin and owner:
                self.owner_by_vin[vin] = owner

        self.vin_records = records
        return len(records)

    def load_dap_file(self, filepath):
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        ws = wb.active
        records = []
        
        header_row_idx = None
        col_map = {}
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=20, values_only=True), start=1):
            row_norm = [normalize_text(c) for c in row if c is not None]
            if any('بارنامه' in str(c) for c in row_norm) and any('پلاک' in str(c) for c in row_norm):
                header_row_idx = row_idx
                for col_idx, cell in enumerate(row):
                    cn = normalize_text(cell)
                    if 'ردیف' in cn:
                        col_map['row_id'] = col_idx
                    elif 'سریال پلاک' in cn:
                        col_map['plate_series'] = col_idx
                    elif 'پلاک' in cn and 'سریال' not in cn:
                        col_map['plate'] = col_idx
                    elif 'شماره بارنامه' in cn or ('بارنامه' in cn and 'سریال' not in cn and 'صحت' not in cn):
                        col_map['bill_no'] = col_idx
                    elif 'سریال بارنامه' in cn:
                        col_map['bill_series'] = col_idx
                    elif 'تاریخ' in cn and 'گزارش' not in cn:
                        col_map['trip_date'] = col_idx
                    elif 'شهر مبدا' in cn or 'مبدا' in cn and 'استان' not in cn:
                        col_map['origin_city'] = col_idx
                    elif 'استان مبدا' in cn:
                        col_map['origin_prov'] = col_idx
                    elif 'شهر مقصد' in cn or 'مقصد' in cn and 'استان' not in cn:
                        col_map['dest_city'] = col_idx
                    elif 'استان مقصد' in cn:
                        col_map['dest_prov'] = col_idx
                    elif 'صحت سنجی' in cn or 'نتیجه' in cn:
                        col_map['status'] = col_idx
                    elif 'مسیر ۱' in cn or 'مسیر 1' in cn:
                        col_map['route_prov1'] = col_idx
                    elif 'مسیر ۲' in cn or 'مسیر 2' in cn:
                        col_map['route_prov2'] = col_idx
                    elif 'مسافت' in cn:
                        col_map['distance'] = col_idx
                    elif 'راننده' in cn and 'شناسه' not in cn and 'موبایل' not in cn:
                        col_map['driver'] = col_idx
                    elif 'موبایل' in cn:
                        col_map['mobile'] = col_idx
                    elif 'شرکت' in cn and 'شناسه' not in cn:
                        col_map['company'] = col_idx
                    elif 'نوع خودرو' in cn:
                        col_map['vehicle_type'] = col_idx
                    elif 'کالا' in cn:
                        col_map['goods_name'] = col_idx
                break
        
        start_r = header_row_idx + 1 if header_row_idx else 11
        
        for row in ws.iter_rows(min_row=start_r, values_only=True):
            if not row or all(c is None for c in row):
                continue
            
            def get_val(key, default=""):
                if key in col_map and col_map[key] < len(row):
                    v = row[col_map[key]]
                    return normalize_text(v) if v is not None else default
                return default
            
            plate = get_val('plate')
            plate_series = get_val('plate_series')
            plate_full = normalize_plate(plate + plate_series)
            bill_no = get_val('bill_no')
            trip_date = get_val('trip_date')
            
            if not plate and not bill_no:
                continue

            rec = {
                'row_id': get_val('row_id'),
                'bill_no': bill_no,
                'bill_series': get_val('bill_series'),
                'trip_date': trip_date,
                'plate': plate,
                'plate_series': plate_series,
                'plate_full': plate_full,
                'origin_city': get_val('origin_city'),
                'origin_prov': get_val('origin_prov'),
                'dest_city': get_val('dest_city'),
                'dest_prov': get_val('dest_prov'),
                'status': get_val('status', 'ثبت شده در داپ'),
                'route_prov1': get_val('route_prov1'),
                'route_prov2': get_val('route_prov2'),
                'distance': clean_float(row[col_map['distance']]) if 'distance' in col_map and col_map['distance'] < len(row) else 0.0,
                'driver': get_val('driver'),
                'mobile': get_val('mobile'),
                'company': get_val('company'),
                'vehicle_type': get_val('vehicle_type'),
                'goods_name': get_val('goods_name')
            }
            records.append(rec)
            
        self.dap_records = records
        return len(records)

    def load_fuel_file(self, filepath):
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
        ws = wb.active
        records = []
        
        header_row_idx = 1
        col_map = {}
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
            row_norm = [normalize_text(c).upper() for c in row if c is not None]
            if any('VIN' in str(c) for c in row_norm) or any('PAN' in str(c) for c in row_norm) or any('SUM' in str(c) for c in row_norm):
                header_row_idx = row_idx
                for col_idx, cell in enumerate(row):
                    cn = normalize_text(cell).upper()
                    if 'VIN' in cn:
                        col_map['vin'] = col_idx
                    elif 'PAN' in cn:
                        col_map['pan'] = col_idx
                    elif 'PERSIAN_D' in cn or 'DATE' in cn or 'تاریخ' in cn:
                        col_map['dt'] = col_idx
                    elif 'SUM_TOTAL' in cn or 'SUM_X1' in cn or 'لیتر' in cn or 'مقدار' in cn:
                        if 'total' not in col_map or 'SUM_TOTAL' in cn:
                            col_map['liters'] = col_idx
                    elif 'GS_NAME' in cn or 'جایگاه' in cn or 'STATION' in cn:
                        col_map['station'] = col_idx
                    elif 'AREA_NAME' in cn or 'منطقه' in cn or 'شهر' in cn:
                        col_map['area'] = col_idx
                    elif 'ZONE_NAME' in cn or 'استان' in cn or 'ناحیه' in cn:
                        col_map['zone'] = col_idx
                break

        start_r = header_row_idx + 1
        for row in ws.iter_rows(min_row=start_r, values_only=True):
            if not row or all(c is None for c in row):
                continue

            def get_val(key, default=""):
                if key in col_map and col_map[key] < len(row):
                    v = row[col_map[key]]
                    return normalize_text(v) if v is not None else default
                return default

            dt_raw = get_val('dt')
            parts = dt_raw.split()
            date_str = parts[0] if len(parts) > 0 else ""
            time_str = parts[1] if len(parts) > 1 else ""
            
            vin = get_val('vin')
            pan = get_val('pan')
            liters = clean_float(row[col_map['liters']]) if 'liters' in col_map and col_map['liters'] < len(row) else 0.0
            
            rec = {
                'dt_str': dt_raw,
                'date_str': date_str,
                'time_str': time_str,
                'vin': vin,
                'pan': pan,
                'liters': liters,
                'station': get_val('station'),
                'area': get_val('area'),
                'zone': get_val('zone')
            }
            records.append(rec)

        self.fuel_records = records
        return len(records)

    def run_comprehensive_audit(self, days_window=3):
        """
        Deep analysis correlating DAP waybills, VIN map, and Fuel Transactions.
        Answers:
        1. Did vehicle traverse the route on waybill date? (DAP validation status)
        2. Was vehicle observed in DAP system in this date or within N-day window?
        3. Did vehicle refuel along the route (Origin, Destination, Route Prov 1, Route Prov 2)?
        4. What was the exact refueling timing, station, province, and liters?
        """
        self.audit_results = []
        
        # Group fuel transactions by VIN and by Date
        fuel_by_vin = defaultdict(list)
        for f in self.fuel_records:
            if f['vin']:
                fuel_by_vin[f['vin']].append(f)
        
        # Counters
        dap_verified_count = 0
        fuel_in_route_count = 0
        fuel_outside_count = 0
        no_fuel_count = 0
        total_liters = 0.0

        for r_idx, dap in enumerate(self.dap_records, start=1):
            plate_full = dap['plate_full']
            vin = self.vin_by_plate.get(plate_full, "")
            vin_status = "تطبیق یافت" if vin else "عدم تطبیق VIN"
            
            w_date_obj = jalali_to_date(dap['trip_date'])
            audit_start_str = dap['trip_date']
            audit_end_str = ""
            
            matched_fuel = []
            fuel_in_corridor = []
            fuel_outside_corridor = []
            
            if w_date_obj:
                end_date_obj = w_date_obj + timedelta(days=days_window)
                audit_end_str = date_to_jalali(end_date_obj)
                
                # Check fuel transactions in window
                candidate_fuels = fuel_by_vin.get(vin, [])
                for f in candidate_fuels:
                    f_date_obj = jalali_to_date(f['date_str'])
                    if f_date_obj and w_date_obj <= f_date_obj <= end_date_obj:
                        matched_fuel.append(f)
            
            # Corridor Provinces
            orig_p = dap['origin_prov']
            dest_p = dap['dest_prov']
            r1_p = dap['route_prov1']
            r2_p = dap['route_prov2']
            
            corridor_provs = set()
            for p in [orig_p, dest_p, r1_p, r2_p]:
                if p and p.strip():
                    for sub in p.split(','):
                        sub_n = normalize_text(sub)
                        if sub_n:
                            corridor_provs.add(sub_n)
            
            for f in matched_fuel:
                f_zone = normalize_text(f['zone'])
                f_area = normalize_text(f['area'])
                # Is zone or area in corridor?
                in_corr = any(cp in f_zone or cp in f_area for cp in corridor_provs) if corridor_provs else False
                if in_corr:
                    fuel_in_corridor.append(f)
                else:
                    fuel_outside_corridor.append(f)

            # Analysis Synthesis
            fuel_count = len(matched_fuel)
            fuel_sum = sum(f['liters'] for f in matched_fuel)
            total_liters += fuel_sum
            
            # Route traversal check
            dap_status = dap['status']
            is_dap_observed = bool(dap_status and "ثبت" in dap_status or "رویت" in dap_status or "تایید" in dap_status)
            if is_dap_observed:
                dap_verified_count += 1
            
            # Fuel in route check
            if len(fuel_in_corridor) > 0:
                route_fuel_verdict = f"سوخت‌گیری در مسیر ({len(fuel_in_corridor)} بار)"
                fuel_in_route_count += 1
            elif len(fuel_outside_corridor) > 0:
                route_fuel_verdict = f"سوخت‌گیری خارج از مسیر ({len(fuel_outside_corridor)} بار)"
                fuel_outside_count += 1
            else:
                route_fuel_verdict = "عدم سوخت‌گیری در بازه"
                no_fuel_count += 1

            # Overall Audit Status
            if is_dap_observed and len(fuel_in_corridor) > 0:
                overall_status = "تایید کامل (داپ + سوخت مسیر)"
                status_color = "#1B5E20" # Green
            elif is_dap_observed and len(matched_fuel) == 0:
                overall_status = "داپ تایید / بدون سوخت"
                status_color = "#E65100" # Orange
            elif "خارج از مسیر" in dap_status:
                overall_status = "هشدار: انحراف از مسیر داپ"
                status_color = "#B71C1C" # Red
            elif not vin:
                overall_status = "فاقد شناسه VIN"
                status_color = "#4A148C" # Purple
            else:
                overall_status = f"وضعیت داپ: {dap_status}"
                status_color = "#0D47A1" # Blue

            fuel_dates_summary = "، ".join([f"{f['date_str']} ({f['liters']} ل - {f['zone']})" for f in matched_fuel[:3]])
            if len(matched_fuel) > 3:
                fuel_dates_summary += f" و {len(matched_fuel)-3} مورد دیگر"

            item = {
                'row_id': r_idx,
                'plate': dap['plate'],
                'plate_series': dap['plate_series'],
                'plate_full': plate_full,
                'bill_no': dap['bill_no'],
                'trip_date': dap['trip_date'],
                'driver': dap['driver'],
                'origin': f"{dap['origin_city']} ({dap['origin_prov']})",
                'dest': f"{dap['dest_city']} ({dap['dest_prov']})",
                'origin_prov': orig_p,
                'dest_prov': dest_p,
                'route_provs': ", ".join(filter(None, [r1_p, r2_p])),
                'distance': dap['distance'],
                'dap_status': dap['status'],
                'vin': vin,
                'vin_status': vin_status,
                'audit_window': f"{audit_start_str} تا {audit_end_str}",
                'fuel_count': fuel_count,
                'fuel_sum': round(fuel_sum, 1),
                'route_fuel_verdict': route_fuel_verdict,
                'overall_status': overall_status,
                'status_color': status_color,
                'fuel_details': fuel_dates_summary,
                'matched_fuel_list': matched_fuel,
                'vehicle_type': dap['vehicle_type'],
                'company': dap['company'],
                'mobile': dap['mobile']
            }
            self.audit_results.append(item)

        self.stats = {
            'total_waybills': len(self.dap_records),
            'total_fuel_tx': len(self.fuel_records),
            'total_vehicles': len(self.vin_records),
            'dap_verified_count': dap_verified_count,
            'fuel_in_route_count': fuel_in_route_count,
            'fuel_outside_count': fuel_outside_count,
            'no_fuel_count': no_fuel_count,
            'total_liters': round(total_liters, 1)
        }
        return self.audit_results

    def export_excel(self, output_path):
        if not HAS_OPENPYXL:
            return False, "کتابخانه openpyxl نصب نیست."
        
        wb = openpyxl.Workbook()
        
        # Styles
        font_header = Font(name="Tahoma", size=10, bold=True, color="FFFFFF")
        font_data = Font(name="Tahoma", size=9)
        fill_header = PatternFill(start_color="1A237E", end_color="1A237E", fill_type="solid")
        fill_alt = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
        fill_green = PatternFill(start_color="E8F5E9", end_color="E8F5E9", fill_type="solid")
        fill_orange = PatternFill(start_color="FFF3E0", end_color="FFF3E0", fill_type="solid")
        fill_red = PatternFill(start_color="FFEBEE", end_color="FFEBEE", fill_type="solid")
        
        thin = Side(border_style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        align_c = Alignment(horizontal="center", vertical="center")
        align_r = Alignment(horizontal="right", vertical="center")

        # Sheet 1: ممیزی تحلیلی جامع
        ws1 = wb.active
        ws1.title = "ممیزی تحلیلی بارنامه و سوخت"
        ws1.views.sheetView[0].rightToLeft = True

        headers1 = [
            "ردیف", "شماره بارنامه", "تاریخ بارنامه", "پلاک خودرو", "سریال", "VIN", 
            "نام راننده", "مبدا", "مقصد", "مسافت (km)", "نتیجه داپ", 
            "بازه تحلیل", "سوخت‌گیری در بازه", "مجموع لیتر", "تحلیل سوخت مسیر", 
            "وضعیت نهایی ممیزی", "جزئیات سوخت‌گیری‌های متناظر"
        ]
        ws1.append(headers1)
        for col_idx, _ in enumerate(headers1, 1):
            cell = ws1.cell(row=1, column=col_idx)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_c

        for r_idx, item in enumerate(self.audit_results, start=2):
            row_data = [
                item['row_id'], item['bill_no'], item['trip_date'], item['plate'], item['plate_series'], item['vin'],
                item['driver'], item['origin'], item['dest'], item['distance'], item['dap_status'],
                item['audit_window'], item['fuel_count'], item['fuel_sum'], item['route_fuel_verdict'],
                item['overall_status'], item['fuel_details']
            ]
            ws1.append(row_data)
            for c_idx in range(1, len(row_data) + 1):
                cell = ws1.cell(row=r_idx, column=c_idx)
                cell.font = font_data
                cell.border = border
                cell.alignment = align_c if c_idx in [1, 2, 3, 4, 5, 6, 10, 12, 13, 14] else align_r
                if "تایید کامل" in item['overall_status']:
                    cell.fill = fill_green
                elif "هشدار" in item['overall_status']:
                    cell.fill = fill_red
                elif "بدون سوخت" in item['overall_status']:
                    cell.fill = fill_orange

        # Adjust Column Widths
        for ws in [ws1]:
            for col in ws.columns:
                max_len = max(len(str(cell.value or '')) for cell in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 11), 40)

        # Sheet 2: ناوگان و VIN
        ws2 = wb.create_sheet(title="ناوگان و VIN")
        ws2.views.sheetView[0].rightToLeft = True
        headers2 = ["ردیف", "نام راننده / مالک", "پلاک", "سریال", "پلاک کامل", "VIN", "PAN کارت سوخت"]
        ws2.append(headers2)
        for col_idx in range(1, len(headers2) + 1):
            c = ws2.cell(row=1, column=col_idx)
            c.font = font_header; c.fill = fill_header; c.alignment = align_c
        for idx, v in enumerate(self.vin_records, start=1):
            ws2.append([idx, v['owner'], v['plate'], v['series'], v['plate_full'], v['vin'], v['pan']])

        # Sheet 3: داپ و بارنامه‌ها
        ws3 = wb.create_sheet(title="بارنامه‌ها و داپ")
        ws3.views.sheetView[0].rightToLeft = True
        headers3 = ["ردیف", "بارنامه", "تاریخ", "پلاک", "مبدا", "مقصد", "مسافت", "وضعیت داپ", "راننده", "مسیر ۱", "مسیر ۲"]
        ws3.append(headers3)
        for col_idx in range(1, len(headers3) + 1):
            c = ws3.cell(row=1, column=col_idx)
            c.font = font_header; c.fill = fill_header; c.alignment = align_c
        for idx, d in enumerate(self.dap_records, start=1):
            ws3.append([idx, d['bill_no'], d['trip_date'], d['plate_full'], d['origin_city'], d['dest_city'], d['distance'], d['status'], d['driver'], d['route_prov1'], d['route_prov2']])

        # Sheet 4: تراکنش‌های سوخت
        ws4 = wb.create_sheet(title="تراکنش‌های سوخت")
        ws4.views.sheetView[0].rightToLeft = True
        headers4 = ["ردیف", "تاریخ و ساعت", "VIN", "PAN", "لیتر", "جایگاه", "منطقه/شهر", "استان/ناحیه"]
        ws4.append(headers4)
        for col_idx in range(1, len(headers4) + 1):
            c = ws4.cell(row=1, column=col_idx)
            c.font = font_header; c.fill = fill_header; c.alignment = align_c
        for idx, f in enumerate(self.fuel_records, start=1):
            ws4.append([idx, f['dt_str'], f['vin'], f['pan'], f['liters'], f['station'], f['area'], f['zone']])

        wb.save(output_path)
        return True, "خروجی با موفقیت ذخیره شد."

# ----------------- Modern Persian Tkinter UI -----------------
class DoustanApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} | نسخه {APP_VERSION}")
        self.root.geometry("1240x820")
        self.root.minsize(1050, 700)
        
        self.engine = AuditEngine()
        self.vin_file_path = tk.StringVar()
        self.dap_file_path = tk.StringVar()
        self.fuel_file_path = tk.StringVar()
        self.audit_window_var = tk.IntVar(value=3)
        self.search_var = tk.StringVar()
        self.filter_status_var = tk.StringVar(value="همه")
        
        self.setup_styles()
        self.build_ui()
        self.auto_load_default_files()

    def setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass

        # Color Palette
        self.BG_MAIN = "#F4F6F9"
        self.BG_CARD = "#FFFFFF"
        self.COLOR_PRIMARY = "#1A237E"    # Deep Indigo
        self.COLOR_ACCENT = "#0288D1"     # Ocean Blue
        self.COLOR_SUCCESS = "#2E7D32"    # Forest Green
        self.COLOR_WARN = "#E65100"       # Orange
        self.COLOR_DANGER = "#C62828"     # Red
        self.TEXT_COLOR = "#212121"

        self.root.configure(bg=self.BG_MAIN)
        
        style.configure("TFrame", background=self.BG_MAIN)
        style.configure("Card.TFrame", background=self.BG_CARD, relief="groove")
        style.configure("TLabel", background=self.BG_MAIN, foreground=self.TEXT_COLOR, font=("Tahoma", 9))
        style.configure("Header.TLabel", font=("Tahoma", 13, "bold"), foreground=self.COLOR_PRIMARY, background=self.BG_MAIN)
        style.configure("SubHeader.TLabel", font=("Tahoma", 9), foreground="#555555", background=self.BG_MAIN)
        
        style.configure("Primary.TButton", font=("Tahoma", 10, "bold"), background=self.COLOR_PRIMARY, foreground="#FFFFFF", borderwidth=1, focuscolor="none", padding=6)
        style.map("Primary.TButton", background=[("active", "#283593")])
        
        style.configure("Accent.TButton", font=("Tahoma", 9, "bold"), background=self.COLOR_ACCENT, foreground="#FFFFFF", padding=5)
        style.map("Accent.TButton", background=[("active", "#039BE5")])
        
        style.configure("Success.TButton", font=("Tahoma", 10, "bold"), background=self.COLOR_SUCCESS, foreground="#FFFFFF", padding=6)
        style.map("Success.TButton", background=[("active", "#388E3C")])

        style.configure("Treeview", font=("Tahoma", 9), rowheight=26, background="#FFFFFF", fieldbackground="#FFFFFF")
        style.configure("Treeview.Heading", font=("Tahoma", 9, "bold"), background="#E0E7FF", foreground=self.COLOR_PRIMARY)
        style.map("Treeview", background=[('selected', '#C5CAE9')], foreground=[('selected', '#000000')])

    def build_ui(self):
        # 1. Top Header Banner
        header_frame = tk.Frame(self.root, bg=self.COLOR_PRIMARY, height=75)
        header_frame.pack(fill="x", side="top")
        
        title_lbl = tk.Label(header_frame, text=APP_NAME, font=("Tahoma", 15, "bold"), fg="#FFFFFF", bg=self.COLOR_PRIMARY)
        title_lbl.pack(anchor="e", padx=25, pady=(10, 0))
        
        sub_lbl = tk.Label(header_frame, text=APP_SUBTITLE, font=("Tahoma", 9), fg="#E0E7FF", bg=self.COLOR_PRIMARY)
        sub_lbl.pack(anchor="e", padx=25, pady=(2, 10))

        # Main Container
        main_box = ttk.Frame(self.root, padding="10")
        main_box.pack(fill="both", expand=True)

        # 2. File Selection & Settings Area
        files_panel = ttk.LabelFrame(main_box, text=" انتخاب و بارگذاری سه فایل اصلی اکسل (VIN، داپ، سوخت) ", padding="10")
        files_panel.pack(fill="x", pady=(0, 8))

        # سه ورودی اصلی، مستقل و برجسته در بالای پنجره
        files_panel.configure(labelanchor="ne")
        files_panel.columnconfigure(2, weight=1)
        self.create_file_selector(files_panel, 0, "فیلد ۱: فایل بارنامه‌های داپ (DAP)", self.dap_file_path, self.load_dap_action, "dap")
        self.create_file_selector(files_panel, 1, "فیلد ۲: فایل تراکنش سوخت (Fuel)", self.fuel_file_path, self.load_fuel_action, "fuel")
        self.create_file_selector(files_panel, 2, "فیلد ۳: فایل نگاشت شاسی و پلاک (VIN)", self.vin_file_path, self.load_vin_action, "vin")

        # Settings & Run Actions Row
        action_bar = ttk.Frame(files_panel)
        action_bar.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(8, 0))

        # Audit Window Setting
        ttk.Label(action_bar, text="بازه بررسی سوخت پس از بارنامه:", font=("Tahoma", 9, "bold")).pack(side="right", padx=(5, 5))
        window_spin = ttk.Spinbox(action_bar, from_=0, to=15, textvariable=self.audit_window_var, width=4, justify="center")
        window_spin.pack(side="right", padx=(0, 15))
        ttk.Label(action_bar, text="روز").pack(side="right", padx=(0, 20))

        # Run Button
        self.btn_run = ttk.Button(action_bar, text=" اجرای ممیزی و تحلیل جامع مسیر ", style="Primary.TButton", command=self.run_audit_action)
        self.btn_run.pack(side="right", padx=10)

        # Export Excel Button
        self.btn_export = ttk.Button(action_bar, text=" دریافت گزارش کامل اکسل (Excel) ", style="Success.TButton", command=self.export_excel_action)
        self.btn_export.pack(side="right", padx=10)

        # Status indicator
        self.lbl_proc_status = ttk.Label(action_bar, text="آماده دریافت فایل‌ها", font=("Tahoma", 9, "bold"), foreground=self.COLOR_ACCENT)
        self.lbl_proc_status.pack(side="left", padx=10)

        # 3. KPI / Summary Cards Frame
        kpi_frame = ttk.Frame(main_box)
        kpi_frame.pack(fill="x", pady=(0, 8))

        self.card_waybills = self.create_kpi_card(kpi_frame, "تعداد کل بارنامه‌ها", "۰", "#1A237E")
        self.card_verified = self.create_kpi_card(kpi_frame, "رویت و تایید داپ", "۰", "#2E7D32")
        self.card_fuel_route = self.create_kpi_card(kpi_frame, "سوخت‌گیری در مسیر", "۰", "#0288D1")
        self.card_fuel_none = self.create_kpi_card(kpi_frame, "بدون سوخت در بازه", "۰", "#E65100")
        self.card_liters = self.create_kpi_card(kpi_frame, "مجموع لیتر سوخت", "۰", "#4A148C")

        # 4. Search and Live Filter Bar
        search_frame = ttk.LabelFrame(main_box, text=" جستجوی پیشرفته لحظه‌ای و فیلترها ", padding="8")
        search_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(search_frame, text="جستجو (پلاک، VIN، بارنامه، راننده، شهر/استان، وضعیت):", font=("Tahoma", 9, "bold")).pack(side="right", padx=5)
        ent_search = ttk.Entry(search_frame, textvariable=self.search_var, width=32, font=("Tahoma", 9))
        ent_search.pack(side="right", padx=5)
        self.search_var.trace_add("write", lambda *args: self.apply_live_filter())

        ttk.Label(search_frame, text="فیلتر وضعیت داپ/سوخت:", font=("Tahoma", 9)).pack(side="right", padx=(20, 5))
        cb_filter = ttk.Combobox(search_frame, textvariable=self.filter_status_var, state="readonly", width=22,
                                 values=["همه", "تایید کامل (داپ + سوخت مسیر)", "سوخت‌گیری در مسیر", "سوخت‌گیری خارج از مسیر", "بدون سوخت در بازه", "انحراف از مسیر داپ", "فاقد شناسه VIN"])
        cb_filter.pack(side="right", padx=5)
        cb_filter.bind("<<ComboboxSelected>>", lambda e: self.apply_live_filter())

        btn_clear_search = ttk.Button(search_frame, text="پاکسازی فیلتر", command=self.clear_filter)
        btn_clear_search.pack(side="right", padx=10)

        # 5. Multi-Tab Notebook with Treeviews
        self.notebook = ttk.Notebook(main_box)
        self.notebook.pack(fill="both", expand=True)

        # Tab 1: Comprehensive Audit
        self.tab_audit = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(self.tab_audit, text="  گزارش ممیزی و تطبیق مسیر و سوخت (تحلیل اصلی)  ")
        self.setup_audit_treeview()

        # Tab 2: Vehicles & VIN Mapping
        self.tab_vin = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(self.tab_vin, text="  اطلاعات ناوگان و VIN  ")
        self.setup_vin_treeview()

        # Tab 3: DAP Waybills
        self.tab_dap = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(self.tab_dap, text="  بارنامه‌ها و سیستم داپ  ")
        self.setup_dap_treeview()

        # Tab 4: Fuel Transactions
        self.tab_fuel = ttk.Frame(self.notebook, padding=5)
        self.notebook.add(self.tab_fuel, text="  تراکنش‌های سوخت  ")
        self.setup_fuel_treeview()

    def create_file_selector(self, parent, row_idx, title, var, cmd, kind):
        """Render one prominent file field with Browse, Clear and live status."""
        lbl = ttk.Label(parent, text=title, font=("Tahoma", 10, "bold"), anchor="e")
        lbl.grid(row=row_idx, column=3, sticky="e", padx=6, pady=5)

        ent = ttk.Entry(parent, textvariable=var, font=("Tahoma", 9), state="readonly")
        ent.grid(row=row_idx, column=2, sticky="ew", padx=5, pady=5)

        browse_btn = ttk.Button(parent, text="Browse", command=cmd)
        browse_btn.grid(row=row_idx, column=1, sticky="ew", padx=4, pady=5)
        clear_btn = ttk.Button(parent, text="Clear", command=lambda k=kind: self.clear_file_selection(k))
        clear_btn.grid(row=row_idx, column=0, sticky="ew", padx=4, pady=5)

        status = ttk.Label(parent, text="Status: فایلی انتخاب نشده", font=("Tahoma", 9, "bold"), foreground="#777777")
        status.grid(row=row_idx, column=4, sticky="w", padx=8, pady=5)
        setattr(self, "lbl_status_" + kind, status)

    def clear_file_selection(self, kind):
        """Clear the selected path and records for only the requested input."""
        path_var = getattr(self, kind + "_file_path")
        path_var.set("")
        if kind == "vin":
            self.engine.vin_records = []
            self.engine.vin_by_plate.clear()
            self.engine.pan_by_vin.clear()
            self.engine.owner_by_vin.clear()
            self.engine.plate_by_vin.clear()
            self.tree_vin.delete(*self.tree_vin.get_children())
        elif kind == "dap":
            self.engine.dap_records = []
            self.tree_dap.delete(*self.tree_dap.get_children())
            self.tree_audit.delete(*self.tree_audit.get_children())
            self.engine.audit_results = []
        elif kind == "fuel":
            self.engine.fuel_records = []
            self.tree_fuel.delete(*self.tree_fuel.get_children())
            self.tree_audit.delete(*self.tree_audit.get_children())
            self.engine.audit_results = []
        getattr(self, "lbl_status_" + kind).config(text="Status: فایلی انتخاب نشده", foreground="#777777")

    def create_kpi_card(self, parent, title, initial_val, color):
        card = tk.Frame(parent, bg="#FFFFFF", highlightbackground="#E0E0E0", highlightthickness=1, bd=0, padx=12, pady=6)
        card.pack(side="right", fill="both", expand=True, padx=4)
        
        lbl_title = tk.Label(card, text=title, font=("Tahoma", 8), fg="#666666", bg="#FFFFFF")
        lbl_title.pack(anchor="center")
        
        lbl_val = tk.Label(card, text=initial_val, font=("Tahoma", 13, "bold"), fg=color, bg="#FFFFFF")
        lbl_val.pack(anchor="center")
        return lbl_val

    def setup_audit_treeview(self):
        cols = (
            "row_id", "bill_no", "trip_date", "plate", "vin", "driver", 
            "origin", "dest", "distance", "dap_status", "fuel_count", 
            "fuel_sum", "route_fuel_verdict", "overall_status", "fuel_details"
        )
        self.tree_audit = ttk.Treeview(self.tab_audit, columns=cols, show="headings", selectmode="browse")
        
        headings = {
            "row_id": "ردیف",
            "bill_no": "شماره بارنامه",
            "trip_date": "تاریخ بارنامه",
            "plate": "پلاک کامل",
            "vin": "VIN",
            "driver": "نام راننده",
            "origin": "مبدا",
            "dest": "مقصد",
            "distance": "مسافت (km)",
            "dap_status": "صحت داپ",
            "fuel_count": "سوخت (تعداد)",
            "fuel_sum": "مجموع لیتر",
            "route_fuel_verdict": "سوخت در مسیر؟",
            "overall_status": "وضعیت نهایی",
            "fuel_details": "جزئیات سوخت‌گیری در بازه"
        }
        
        col_widths = {
            "row_id": 45, "bill_no": 90, "trip_date": 85, "plate": 90, "vin": 140, "driver": 110,
            "origin": 120, "dest": 120, "distance": 70, "dap_status": 120, "fuel_count": 80,
            "fuel_sum": 80, "route_fuel_verdict": 130, "overall_status": 150, "fuel_details": 220
        }
        
        for c in cols:
            self.tree_audit.heading(c, text=headings[c], anchor="center")
            self.tree_audit.column(c, width=col_widths.get(c, 100), anchor="center")

        sb_y = ttk.Scrollbar(self.tab_audit, orient="vertical", command=self.tree_audit.yview)
        sb_x = ttk.Scrollbar(self.tab_audit, orient="horizontal", command=self.tree_audit.xview)
        self.tree_audit.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)

        sb_y.pack(side="left", fill="y")
        sb_x.pack(side="bottom", fill="x")
        self.tree_audit.pack(fill="both", expand=True)
        
        # Tags for colored rows
        self.tree_audit.tag_configure("tag_green", background="#E8F5E9")
        self.tree_audit.tag_configure("tag_orange", background="#FFF3E0")
        self.tree_audit.tag_configure("tag_red", background="#FFEBEE")
        self.tree_audit.tag_configure("tag_purple", background="#F3E5F5")
        
        self.tree_audit.bind("<Double-1>", self.on_audit_double_click)

    def setup_vin_treeview(self):
        cols = ("idx", "owner", "plate", "series", "plate_full", "vin", "pan")
        self.tree_vin = ttk.Treeview(self.tab_vin, columns=cols, show="headings")
        headings = {"idx": "ردیف", "owner": "نام راننده / مالک", "plate": "پلاک", "series": "سریال", "plate_full": "پلاک کامل", "vin": "VIN", "pan": "PAN"}
        for c in cols:
            self.tree_vin.heading(c, text=headings[c], anchor="center")
            self.tree_vin.column(c, width=120, anchor="center")
        sb = ttk.Scrollbar(self.tab_vin, orient="vertical", command=self.tree_vin.yview)
        self.tree_vin.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.tree_vin.pack(fill="both", expand=True)

    def setup_dap_treeview(self):
        cols = ("idx", "bill_no", "trip_date", "plate", "origin", "dest", "distance", "status", "driver")
        self.tree_dap = ttk.Treeview(self.tab_dap, columns=cols, show="headings")
        headings = {"idx": "ردیف", "bill_no": "بارنامه", "trip_date": "تاریخ", "plate": "پلاک", "origin": "مبدا", "dest": "مقصد", "distance": "مسافت", "status": "نتیجه صحت سنجی داپ", "driver": "راننده"}
        for c in cols:
            self.tree_dap.heading(c, text=headings[c], anchor="center")
            self.tree_dap.column(c, width=110, anchor="center")
        sb = ttk.Scrollbar(self.tab_dap, orient="vertical", command=self.tree_dap.yview)
        self.tree_dap.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.tree_dap.pack(fill="both", expand=True)

    def setup_fuel_treeview(self):
        cols = ("idx", "dt", "vin", "pan", "liters", "station", "area", "zone")
        self.tree_fuel = ttk.Treeview(self.tab_fuel, columns=cols, show="headings")
        headings = {"idx": "ردیف", "dt": "تاریخ و ساعت", "vin": "VIN", "pan": "PAN", "liters": "مقدار (لیتر)", "station": "جایگاه", "area": "منطقه", "zone": "استان/ناحیه"}
        for c in cols:
            self.tree_fuel.heading(c, text=headings[c], anchor="center")
            self.tree_fuel.column(c, width=120, anchor="center")
        sb = ttk.Scrollbar(self.tab_fuel, orient="vertical", command=self.tree_fuel.yview)
        self.tree_fuel.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        self.tree_fuel.pack(fill="both", expand=True)

    # ----------------- Actions & Loading -----------------
    def auto_load_default_files(self):
        """Auto detects files in directory or app folder."""
        search_dirs = [os.getcwd(), os.path.dirname(os.path.abspath(__file__)), "/mnt/data"]
        for d in search_dirs:
            if not os.path.isdir(d):
                continue
            for f in os.listdir(d):
                fp = os.path.join(d, f)
                fn = f.lower()
                if ('vin' in fn or 'ناوگان' in fn) and fn.endswith(('.xlsx', '.xls')) and not self.vin_file_path.get():
                    self.vin_file_path.set(fp)
                    self.load_vin_action(fp)
                elif ('داپ' in fn or 'dap' in fn) and fn.endswith(('.xlsx', '.xls')) and not self.dap_file_path.get():
                    self.dap_file_path.set(fp)
                    self.load_dap_action(fp)
                elif ('تراكنش' in fn or 'تراکنش' in fn or 'سوخت' in fn or 'fuel' in fn) and fn.endswith(('.xlsx', '.xls')) and not self.fuel_file_path.get():
                    self.fuel_file_path.set(fp)
                    self.load_fuel_action(fp)

        if self.vin_records_loaded and self.dap_records_loaded and self.fuel_records_loaded:
            self.run_audit_action()

    @property
    def vin_records_loaded(self):
        return len(self.engine.vin_records) > 0

    @property
    def dap_records_loaded(self):
        return len(self.engine.dap_records) > 0

    @property
    def fuel_records_loaded(self):
        return len(self.engine.fuel_records) > 0

    def load_vin_action(self, filepath=None):
        if not filepath:
            filepath = filedialog.askopenfilename(title="انتخاب فایل اکسل VIN و ناوگان", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not filepath or not os.path.isfile(filepath):
            return
        self.vin_file_path.set(filepath)
        try:
            cnt = self.engine.load_vin_file(filepath)
            self.lbl_status_vin.config(text=f"Status: بارگذاری شد ({cnt} خودرو)", foreground=self.COLOR_SUCCESS)
            self.populate_vin_tree()
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در خواندن فایل VIN: {str(e)}")

    def load_dap_action(self, filepath=None):
        if not filepath:
            filepath = filedialog.askopenfilename(title="انتخاب فایل اکسل داپ و بارنامه‌ها", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not filepath or not os.path.isfile(filepath):
            return
        self.dap_file_path.set(filepath)
        try:
            cnt = self.engine.load_dap_file(filepath)
            self.lbl_status_dap.config(text=f"Status: بارگذاری شد ({cnt} بارنامه)", foreground=self.COLOR_SUCCESS)
            self.populate_dap_tree()
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در خواندن فایل داپ: {str(e)}")

    def load_fuel_action(self, filepath=None):
        if not filepath:
            filepath = filedialog.askopenfilename(title="انتخاب فایل اکسل تراکنش‌های سوخت", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not filepath or not os.path.isfile(filepath):
            return
        self.fuel_file_path.set(filepath)
        try:
            cnt = self.engine.load_fuel_file(filepath)
            self.lbl_status_fuel.config(text=f"Status: بارگذاری شد ({cnt} تراکنش)", foreground=self.COLOR_SUCCESS)
            self.populate_fuel_tree()
        except Exception as e:
            messagebox.showerror("خطا", f"خطا در خواندن فایل سوخت: {str(e)}")

    def populate_vin_tree(self):
        self.tree_vin.delete(*self.tree_vin.get_children())
        for idx, r in enumerate(self.engine.vin_records, start=1):
            self.tree_vin.insert("", "end", values=(idx, r['owner'], r['plate'], r['series'], r['plate_full'], r['vin'], r['pan']))

    def populate_dap_tree(self):
        self.tree_dap.delete(*self.tree_dap.get_children())
        for idx, d in enumerate(self.engine.dap_records, start=1):
            self.tree_dap.insert("", "end", values=(idx, d['bill_no'], d['trip_date'], d['plate_full'], d['origin_city'], d['dest_city'], d['distance'], d['status'], d['driver']))

    def populate_fuel_tree(self):
        self.tree_fuel.delete(*self.tree_fuel.get_children())
        for idx, f in enumerate(self.engine.fuel_records[:1000], start=1): # Limit 1000 for speedy display
            self.tree_fuel.insert("", "end", values=(idx, f['dt_str'], f['vin'], f['pan'], f['liters'], f['station'], f['area'], f['zone']))

    def run_audit_action(self):
        if not self.dap_records_loaded:
            messagebox.showwarning("هشدار", "لطفاً ابتدا فایل داپ / بارنامه‌ها را انتخاب و بارگذاری کنید.")
            return
        
        w_days = self.audit_window_var.get()
        self.lbl_proc_status.config(text="در حال تحلیل مسیر و ممیزی سوخت...", foreground=self.COLOR_WARN)
        self.root.update_idletasks()
        
        try:
            results = self.engine.run_comprehensive_audit(days_window=w_days)
            self.update_kpi_display()
            self.apply_live_filter()
            self.lbl_proc_status.config(text=f"✓ ممیزی انجام شد ({len(results)} رکورد)", foreground=self.COLOR_SUCCESS)
        except Exception as e:
            messagebox.showerror("خطای ممیزی", f"خطا در پردازش ممیزی\n{str(e)}")
            self.lbl_proc_status.config(text="خطا در پردازش", foreground=self.COLOR_DANGER)

    def update_kpi_display(self):
        st = self.engine.stats
        self.card_waybills.config(text=f"{st['total_waybills']:,}")
        self.card_verified.config(text=f"{st['dap_verified_count']:,}")
        self.card_fuel_route.config(text=f"{st['fuel_in_route_count']:,}")
        self.card_fuel_none.config(text=f"{st['no_fuel_count']:,}")
        self.card_liters.config(text=f"{st['total_liters']:,.0f} L")

    def apply_live_filter(self):
        q = normalize_text(self.search_var.get()).lower()
        f_status = self.filter_status_var.get()

        self.tree_audit.delete(*self.tree_audit.get_children())
        
        for item in self.engine.audit_results:
            # Status filter
            if f_status != "همه":
                if f_status not in item['overall_status'] and f_status not in item['route_fuel_verdict']:
                    continue
            
            # Text search query
            if q:
                searchable = f"{item['plate_full']} {item['vin']} {item['bill_no']} {item['driver']} {item['origin']} {item['dest']} {item['dap_status']} {item['overall_status']} {item['fuel_details']}".lower()
                if q not in searchable:
                    continue

            tag = "tag_green" if "تایید کامل" in item['overall_status'] else (
                "tag_red" if "هشدار" in item['overall_status'] else (
                    "tag_orange" if "بدون سوخت" in item['overall_status'] else (
                        "tag_purple" if "فاقد" in item['overall_status'] else ""
                    )
                )
            )

            row_vals = (
                item['row_id'], item['bill_no'], item['trip_date'], item['plate_full'],
                item['vin'], item['driver'], item['origin'], item['dest'], item['distance'],
                item['dap_status'], item['fuel_count'], item['fuel_sum'],
                item['route_fuel_verdict'], item['overall_status'], item['fuel_details']
            )
            self.tree_audit.insert("", "end", values=row_vals, tags=(tag,))

    def clear_filter(self):
        self.search_var.set("")
        self.filter_status_var.set("همه")
        self.apply_live_filter()

    def on_audit_double_click(self, event):
        item_id = self.tree_audit.selection()
        if not item_id:
            return
        vals = self.tree_audit.item(item_id, 'values')
        if not vals:
            return
        
        row_id = int(vals[0])
        match = next((x for x in self.engine.audit_results if x['row_id'] == row_id), None)
        if not match:
            return

        # Detail Modal
        win = tk.Toplevel(self.root)
        win.title(f"پرونده ممیزی بارنامه {match['bill_no']} - خودرو {match['plate_full']}")
        win.geometry("680x520")
        win.configure(bg=self.BG_MAIN)
        win.transient(self.root)

        hdr = tk.Label(win, text=f"تحلیل تفصیلی بارنامه {match['bill_no']} ({match['trip_date']})", font=("Tahoma", 11, "bold"), fg=self.COLOR_PRIMARY, bg=self.BG_MAIN)
        hdr.pack(pady=10)

        info_frame = ttk.LabelFrame(win, text=" مشخصات مسیر و بارنامه ", padding=10)
        info_frame.pack(fill="x", padx=15, pady=5)

        info_txt = f"""
پلاک خودرو: {match['plate_full']}        |    شناسه VIN: {match['vin'] or 'ثبت نشده'}
راننده: {match['driver']}             |    موبایل: {match['mobile'] or 'ثبت نشده'}
مسیر: {match['origin']}  به  {match['dest']}  (مسافت: {match['distance']} کیلومتر)
وضعیت ثبت در سیستم داپ: {match['dap_status']}
بازه بررسی سوخت‌گیری: {match['audit_window']}
نتیجه تطبیق سوخت مسیر: {match['route_fuel_verdict']}
وضعیت کل ممیزی: {match['overall_status']}
"""
        tk.Label(info_frame, text=info_txt, font=("Tahoma", 9), justify="right", anchor="e", bg=self.BG_CARD).pack(fill="x")

        fuel_frame = ttk.LabelFrame(win, text=" تراکنش‌های سوخت‌گیری متناظر در بازه زمانی ", padding=10)
        fuel_frame.pack(fill="both", expand=True, padx=15, pady=5)

        tree = ttk.Treeview(fuel_frame, columns=("dt", "liters", "station", "zone", "corridor"), show="headings")
        tree.heading("dt", text="تاریخ و ساعت"); tree.column("dt", width=120, anchor="center")
        tree.heading("liters", text="لیتر"); tree.column("liters", width=70, anchor="center")
        tree.heading("station", text="نام جایگاه"); tree.column("station", width=140, anchor="center")
        tree.heading("zone", text="استان/ناحیه"); tree.column("zone", width=110, anchor="center")
        tree.heading("corridor", text="انطباق با مسیر"); tree.column("corridor", width=100, anchor="center")
        
        tree.pack(fill="both", expand=True)

        for f in match['matched_fuel_list']:
            f_zone = normalize_text(f['zone'])
            f_area = normalize_text(f['area'])
            in_corr = (match['origin_prov'] in f_zone or match['dest_prov'] in f_zone or match['origin_prov'] in f_area or match['dest_prov'] in f_area)
            c_txt = "در مسیر ✓" if in_corr else "خارج مسیر ✗"
            tree.insert("", "end", values=(f['dt_str'], f['liters'], f['station'], f['zone'], c_txt))

        ttk.Button(win, text="بستن", command=win.destroy).pack(pady=8)

    def export_excel_action(self):
        if not self.engine.audit_results:
            messagebox.showwarning("هشدار", "ابتدا عملیات ممیزی را اجرا کنید.")
            return
        
        save_path = filedialog.asksaveasfilename(
            title="ذخیره گزارش جامع اکسل",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile="Doustan_Fleet_Audit_Report.xlsx"
        )
        if not save_path:
            return
        
        ok, msg = self.engine.export_excel(save_path)
        if ok:
            messagebox.showinfo("موفقیت", f"{msg}\nفایل در مسیر زیر ذخیره شد:\n{save_path}")
        else:
            messagebox.showerror("خطا", msg)

def main():
    if not TK_AVAILABLE:
        print("[ERROR] Tkinter is not available.")
        sys.exit(1)
    root = tk.Tk()
    app = DoustanApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
