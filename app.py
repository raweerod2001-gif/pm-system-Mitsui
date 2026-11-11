import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import sqlite3
import hashlib
import os
import io
import base64
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.drawing.image import Image as OpenpyxlImage
import smtplib
from email.message import EmailMessage
import cv2
import numpy as np
from urllib.parse import urlparse, parse_qs

# ==========================================
# PAGE CONFIG & LOGO SETUP
# ==========================================
LOGO_BASE64 = "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"
LOGO_PATH = "Mitsui_High-tec_logo.png"

try:
    if not os.path.exists(LOGO_PATH):
        with open(LOGO_PATH, "wb") as f:
            f.write(base64.b64decode(LOGO_BASE64))
except:
    pass

if os.path.exists(LOGO_PATH):
    st.set_page_config(page_title="MHT Preventive Maintenance", page_icon=LOGO_PATH, layout="wide")
else:
    st.set_page_config(page_title="MHT Preventive Maintenance", page_icon="⚙️", layout="wide")

DB_PATH = "pm_factory_v4.db"

ROLE_LABELS = {
    "technician": "ช่างผู้ตรวจเช็ก",
    "supervisor": "หัวหน้างาน",
    "admin": "ผู้ดูแลระบบ",
}

RADIO_OPTIONS = "OK|Need correction|No have PM item"
ABNORMAL_OPTIONS = "Need correction"

# ==========================================
# SEED DATA (อัปเดต 12 เครื่อง + Auto Stack Lamcor)
# ==========================================
MACHINE_TYPES = [
    {
        "prefix": "UN", "category": "Uncoiler", "name_template": "Uncoiler {code}", "count": 12,
        "items": [
            {"period": "Monthly", "items": [
                "Check noise of Drum turning condition",
                "Check smoothly of dual drum head rotates",
                "Check structure, locked of main shalf & welding point",
                "Check Hydraulic pump and oil level",
                "Check control operation",
                "Check clutch & brake,fixed point of motor",
                "Check all electrical connection and all equipment inside enclosure box must clean & secure",
                "Grease-up dual drum head rotates 1 shot / point",
                "Check oil leak",
                "Check un-coiler close-open  condition"
            ]},
            {"period": "Half-Year", "items": [
                "All item listed under the  monthly maintenance",
                "Re-tighten all Bolt & Nut (Mechanical&Electrical)"
            ]},
            {"period": "Yearly", "items": [
                "All item listed under the half - year mantenance",
                "Re-tighten all Bolt & Nut (Mechanical&Electrical)",
                "Check Resistivity of motor by M-ohm meter"
            ]}
        ]
    },
    {
        "prefix": "LV", "category": "Leveller", "name_template": "Leveller {code}", "count": 12,
        "items": [
            {"period": "Monthly", "items": [
                "All clean",
                "Check motor current……….A. (by INV. Display)",
                "Roller Checking by naked eye must clean and smoothly movement",
                "Check control operation (S-shaped looper)",
                "Check Lubrication oil and grease 1 shot/point",
                "Check rotation roller",
                "Check emergency s/w condition"
            ]},
            {"period": "Half-Year", "items": [
                "All item listed under monthly maintenance",
                "Check sprocket condition",
                "Check roller condition",
                "Check and tighten all electrical terminal"
            ]},
            {"period": "Yearly", "items": [
                "All item listed under the half - year maintenance",
                "Clean &check condition contactor",
                "Retighten all screw,nuts,bolts",
                "Check resistively main motor by M-ohm meter"
            ]}
        ]
    },
    {
        "prefix": "MS", "category": "Press Machine", "name_template": "Press Machine {code}", "count": 12,
        "items": [
            {"period": "Monthly", "items": [
                "Clean auto drain filter of pneumatic (water separator)",
                "Check belt main motor visual check by naked eye",
                "check coupling encoder visual check by naked eye",
                "Check condition magnetic contactor and over load",
                "Visual check wiring terminal, electric equipment inside enclosure panel",
                "Check clutch brake condition by two hand control pedestal (normal working &secure)",
                "Check function top stop of main press at low& Hi-speed must stop at accept zone",
                "Check condition of front and rear guard (Baring slide)",
                "Check operation of safety guard sensor",
                "Check and cleanliness of motor pump",
                "Check oil level (Press, Fedder & MAC unit)",
                "Check leaked point of oil ( Feeder , Press & MAC Unit)",
                "Check operation of oil cooler and clean-up air filter",
                "Grease-up feeder roller",
                "Check condition and Operation function of feeder Unit",
                "Check noise of main press",
                "Check condition Noise enclosure over of machine",
                "Check and clean inverter of main press",
                "Check condition  level sensor cutting oil  and pump",
                "Check condition scrap conveyor",
                "Check machine vibration (M/C bolt and nut stopper for die change)",
                "Check rotation codition,coupling,key,universal joint"
            ]},
            {"period": "Half-Year", "items": [
                "All item list under the monthly maintenance",
                "Check oil condition",
                "Check and operation of die height indicator",
                "Check condition and function of feeder Unit",
                "Check Gab balance of Roller feeder",
                "Check coupling,encoder shaft gap check by hand",
                "Check Torque lock socket pre-roller"
            ]},
            {"period": "Yearly", "items": [
                "All item listed under the half - year maintenance",
                "Replace Hydraulic Oil & Clean-up filter",
                "Replace Feeder Oil & clean-up filter",
                "Check resistively of motor  by M-ohm meter",
                "Check driver skew MAC",
                "Check varicam condition",
                "Check codition air mountting and air brake"
            ]}
        ]
    },
    {
        "prefix": "AS", "category": "Auto Stack", "name_template": "Auto Stack Machine {code}", "count": 17,
        "items": [
            {"period": "Monthly", "items": [
                "All clean",
                "Check condition of Roller",
                "Check cleaniness of  sensor",
                "Check air leak of cylinder",
                "Check control operation",
                "Check condition of clipper",
                "Check emergency s/w condition",
                "Check condition jig jumping",
                "Check function lock button Zero shift"
            ]},
            {"period": "Half-Year", "items": [
                "All item listed under monthly maintenance",
                "Check timing belt condition",
                "Check sensor condition"
            ]},
            {"period": "Yearly", "items": [
                "All item listed under the half - year maintenance",
                "Clean &check condition machine",
                "Calibration sensor",
                "Send machine to QC check MSA",
                "Retighten the nuts and bolts"
            ]}
        ]
    },
    {
        "prefix": "ASLamcor", "category": "Auto Stack", "name_template": "Auto stack Lamcor No.{num}", "count": 3,
        "items": [
            {"period": "Monthly", "items": [
                "All clean",
                "Check condition of Roller",
                "Check cleaniness of  sensor",
                "Check air leak of cylinder",
                "Check control operation",
                "Check condition of clipper",
                "Check emergency s/w condition",
                "Check condition jig jumping",
                "Check function lock button Zero shift"
            ]},
            {"period": "Half-Year", "items": [
                "All item listed under monthly maintenance",
                "Check timing belt condition",
                "Check sensor condition"
            ]},
            {"period": "Yearly", "items": [
                "All item listed under the half - year maintenance",
                "Clean &check condition machine",
                "Calibration sensor",
                "Send machine to QC check MSA",
                "Retighten the nuts and bolts"
            ]}
        ]
    }
]

# ==========================================
# MACHINE MODELS MAPPING
# ==========================================
MACHINE_MODELS = {
    "MS-01": "AMX-3300-230", "MS-02": "HMX-3000 (1)", "MS-03": "HMX-3000 (1)",
    "MS-04": "PDA-8L", "MS-05": "BSTA 80H", "MS-06": "PDA-20L",
    "MS-07": "PDA-20L", "MS-08": "MSP-2200-200", "MS-09": "PDA-20L",
    "MS-10": "MAC-1250-200", "MS-11": "PDA-20L", "MS-12": "HMX-2000",
    
    "LV-01": "RETAIN", "LV-02": "MATSUMOTO", "LV-03": "MATSUMOTO",
    "LV-04": "MITSUI", "LV-05": "MITSUI", "LV-06": "fastechnology",
    "LV-07": "fastechnology", "LV-08": "MATSUMOTO", "LV-09": "AIDA",
    "LV-10": "MATSUMOTO", "LV-11": "fastechnology", "LV-12": "fastechnology",
    
    "UN-01": "MITSUI", "UN-02": "MATSUMOTO", "UN-03": "MATSUMOTO",
    "UN-04": "fastechnology", "UN-05": "SATSUKI KIZAI", "UN-06": "fastechnology",
    "UN-07": "fastechnology", "UN-08": "MATSUMOTO", "UN-09": "MATSUMOTO",
    "UN-10": "YING YU", "UN-11": "MATSUMOTO", "UN-12": "fastechnology"
}

# ==========================================
# DATABASE CONNECTION & INIT
# ==========================================
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn

@st.cache_resource
def init_db_once():
    init_db()
    return True

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL, salt TEXT NOT NULL, full_name TEXT,
        role TEXT NOT NULL DEFAULT 'technician', active INTEGER DEFAULT 1
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS machines (
        id INTEGER PRIMARY KEY AUTOINCREMENT, machine_code TEXT UNIQUE NOT NULL,
        machine_name TEXT NOT NULL, category TEXT, active INTEGER DEFAULT 1
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS checklist_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT, machine_id INTEGER NOT NULL,
        period TEXT NOT NULL, item_name TEXT NOT NULL, item_type TEXT NOT NULL,
        options TEXT, abnormal_options TEXT, order_no INTEGER DEFAULT 0,
        FOREIGN KEY (machine_id) REFERENCES machines(id)
    )""")
    c.execute("""CREATE TABLE IF NOT EXISTS pm_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, machine_id INTEGER,
        machine_code TEXT, machine_name TEXT, inspector_username TEXT, inspector_name TEXT,
        preventive_type TEXT, model TEXT, location TEXT, time_start TEXT, time_finish TEXT,
        part_changing TEXT, note TEXT, image_data TEXT, status TEXT,
        is_approved INTEGER DEFAULT 0, approved_by TEXT, approved_at TEXT
    )""")
    
    try:
        c.execute("ALTER TABLE pm_logs ADD COLUMN work_week TEXT")
    except sqlite3.OperationalError:
        pass
        
    c.execute("""CREATE TABLE IF NOT EXISTS pm_log_details (
        id INTEGER PRIMARY KEY AUTOINCREMENT, log_id INTEGER, period TEXT,
        item_name TEXT, item_type TEXT, value TEXT, remark TEXT, is_abnormal INTEGER,
        FOREIGN KEY (log_id) REFERENCES pm_logs(id)
    )""")
    c.execute("CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)")
    conn.commit()

    c.execute("CREATE INDEX IF NOT EXISTS idx_machines_code ON machines(machine_code)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_checklist_machine ON checklist_items(machine_id)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_logs_machine_code ON pm_logs(machine_code)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_logs_status ON pm_logs(status)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_logs_approved ON pm_logs(is_approved)")
    c.execute("CREATE INDEX IF NOT EXISTS idx_details_logid ON pm_log_details(log_id)")
    conn.commit()

    c.execute("SELECT COUNT(*) as cnt FROM users")
    if c.fetchone()["cnt"] == 0:
        salt, pw_hash = hash_password("admin1234")
        c.execute("INSERT INTO users (username, password_hash, salt, full_name, role) VALUES (?,?,?,?,?)",
                  ("admin", pw_hash, salt, "ผู้ดูแลระบบ", "admin"))
        conn.commit()

    sync_machines(conn)
    conn.close()

def sync_machines(conn):
    c = conn.cursor()
    for mtype in MACHINE_TYPES:
        for i in range(1, mtype.get("count", 12) + 1):
            code = f"{mtype['prefix']}-{i:02d}"
            name = mtype["name_template"].format(code=code, num=i)
            
            c.execute("SELECT id FROM machines WHERE machine_code=?", (code,))
            row = c.fetchone()
            
            if not row:
                c.execute("INSERT INTO machines (machine_code, machine_name, category) VALUES (?,?,?)",
                          (code, name, mtype["category"]))
                machine_id = c.lastrowid
            else:
                machine_id = row["id"]
            
            c.execute("DELETE FROM checklist_items WHERE machine_id=?", (machine_id,))
            
            for p_idx, period_group in enumerate(mtype["items"]):
                period = period_group["period"]
                for order_no, item_name in enumerate(period_group["items"]):
                    c.execute("""INSERT INTO checklist_items 
                                 (machine_id, period, item_name, item_type, options, abnormal_options, order_no)
                                 VALUES (?,?,?,?,?,?,?)""",
                              (machine_id, period, item_name, "radio", RADIO_OPTIONS, ABNORMAL_OPTIONS, order_no + (p_idx*100)))
    conn.commit()

# ==========================================
# SETTINGS & EMAIL HELPERS
# ==========================================
def get_setting(key, default=""):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT value FROM settings WHERE key=?", (key,))
    row = c.fetchone()
    conn.close()
    return row["value"] if row else default

def set_setting(key, value):
    conn = get_conn()
    c = conn.cursor()
    c.execute("INSERT INTO settings (key, value) VALUES (?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value", (key, value))
    conn.commit()
    conn.close()

def send_email_with_excel(excel_bytes, machine_code, pm_type, date_str):
    sender_email = get_setting("smtp_email")
    sender_password = get_setting("smtp_password")
    receiver_email = get_setting("manager_email")
    
    if not sender_email or not sender_password or not receiver_email:
        return False, "ยังไม่ได้ตั้งค่าอีเมลในระบบ (กรุณาให้ Admin ตั้งค่า)"

    try:
        msg = EmailMessage()
        msg['Subject'] = f'[MHT PM System] รายงานตรวจสอบเครื่องจักร {machine_code} ({pm_type})'
        msg['From'] = sender_email
        msg['To'] = receiver_email
        
        body = f"""เรียน ผู้จัดการ (Manager),

ระบบได้ทำการบันทึกและตรวจสอบ (Review) การบำรุงรักษาเครื่องจักรเรียบร้อยแล้ว
รายละเอียดเบื้องต้น:
- รหัสเครื่องจักร: {machine_code}
- รอบการตรวจสอบ: {pm_type}
- วันที่: {date_str}

ได้ทำการแนบไฟล์รายงาน Excel ฉบับสมบูรณ์มาพร้อมกับอีเมลฉบับนี้

ด้วยความเคารพ,
MHT PM System
"""
        msg.set_content(body)
        
        filename = f"PM_Report_{machine_code}_{date_str.replace('-', '')}.xlsx"
        msg.add_attachment(excel_bytes, maintype='application', 
                           subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                           filename=filename)

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            
        return True, "ส่งอีเมลสำเร็จ"
    except Exception as e:
        return False, str(e)

# ==========================================
# USER & MACHINE MANAGEMENT HELPERS
# ==========================================
def hash_password(password, salt=None):
    if salt is None: salt = os.urandom(16).hex()
    pw_hash = hashlib.pbkdf2_hmac("sha256", password.encode(), bytes.fromhex(salt), 100000).hex()
    return salt, pw_hash

def verify_password(password, salt, pw_hash):
    _, check_hash = hash_password(password, salt)
    return check_hash == pw_hash

def authenticate(username, password):
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM users WHERE username=? AND active=1", (username,))
    row = c.fetchone()
    conn.close()
    if row and verify_password(password, row["salt"], row["password_hash"]):
        return dict(row)
    return None

@st.cache_data(ttl=60)
def list_users():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT id, username, full_name, role, active FROM users ORDER BY username")
    rows = c.fetchall()
    conn.close()
    return [dict(r) for r in rows]

def add_user(username, password, full_name, role):
    salt, pw_hash = hash_password(password)
    conn = get_conn()
    c = conn.cursor()
    c.execute("INSERT INTO users (username, password_hash, salt, full_name, role) VALUES (?,?,?,?,?)",
              (username, pw_hash, salt, full_name, role))
    conn.commit()
    conn.close()
    st.cache_data.clear()

def toggle_user_active(user_id, current_active):
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE users SET active=? WHERE id=?", (0 if current_active else 1, user_id))
    conn.commit()
    conn.close()
    st.cache_data.clear()

def delete_user(user_id):
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM users WHERE id=?", (user_id,))
    conn.commit()
    conn.close()
    st.cache_data.clear()

def reset_user_password(user_id, new_password):
    salt, pw_hash = hash_password(new_password)
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE users SET password_hash=?, salt=? WHERE id=?", (pw_hash, salt, user_id))
    conn.commit()
    conn.close()

@st.cache_data(ttl=60)
def get_all_machines():
    conn = get_conn()
    c = conn.cursor()
    c.execute("SELECT * FROM machines ORDER BY machine_code")
    rows = c.fetchall()
    conn.close()
    return [dict(r) for r in rows]

@st.cache_data(ttl=60)
def load_active_machines_df():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM machines WHERE active=1 ORDER BY machine_code", conn)
    conn.close()
    return df

@st.cache_data(ttl=60)
def load_checklist_items_df(machine_id):
    conn = get_conn()
    df = pd.read_sql_query(
        "SELECT * FROM checklist_items WHERE machine_id=? ORDER BY order_no", conn, params=(machine_id,)
    )
    conn.close()
    return df

def toggle_machine_active(machine_id, current_active):
    conn = get_conn()
    c = conn.cursor()
    c.execute("UPDATE machines SET active=? WHERE id=?", (0 if current_active else 1, machine_id))
    conn.commit()
    conn.close()
    st.cache_data.clear()

def delete_pm_log(log_id):
    conn = get_conn()
    c = conn.cursor()
    c.execute("DELETE FROM pm_log_details WHERE log_id=?", (log_id,))
    c.execute("DELETE FROM pm_logs WHERE id=?", (log_id,))
    conn.commit()
    conn.close()

# ==========================================
# EXCEL GENERATOR (แก้โคลอนตรงกันเป๊ะ 100%)
# ==========================================
@st.cache_data
def generate_paper_like_excel(df_logs, df_details):
    wb = Workbook()
    
    ws_summary = wb.active
    ws_summary.title = "Summary (ภาพรวม)"
    headers = list(df_logs.columns)
    ws_summary.append(headers)
    for _, row in df_logs.iterrows():
        ws_summary.append(list(row))
    for cell in ws_summary[1]:
        cell.font = Font(bold=True)
        
    border_thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center', wrap_text=True)
    font_bold = Font(bold=True)
    
    for _, log in df_logs.iterrows():
        sheet_name = f"{log['machine_code']}_{log['timestamp'][:10].replace('-','')}"[:31]
        ws = wb.create_sheet(title=sheet_name)
        
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
        ws.sheet_properties.pageSetUpPr.fitToPage = True  
        ws.page_setup.fitToWidth = 1  
        ws.page_setup.fitToHeight = 1  
        
        ws.page_margins.left = 0.25
        ws.page_margins.right = 0.25
        ws.page_margins.top = 0.25     
        ws.page_margins.bottom = 0.25  
        ws.page_margins.header = 0.0
        ws.page_margins.footer = 0.0
        
        # --- ใช้ตาราง 6 คอลัมน์ (A-F) ---
        ws.column_dimensions['A'].width = 7.5  
        ws.column_dimensions['B'].width = 16   
        ws.column_dimensions['C'].width = 32   
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 15
        
        ws.row_dimensions[1].height = 40  
        
        try:
            img_io = io.BytesIO(base64.b64decode(LOGO_BASE64))
            img = OpenpyxlImage(img_io)
            img.width = 45   
            img.height = 45  
            ws.add_image(img, 'A1')
        except Exception:
            ws['A1'] = "MHT"
            ws['A1'].font = Font(bold=True, size=16, color="000000")
            ws['A1'].alignment = align_center

        ws.merge_cells('B1:F1')
        ws['B1'] = " Mitsui High-Tec (Thailand) Co., Ltd."
        ws['B1'].font = Font(name='Times New Roman', bold=True, size=16)
        ws['B1'].alignment = Alignment(horizontal='left', vertical='center')
        
        ws.merge_cells('A3:F3')
        machine_name_str = str(log['machine_name']).lower()
        machine_cat = machine_name_str.split(' ')[0].capitalize()
        if "press" in machine_name_str: machine_cat = "Press Machine"
        elif "auto stack" in machine_name_str or "as-" in str(log['machine_code']).lower() or "aslamcor" in str(log['machine_code']).lower(): machine_cat = "auto stack machine"
        ws['A3'] = f"Preventive maintenance (PM) check list for {machine_cat}"
        ws['A3'].font = Font(bold=True, size=14)
        ws['A3'].alignment = align_center
        
        # --- จัดระเบียบช่องส่วนหัวใหม่ ให้ข้อมูลคอลัมน์ C และ E ตรงกันเป๊ะ ---
        ws.merge_cells('A5:B5')
        ws['A5'] = "Type of Preventive"
        ws['C5'] = f": {log['preventive_type']}"
        ws['D5'] = "Machine Code"
        ws.merge_cells('E5:F5')
        ws['E5'] = f": {log['machine_code']}"
        
        ws.merge_cells('A6:B6')
        ws['A6'] = "Model"
        ws['C6'] = f": {log['model'] or '-'}"
        ws['D6'] = "Machine Location"
        ws.merge_cells('E6:F6')
        ws['E6'] = f": {log['location'] or '-'}"
        
        ws.merge_cells('A7:B7')
        ws['A7'] = "Preventive date"
        ws['C7'] = f": {log['timestamp'][:10]}"
        ws['D7'] = "Completed date"
        ws.merge_cells('E7:F7')
        ws['E7'] = f": {log['timestamp'][:10]}"
        
        ws.merge_cells('A8:B8')
        ws['A8'] = "Time start"
        ws['C8'] = f": {log['time_start']}"
        ws['D8'] = "Time Finish"
        ws.merge_cells('E8:F8')
        ws['E8'] = f": {log['time_finish']}"
        
        ws['D9'] = "Work week"
        ws.merge_cells('E9:F9')
        work_week_val = log['work_week'] if 'work_week' in log.keys() and log['work_week'] else '-'
        ws['E9'] = f": {work_week_val}"
        
        ws.merge_cells('A11:F11')
        ws['A11'] = "[ / ] OK                      [ X ] Need correction                      [ 0 ] No have PM item"
        ws['A11'].alignment = align_center
        ws['A11'].font = font_bold
        
        ws['A13'] = "Item"
        ws.merge_cells('B13:C13')
        ws['B13'] = "Description"
        ws['D13'] = "Condition"
        ws.merge_cells('E13:F13')
        ws['E13'] = "Remark"
        for col in ['A','B','C','D','E','F']:
            ws[f'{col}13'].font = font_bold
            ws[f'{col}13'].alignment = align_center
            ws[f'{col}13'].border = border_thin
            
        current_row = 14
        log_dets = df_details[df_details['log_id'] == log['id']]
        
        for period in ['Monthly', 'Half-Year', 'Yearly']:
            period_dets = log_dets[log_dets['period'] == period]
            if period_dets.empty: continue
            
            is_checked = '/' if log['preventive_type'] in [period, 'Half-Year', 'Yearly'] else ' '
            
            # --- ใส่กรอบให้ครบทุกช่อง "ก่อน" สั่ง Merge เพื่อป้องกันเส้นหาย ---
            for c_idx in range(1, 7):
                ws.cell(row=current_row, column=c_idx).border = border_thin
                
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)
            ws.cell(row=current_row, column=2).value = f"[ {is_checked} ] {period}"
            ws.cell(row=current_row, column=2).font = font_bold
            ws.cell(row=current_row, column=2).alignment = align_left
            
            ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=6)

            current_row += 1
            
            item_no = 1
            for _, det in period_dets.iterrows():
                # --- ใส่กรอบให้ครบทุกช่อง "ก่อน" สั่ง Merge เสมอ ---
                for c_idx in range(1, 7):
                    ws.cell(row=current_row, column=c_idx).border = border_thin
                    
                ws.cell(row=current_row, column=1).value = item_no
                ws.cell(row=current_row, column=1).alignment = align_center
                
                ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=3)
                ws.cell(row=current_row, column=2).value = det['item_name']
                ws.cell(row=current_row, column=2).alignment = align_left
                
                sym = det['value']
                if sym == "OK": sym = "/"
                elif sym == "Need correction": sym = "X"
                elif sym == "No have PM item": sym = "0"
                
                ws.cell(row=current_row, column=4).value = sym
                ws.cell(row=current_row, column=4).alignment = align_center
                if sym == "X": ws.cell(row=current_row, column=4).font = Font(color="FF0000", bold=True)
                
                ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=6)
                ws.cell(row=current_row, column=5).value = det['remark']
                ws.cell(row=current_row, column=5).alignment = align_left
                
                current_row += 1
                item_no += 1
                
        current_row += 1
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
        ws.cell(row=current_row, column=1).value = "Part changing :"
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=6)
        ws.cell(row=current_row, column=3).value = log['part_changing'] or ""
        for c in range(3, 7):
            ws.cell(row=current_row, column=c).border = Border(bottom=Side(style='thin'))
        
        current_row += 1
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=6)
        for c in range(3, 7):
            ws.cell(row=current_row, column=c).border = Border(bottom=Side(style='thin'))

        current_row += 1
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
        ws.cell(row=current_row, column=1).value = "Note :"
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=6)
        ws.cell(row=current_row, column=3).value = log['note'] or ""
        for c in range(3, 7):
            ws.cell(row=current_row, column=c).border = Border(bottom=Side(style='thin'))
        
        current_row += 2
        
        is_auto_stack = "auto stack" in machine_name_str or "as-" in str(log['machine_code']).lower() or "aslamcor" in str(log['machine_code']).lower()
        is_press = "press" in machine_name_str or "ms-" in str(log['machine_code']).lower()
        is_uncoiler = "uncoiler" in machine_name_str or "un-" in str(log['machine_code']).lower()
        is_leveller = "leveller" in machine_name_str or "lv-" in str(log['machine_code']).lower()
        
        if is_press:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            ws.cell(row=current_row, column=1).value = "Recommended"
            ws.cell(row=current_row, column=1).font = Font(size=9)
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=3)
            ws.cell(row=current_row, column=3).value = "     * Grease         :  Grease no.0"
            ws.cell(row=current_row, column=3).font = Font(size=9)
            
            ws.cell(row=current_row, column=4).value = "Preventive by"
            ws.cell(row=current_row, column=5).value = "Review by"
            ws.cell(row=current_row, column=6).value = "Approved by"
            for c in [4,5,6]:
                ws.cell(row=current_row, column=c).alignment = align_center
                ws.cell(row=current_row, column=c).border = border_thin
                
            current_row += 1
            
            # --- ให้ช่องน้ำมันตัดขึ้นบรรทัดใหม่ ไม่ล้นทับกรอบลายเซ็น ---
            text_oil = "     * Machine oil      :JIS No.1 , (MS-01,02,03,05,08,10,12 Lubrcation oil # 32 ,\n                                       MS-04,06,07,09,11 Lubrication oil # 100)"
            ws.cell(row=current_row, column=3).value = text_oil
            ws.cell(row=current_row, column=3).font = Font(size=9)
            ws.cell(row=current_row, column=3).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row+1, end_column=3)
            ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row+1, end_column=4)
            ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row+1, end_column=5)
            ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row+1, end_column=6)
            
            ws.cell(row=current_row, column=4).value = log['inspector_name']
            ws.cell(row=current_row, column=5).value = log['approved_by'] if log['is_approved'] else ""
            ws.cell(row=current_row, column=6).value = ""
            
            for c in [4,5,6]:
                ws.cell(row=current_row, column=c).alignment = align_center
                ws.cell(row=current_row, column=c).border = border_thin
                ws.cell(row=current_row+1, column=c).border = border_thin

            current_row += 2
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=3)
            ws.cell(row=current_row, column=3).value = "     * MAC Unit oil    :Hydraulic oil  # 32"
            ws.cell(row=current_row, column=3).font = Font(size=9)
            
        elif is_uncoiler:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            ws.cell(row=current_row, column=1).value = "Recommended"
            ws.cell(row=current_row, column=1).font = Font(size=9)
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=3)
            ws.cell(row=current_row, column=3).value = "     * Grease                 :  Grease no.0"
            ws.cell(row=current_row, column=3).font = Font(size=9)
            
            ws.cell(row=current_row, column=4).value = "Preventive by"
            ws.cell(row=current_row, column=5).value = "Review by"
            ws.cell(row=current_row, column=6).value = "Approved by"
            for c in [4,5,6]:
                ws.cell(row=current_row, column=c).alignment = align_center
                ws.cell(row=current_row, column=c).border = border_thin
                
            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=3)
            ws.cell(row=current_row, column=3).value = "1. Confirm  oil seal damage"
            ws.cell(row=current_row, column=3).font = Font(size=9)
            
            ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row+3, end_column=4)
            ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row+3, end_column=5)
            ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row+3, end_column=6)
            
            ws.cell(row=current_row, column=4).value = log['inspector_name']
            ws.cell(row=current_row, column=5).value = log['approved_by'] if log['is_approved'] else ""
            ws.cell(row=current_row, column=6).value = ""
            
            for c in [4,5,6]:
                ws.cell(row=current_row, column=c).alignment = align_center
                for r_idx in range(current_row, current_row+4):
                    ws.cell(row=r_idx, column=c).border = border_thin

            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=3)
            ws.cell(row=current_row, column=3).value = "2. Confirm bearing damage "
            ws.cell(row=current_row, column=3).font = Font(size=9)
            
            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=3)
            ws.cell(row=current_row, column=3).value = "3. Confirm dimension cylinder wear off"
            ws.cell(row=current_row, column=3).font = Font(size=9)
            
            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=3)
            ws.cell(row=current_row, column=3).value = "4. Confirm hydraulic tube wear off"
            ws.cell(row=current_row, column=3).font = Font(size=9)

        elif is_leveller:
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
            ws.cell(row=current_row, column=1).value = "Recommended"
            ws.cell(row=current_row, column=1).font = Font(size=9)
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=3)
            ws.cell(row=current_row, column=3).value = "     * Grease                 :  Grease no.0"
            ws.cell(row=current_row, column=3).font = Font(size=9)
            
            ws.cell(row=current_row, column=4).value = "Preventive by"
            ws.cell(row=current_row, column=5).value = "Review by"
            ws.cell(row=current_row, column=6).value = "Approved by"
            for c in [4,5,6]:
                ws.cell(row=current_row, column=c).alignment = align_center
                ws.cell(row=current_row, column=c).border = border_thin
                
            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row+1, end_column=4)
            ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row+1, end_column=5)
            ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row+1, end_column=6)
            
            ws.cell(row=current_row, column=4).value = log['inspector_name']
            ws.cell(row=current_row, column=5).value = log['approved_by'] if log['is_approved'] else ""
            ws.cell(row=current_row, column=6).value = ""
            
            for c in [4,5,6]:
                ws.cell(row=current_row, column=c).alignment = align_center
                ws.cell(row=current_row, column=c).border = border_thin
                ws.cell(row=current_row+1, column=c).border = border_thin
            current_row += 1

        else:
            # โซน Auto Stack
            ws.cell(row=current_row, column=4).value = "Preventive by"
            ws.cell(row=current_row, column=5).value = "Review by"
            ws.cell(row=current_row, column=6).value = "Approved by"
            for c in [4,5,6]:
                ws.cell(row=current_row, column=c).alignment = align_center
                ws.cell(row=current_row, column=c).border = border_thin
                
            current_row += 1
            ws.merge_cells(start_row=current_row, start_column=4, end_row=current_row+1, end_column=4)
            ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row+1, end_column=5)
            ws.merge_cells(start_row=current_row, start_column=6, end_row=current_row+1, end_column=6)
            
            ws.cell(row=current_row, column=4).value = log['inspector_name']
            ws.cell(row=current_row, column=5).value = log['approved_by'] if log['is_approved'] else ""
            ws.cell(row=current_row, column=6).value = ""
            
            for c in [4,5,6]:
                ws.cell(row=current_row, column=c).alignment = align_center
                ws.cell(row=current_row, column=c).border = border_thin
                ws.cell(row=current_row+1, column=c).border = border_thin
                
            current_row += 1
        
        current_row += 2
        
        doc_info = "QFR-MCS-055, Rev.17"
        eff_date = "27 Feb 2026"
        if is_auto_stack:
            doc_info = "QFR-MCS-063, Rev.04"
            eff_date = "27 Feb 2026"
        elif is_leveller:
            doc_info = "QFR-MCS-054, Rev.04"
            eff_date = "25 Jun 2025"
        elif is_uncoiler:
            doc_info = "QFR-MCS-056, Rev.04"
            eff_date = "25 Jun 2025"
            
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=2)
        ws.cell(row=current_row, column=1).value = f"Effective Date : {eff_date}"
        ws.cell(row=current_row, column=1).font = Font(size=9)
        ws.merge_cells(start_row=current_row, start_column=5, end_row=current_row, end_column=6)
        ws.cell(row=current_row, column=5).value = doc_info
        ws.cell(row=current_row, column=5).font = Font(size=9)
        ws.cell(row=current_row, column=5).alignment = align_right

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


# ==========================================
# PAGE: LOGIN
# ==========================================
def login_page():
    try:
        st.image(base64.b64decode(LOGO_BASE64), width=80)
    except:
        pass
        
    st.title("MHT Preventive Maintenance")
    st.subheader("เข้าสู่ระบบ")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        if st.form_submit_button("เข้าสู่ระบบ"):
            user = authenticate(username, password)
            if user:
                st.session_state.user = user
                st.rerun()
            else:
                st.error("❌ ข้อมูลไม่ถูกต้อง")

# ==========================================
# PAGE: PM FORM (ปรับหน้าตาฟอร์มให้แสดงทั้งหมด + มี Smart Default + Auto-fill Model)
# ==========================================
def pm_form_page(user):
    st.markdown("<h2 style='text-align: center;'>Preventive Maintenance (PM) Check List</h2>", unsafe_allow_html=True)
    
    machines_df = load_active_machines_df()
    if machines_df.empty:
        st.warning("ยังไม่มีข้อมูลเครื่องจักรในระบบ")
        return
    machines = machines_df.to_dict("records")
    codes = [m["machine_code"] for m in machines]

    # --- ระบบสแกน QR ภายในแอป ---
    with st.expander("📷 กดที่นี่เพื่อสแกน QR Code (ไม่ต้องล็อกอินใหม่)"):
        cam_image = st.camera_input("ถ่ายรูป QR Code ของเครื่องจักร (ให้เห็น QR ชัดๆ ในกรอบ)", label_visibility="collapsed")
        if cam_image:
            try:
                bytes_data = cam_image.getvalue()
                cv2_img = cv2.imdecode(np.frombuffer(bytes_data, np.uint8), cv2.IMREAD_COLOR)
                detector = cv2.QRCodeDetector()
                data, bbox, _ = detector.detectAndDecode(cv2_img)
                
                if data:
                    parsed_url = urlparse(data)
                    query_params = parse_qs(parsed_url.query)
                    if 'machine' in query_params:
                        scanned_code = query_params['machine'][0]
                        if scanned_code in codes:
                            st.session_state.selected_machine_code = scanned_code
                            st.success(f"✅ สแกนสำเร็จ! เปลี่ยนเป็นเครื่อง: {scanned_code} ให้แล้วครับ เลื่อนลงไปกรอกข้อมูลได้เลย")
                        else:
                            st.warning(f"❌ สแกนได้รหัส {scanned_code} แต่ไม่มีเครื่องนี้ในระบบ")
                    else:
                        st.warning("❌ QR Code นี้ไม่ใช่ของระบบ MHT PM")
                else:
                    st.error("❌ ระบบอ่าน QR Code ไม่ได้ กรุณาถ่ายให้อยู่ในกรอบชัดเจนและไม่เบลอ")
            except Exception as e:
                st.error("❌ เกิดข้อผิดพลาดในการประมวลผลรูปภาพ")

    st.markdown("---")

    # กำหนดค่า default_idx ให้เลือกเครื่องที่สแกนมา
    if "selected_machine_code" not in st.session_state:
        if st.query_params.get("machine") in codes:
            st.session_state.selected_machine_code = st.query_params.get("machine")
        else:
            st.session_state.selected_machine_code = codes[0]
            
    if st.session_state.selected_machine_code not in codes:
        st.session_state.selected_machine_code = codes[0]

    default_idx = codes.index(st.session_state.selected_machine_code)

    st.markdown("### 📌 ข้อมูลการตรวจเช็ก (Inspection Info)")
    col_m1, col_m2 = st.columns(2)
    with col_m1:
        machine_code = st.selectbox("⚙️ เลือกรหัสเครื่องจักร (Machine Code):", codes, index=default_idx)
        st.session_state.selected_machine_code = machine_code
    with col_m2:
        # เปลี่ยนเป็นติ๊ก (Radio) เพื่อเลือกรอบการทำ PM 
        pm_type = st.radio("📅 เลือกรอบการตรวจเช็ก (Type of Preventive):", ["Monthly", "Half-Year", "Yearly"], horizontal=True)
        
    selected_machine = next(m for m in machines if m["machine_code"] == machine_code)
    items_df = load_checklist_items_df(selected_machine["id"])

    # แสดง Checklist ทั้งหมดเสมอ (ไม่ซ่อนตามที่เลือก)
    periods_to_show = ["Monthly", "Half-Year", "Yearly"]

    with st.form("pm_submission_form", clear_on_submit=False):
        col1, col2 = st.columns(2)
        with col1:
            # ดึง Model ตามเครื่องที่เลือกมาโชว์ให้อัตโนมัติ
            default_model = MACHINE_MODELS.get(machine_code, "")
            model_input = st.text_input("Model:", value=default_model)
            pm_date = st.date_input("Preventive date:")
            time_start = st.time_input("Time start:", value=datetime.strptime("08:00", "%H:%M").time())
        with col2:
            location_input = st.text_input("Machine Location:", value="Motor Core")
            comp_date = st.date_input("Completed date:")
            time_finish = st.time_input("Time Finish:", value=datetime.strptime("17:00", "%H:%M").time())
        work_week = st.text_input("Work week:")

        st.markdown("---")

        th_col1, th_col2, th_col3, th_col4 = st.columns([1, 4, 3, 3])
        th_col1.markdown("**Item**")
        th_col2.markdown("**Description**")
        th_col3.markdown("**Condition**")
        th_col4.markdown("**Remark**")
        st.markdown("<hr style='margin: 0; padding: 0;'>", unsafe_allow_html=True)

        results = []
        has_abnormal = False

        for period in periods_to_show:
            period_items = items_df[items_df['period'] == period]
            if period_items.empty: continue

            st.markdown(f"<div style='background-color: #f0f2f6; padding: 5px; margin-top: 10px;'><b>📍 {period}</b></div>", unsafe_allow_html=True)

            for idx, row in period_items.reset_index().iterrows():
                c1, c2, c3, c4 = st.columns([1, 4, 3, 3])
                c1.write(f"{idx + 1}")
                c2.write(row['item_name'])

                # --- SMART DEFAULT LOGIC ---
                # ระบบช่วยติ๊กให้อัตโนมัติ: เลือกรอบไหน ก็จะขึ้น 'OK' นอกนั้นขึ้น 'No have PM item'
                if pm_type == "Monthly" and period in ["Half-Year", "Yearly"]:
                    default_idx = 2  # ชี้ไปที่ "No have PM item"
                elif pm_type == "Half-Year" and period == "Yearly":
                    default_idx = 2  # ชี้ไปที่ "No have PM item"
                else:
                    default_idx = 0  # ชี้ไปที่ "OK"

                condition = c3.radio("Condition", options=["OK", "Need correction", "No have PM item"],
                                     index=default_idx,
                                     key=f"cond_{row['id']}", horizontal=True, label_visibility="collapsed")
                remark = c4.text_input("Remark", key=f"rem_{row['id']}", label_visibility="collapsed", placeholder="...")

                is_abn = (condition == "Need correction")
                if is_abn: has_abnormal = True

                results.append({
                    "period": period, "item_name": row['item_name'], "item_type": "radio",
                    "value": condition, "remark": remark, "is_abnormal": is_abn
                })
                st.markdown("<hr style='margin: 0; border-top: 1px dashed #ddd;'>", unsafe_allow_html=True)

        st.markdown("---")
        part_changing = st.text_input("Part changing :")
        note = st.text_area("Note :")

        st.markdown("📸 **หลักฐานความผิดปกติ (รูปภาพ)**")
        uploaded_file = st.file_uploader("ถ่ายรูปหรือเลือกไฟล์เพื่อแนบเข้าระบบ", type=["png", "jpg", "jpeg"])

        submitted = st.form_submit_button("💾 บันทึกข้อมูล (Submit PM)")

    if submitted:
        image_data = base64.b64encode(uploaded_file.getvalue()).decode("utf-8") if uploaded_file else None
        if uploaded_file:
            st.image(uploaded_file, width=250)

        if has_abnormal:
            st.warning("⚠️ พบรายการ Need correction ระบบจะบันทึกเป็น 'พบข้อบกพร่อง' และรอให้หัวหน้างาน ตรวจสอบ (Review)")

        conn = get_conn()
        c = conn.cursor()
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status = "abnormal" if has_abnormal else "normal"

        c.execute("""INSERT INTO pm_logs 
                     (timestamp, machine_id, machine_code, machine_name, inspector_username, inspector_name, 
                      preventive_type, model, location, time_start, time_finish, work_week, part_changing, note, image_data, status)
                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (now, selected_machine["id"], selected_machine["machine_code"], selected_machine["machine_name"],
                   user["username"], user["full_name"], pm_type, model_input, location_input, 
                   time_start.strftime("%H:%M"), time_finish.strftime("%H:%M"), work_week, part_changing, note, image_data, status))

        log_id = c.lastrowid
        for res in results:
            c.execute("""INSERT INTO pm_log_details (log_id, period, item_name, item_type, value, remark, is_abnormal)
                         VALUES (?,?,?,?,?,?,?)""",
                      (log_id, res["period"], res["item_name"], res["item_type"], res["value"], res["remark"], res["is_abnormal"]))
        conn.commit()
        conn.close()
        st.success("บันทึกข้อมูลเรียบร้อยแล้ว!")
        st.balloons()


# ==========================================
# PAGE: DASHBOARD & APPROVAL
# ==========================================
def dashboard_page(user):
    st.header("📊 แดชบอร์ดตรวจสอบเอกสาร (Review)")
    conn = get_conn()
    
    c1, c2, c3 = st.columns(3)
    with c1:
        f_machine = st.selectbox("กรองเครื่องจักร", ["ทั้งหมด"] + [m["machine_code"] for m in get_all_machines()])
    with c2:
        f_status = st.selectbox("กรองผลการประเมิน", ["ทั้งหมด", "normal", "abnormal"])
    with c3:
        f_approve = st.selectbox("กรองการตรวจสอบ", ["ทั้งหมด", "รอตรวจสอบ", "ตรวจสอบแล้ว"])

    query = "SELECT * FROM pm_logs WHERE 1=1"
    params = []
    if f_machine != "ทั้งหมด":
        query += " AND machine_code = ?"
        params.append(f_machine)
    if f_status != "ทั้งหมด":
        query += " AND status = ?"
        params.append(f_status)
    if f_approve == "รอตรวจสอบ": query += " AND is_approved = 0"
    elif f_approve == "ตรวจสอบแล้ว": query += " AND is_approved = 1"
    query += " ORDER BY id DESC LIMIT 100"

    df_logs = pd.read_sql_query(query, conn, params=params)
    
    if df_logs.empty:
        st.info("ยังไม่มีประวัติการบันทึก PM ตามเงื่อนไขที่เลือก")
        conn.close()
        return

    st.markdown("---")
    st.subheader("📥 โหลดรายงาน (Excel)")
    
    log_ids = df_logs['id'].tolist()
    placeholders = ",".join(["?"] * len(log_ids))
    df_details = pd.read_sql_query(
        f"SELECT * FROM pm_log_details WHERE log_id IN ({placeholders})", conn, params=log_ids
    )
    
    excel_data = generate_paper_like_excel(df_logs, df_details)
    st.download_button(
        label="🖨️ ดาวน์โหลดไฟล์ Excel รวมทุกเอกสาร",
        data=excel_data,
        file_name=f"PM_Report_MHT_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.markdown("---")

    st.subheader("📝 ประวัติการตรวจเช็ก (รูปแบบตารางทางการ)")
    
    display_df = df_logs.copy()
    display_df = display_df.rename(columns={
        "timestamp": "วัน-เวลาที่บันทึกข้อมูล",
        "machine_code": "รหัสเครื่องจักร",
        "preventive_type": "รอบการบำรุงรักษา",
        "inspector_name": "ผู้รับผิดชอบ (ช่าง)",
        "status": "ผลการประเมิน",
        "is_approved": "สถานะการตรวจสอบ",
        "approved_by": "ผู้ตรวจสอบ (Review By)"
    })
    
    display_df["ผลการประเมิน"] = display_df["ผลการประเมิน"].map({"normal": "🟢 ปกติ", "abnormal": "🔴 พบข้อบกพร่อง"})
    display_df["สถานะการตรวจสอบ"] = display_df["สถานะการตรวจสอบ"].apply(lambda x: "✅ ตรวจสอบแล้ว" if x == 1 else "⏳ รอการตรวจสอบ")
    
    cols_to_show = ["วัน-เวลาที่บันทึกข้อมูล", "รหัสเครื่องจักร", "รอบการบำรุงรักษา", "ผลการประเมิน", "สถานะการตรวจสอบ", "ผู้รับผิดชอบ (ช่าง)", "ผู้ตรวจสอบ (Review By)"]
    st.dataframe(display_df[cols_to_show], use_container_width=True, hide_index=True)

    st.markdown("---")
    st.subheader("🔍 รายละเอียดเชิงลึก (คลิกเพื่อดูรายละเอียด / แนบรูป / ยืนยันตรวจสอบ / ส่งเมลให้นายใหญ่)")

    for _, log in df_logs.iterrows():
        status_icon = "🔴 พบข้อบกพร่อง" if log["status"] == "abnormal" else "🟢 ปกติ"
        approve_icon = "✅ ตรวจสอบแล้ว" if log["is_approved"] else "⏳ รอการตรวจสอบ"
        
        with st.expander(f"ข้อมูลวันที่ {log['timestamp'][:16]} | รหัสเครื่อง: {log['machine_code']} | รอบ: {log['preventive_type']} | {status_icon} | {approve_icon}"):
            col1, col2 = st.columns(2)
            col1.write(f"**ผู้รับผิดชอบการตรวจสอบ:** {log['inspector_name']} | **ช่วงเวลาดำเนินการ:** {log['time_start']} - {log['time_finish']}")
            work_week_display = log['work_week'] if 'work_week' in log and log['work_week'] else '-'
            col1.write(f"**Work week:** {work_week_display}")
            col1.write(f"**รายการเปลี่ยนอะไหล่ (Part changing):** {log['part_changing'] or '-'}")
            col1.write(f"**ข้อเสนอแนะเพิ่มเติม (Note):** {log['note'] or '-'}")
            
            if pd.notna(log['image_data']):
                img_val = str(log['image_data']).strip()
                if img_val and img_val.lower() not in ('nan', 'none'):
                    try:
                        img_bytes = base64.b64decode(img_val)
                        col2.image(img_bytes, caption="เอกสาร/ภาพถ่ายแนบประกอบ", width=250)
                    except Exception:
                        col2.warning("ไม่สามารถแสดงผลภาพถ่ายได้")

            df_det = df_details[df_details['log_id'] == log['id']][['item_name', 'value', 'remark']]
            df_det.columns = ['รายการที่ทำการตรวจสอบ (Inspection Item)', 'ผลการประเมิน (Condition)', 'หมายเหตุ (Remark)']
            st.dataframe(df_det, use_container_width=True, hide_index=True)
            
            btn_col1, btn_col2 = st.columns([1, 1])
            
            if not log["is_approved"] and user["role"] in ["supervisor", "admin"]:
                if btn_col1.button(f"✅ ยืนยันตรวจสอบ & ส่ง Excel หาผู้จัดการ", key=f"app_{log['id']}", type="primary"):
                    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    c = conn.cursor()
                    c.execute("UPDATE pm_logs SET is_approved=1, approved_by=?, approved_at=? WHERE id=?", 
                              (user['full_name'], now_str, log['id']))
                    conn.commit()
                    
                    df_log_single = pd.read_sql_query("SELECT * FROM pm_logs WHERE id=?", conn, params=(log['id'],))
                    df_details_single = pd.read_sql_query("SELECT * FROM pm_log_details WHERE log_id=?", conn, params=(log['id'],))
                    excel_bytes = generate_paper_like_excel(df_log_single, df_details_single)
                    
                    success, email_msg = send_email_with_excel(excel_bytes, log['machine_code'], log['preventive_type'], log['timestamp'][:10])
                    
                    if success:
                        st.success("บันทึกการตรวจสอบและส่งเอกสารแนบให้ผู้จัดการทางอีเมลสำเร็จ!")
                    else:
                        st.warning(f"บันทึกการตรวจสอบเรียบร้อยแล้ว แต่ไม่สามารถส่งอีเมลได้: {email_msg}")
                        
                    st.rerun()
            elif log["is_approved"]:
                btn_col1.success(f"ผ่านการตรวจสอบ (Review) โดย: {log['approved_by']} (เมื่อเวลา {log['approved_at']})")
                
            if user["role"] == "admin":
                if btn_col2.button(f"🗑️ ลบเอกสารชุดนี้ (Admin Only)", key=f"del_log_{log['id']}"):
                    delete_pm_log(log['id'])
                    st.success("ลบเอกสารออกจากระบบเรียบร้อยแล้ว!")
                    st.rerun()

    conn.close()

# ==========================================
# PAGE: ADMIN - MACHINES
# ==========================================
def admin_machines_page():
    st.header("🏭 จัดการเครื่องจักรและรายการตรวจเช็ก")
    tab1, tab2 = st.tabs(["➕ เพิ่ม/แก้ไขเครื่องจักร", "📋 จัดการรายการตรวจ (Checklist)"])

    with tab1:
        st.subheader("เพิ่มเครื่องจักรใหม่")
        with st.form("add_machine_form", clear_on_submit=True):
            code = st.text_input("รหัสเครื่องจักร (เช่น UN-12)")
            name = st.text_input("ชื่อเครื่องจักร (เช่น Uncoiler UN-12)")
            category = st.text_input("หมวดหมู่ (เช่น Uncoiler)")
            submitted = st.form_submit_button("เพิ่มเครื่องจักร")
            if submitted:
                if not code or not name:
                    st.error("กรุณากรอกรหัสและชื่อเครื่องจักร")
                else:
                    try:
                        conn = get_conn()
                        c = conn.cursor()
                        c.execute("INSERT INTO machines (machine_code, machine_name, category) VALUES (?,?,?)",
                                  (code, name, category))
                        conn.commit()
                        conn.close()
                        st.cache_data.clear()
                        st.success(f"เพิ่มเครื่องจักร {name} เรียบร้อยแล้ว")
                        st.rerun()
                    except sqlite3.IntegrityError:
                        st.error("รหัสเครื่องจักรนี้มีอยู่แล้วในระบบ")

        st.markdown("---")
        st.subheader("รายชื่อเครื่องจักรทั้งหมด")
        for m in get_all_machines():
            col1, col2, col3 = st.columns([3, 1, 1])
            col1.write(f"**{m['machine_name']}** ({m['machine_code']}) — {m['category'] or '-'}")
            col2.write("🟢 ใช้งาน" if m["active"] else "🔴 ปิดใช้งาน")
            if col3.button("สลับสถานะ", key=f"toggle_machine_{m['id']}"):
                toggle_machine_active(m["id"], m["active"])
                st.rerun()

    with tab2:
        st.subheader("เพิ่มรายการตรวจเช็กให้เครื่องจักร")
        all_machines = get_all_machines()
        if not all_machines:
            st.info("กรุณาเพิ่มเครื่องจักรก่อน")
        else:
            machine_names = [f"{m['machine_name']} ({m['machine_code']})" for m in all_machines]
            selected = st.selectbox("เลือกเครื่องจักร", machine_names, key="cl_machine_select")
            machine = all_machines[machine_names.index(selected)]

            with st.form("add_checklist_form", clear_on_submit=True):
                period = st.selectbox("รอบการตรวจ (Period)", ["Monthly", "Half-Year", "Yearly"])
                item_name = st.text_input("ชื่อรายการเช็ก (Description)")
                submitted = st.form_submit_button("เพิ่มรายการเช็ก")
                if submitted:
                    if not item_name:
                        st.error("กรุณากรอกชื่อรายการเช็ก")
                    else:
                        conn = get_conn()
                        c = conn.cursor()
                        c.execute("SELECT COALESCE(MAX(order_no), 0) + 1 AS next_order FROM checklist_items WHERE machine_id=? AND period=?", (machine["id"], period))
                        next_order = c.fetchone()["next_order"]
                        c.execute("""INSERT INTO checklist_items 
                                     (machine_id, period, item_name, item_type, options, abnormal_options, order_no)
                                     VALUES (?,?,?,?,?,?,?)""",
                                  (machine["id"], period, item_name, "radio", RADIO_OPTIONS, ABNORMAL_OPTIONS, next_order))
                        conn.commit()
                        conn.close()
                        st.cache_data.clear()
                        st.success(f"เพิ่มรายการ '{item_name}' (รอบ {period}) เรียบร้อยแล้ว")
                        st.rerun()

            st.markdown("---")
            st.subheader(f"รายการตรวจเช็กปัจจุบันของ {machine['machine_name']}")
            conn = get_conn()
            items_df = pd.read_sql_query("SELECT * FROM checklist_items WHERE machine_id=? ORDER BY period, order_no", conn, params=(machine["id"],))
            conn.close()
            
            if items_df.empty:
                st.info("ยังไม่มีรายการเช็ก")
            else:
                for p in ["Monthly", "Half-Year", "Yearly"]:
                    p_items = items_df[items_df['period'] == p]
                    if not p_items.empty:
                        st.markdown(f"**📍 {p}**")
                        for _, row in p_items.iterrows():
                            c1, c2 = st.columns([4, 1])
                            c1.write(f"- {row['item_name']}")
                            if c2.button("ลบ", key=f"del_item_{row['id']}"):
                                conn = get_conn()
                                c = conn.cursor()
                                c.execute("DELETE FROM checklist_items WHERE id=?", (row["id"],))
                                conn.commit()
                                conn.close()
                                st.cache_data.clear()
                                st.rerun()

# ==========================================
# PAGE: ADMIN - USERS
# ==========================================
def admin_users_page():
    st.header("👥 จัดการผู้ใช้งานระบบ (User Management)")
    st.subheader("➕ เพิ่มผู้ใช้งานใหม่")
    with st.form("add_user_form", clear_on_submit=True):
        username = st.text_input("ชื่อบัญชีผู้ใช้ (Username)")
        full_name = st.text_input("ชื่อ-นามสกุล (Full Name)")
        password = st.text_input("รหัสผ่านเริ่มต้น (Initial Password)", type="password")
        role = st.selectbox(
            "ระดับสิทธิ์การใช้งาน (Role)",
            ["technician", "supervisor", "admin"],
            format_func=lambda r: ROLE_LABELS.get(r, r),
        )
        submitted = st.form_submit_button("เพิ่มบัญชีผู้ใช้")
        if submitted:
            if not username or not password:
                st.error("⚠️ กรุณาระบุชื่อบัญชีผู้ใช้และรหัสผ่านให้ครบถ้วน")
            else:
                try:
                    add_user(username, password, full_name, role)
                    st.success(f"✅ เพิ่มผู้ใช้งาน '{username}' เข้าสู่ระบบเรียบร้อยแล้ว")
                    st.rerun()
                except sqlite3.IntegrityError:
                    st.error("⚠️ ชื่อบัญชีผู้ใช้นี้ถูกใช้งานในระบบแล้ว โปรดใช้ชื่ออื่น")

    st.markdown("---")
    st.subheader("📋 รายนามผู้ใช้งานในระบบ")
    for u in list_users():
        col1, col2, col3, col4, col5 = st.columns([2, 2, 1, 1, 1])
        col1.write(f"**{u['username']}**")
        col2.write(f"{u['full_name']} — {ROLE_LABELS.get(u['role'], u['role'])}")
        col3.write("🟢 เปิดใช้งาน" if u["active"] else "🔴 ระงับการใช้งาน")
        if col4.button("สลับสถานะ", key=f"toggle_user_{u['id']}"):
            toggle_user_active(u["id"], u["active"])
            st.rerun()
        
        if u["username"] != "admin":
            if col5.button("🗑️ ลบ", key=f"del_user_{u['id']}"):
                delete_user(u["id"])
                st.rerun()
        else:
            col5.write("")
            
        with st.expander(f"🔑 จัดการรหัสผ่านสำหรับ '{u['username']}'"):
            new_pw = st.text_input("รหัสผ่านใหม่", type="password", key=f"newpw_{u['id']}")
            if st.button("บันทึกการเปลี่ยนแปลงรหัสผ่าน", key=f"confirmpw_{u['id']}"):
                if new_pw:
                    reset_user_password(u["id"], new_pw)
                    st.success("เปลี่ยนรหัสผ่านเสร็จสมบูรณ์")

# ==========================================
# PAGE: ADMIN - SETTINGS (ตั้งค่าอีเมล)
# ==========================================
def admin_settings_page():
    st.header("⚙️ ตั้งค่าระบบ (อีเมลแจ้งเตือน)")
    st.markdown("""
    ระบบจะทำการส่งไฟล์ Excel แนบไปกับอีเมลโดยอัตโนมัติ เมื่อหัวหน้างานกดปุ่ม **"ยืนยันตรวจสอบ (Review)"**
    > **หมายเหตุ:** หากใช้อีเมลของ Gmail ผู้ส่งจะต้องใช้ **รหัสผ่านสำหรับแอป (App Password)** ไม่ใช่รหัสผ่านอีเมลปกติ เพื่อความปลอดภัย
    """)
    
    current_sender_email = get_setting("smtp_email")
    current_sender_pass = get_setting("smtp_password")
    current_manager_email = get_setting("manager_email")
    
    with st.form("email_settings_form"):
        st.subheader("✉️ ข้อมูลผู้ส่ง (System Email)")
        smtp_email = st.text_input("อีเมลผู้ส่ง (เช่น email@gmail.com)", value=current_sender_email)
        smtp_password = st.text_input("รหัสผ่านสำหรับแอป (App Password 16 หลัก)", value=current_sender_pass, type="password")
        
        st.markdown("---")
        st.subheader("👔 ข้อมูลผู้รับ (ผู้จัดการ/นายใหญ่)")
        manager_email = st.text_input("อีเมลผู้จัดการ", value=current_manager_email)
        
        submitted = st.form_submit_button("💾 บันทึกการตั้งค่าอีเมล")
        
        if submitted:
            set_setting("smtp_email", smtp_email)
            set_setting("smtp_password", smtp_password)
            set_setting("manager_email", manager_email)
            st.success("✅ บันทึกข้อมูลการตั้งค่าอีเมลเรียบร้อยแล้ว!")

# ==========================================
# MAIN ROUTING
# ==========================================
init_db_once()
if "user" not in st.session_state: st.session_state.user = None
if st.session_state.user is None:
    login_page()
    st.stop()

current_user = st.session_state.user

st.sidebar.markdown(f"<h2 style='text-align: center;'>MHT PM System</h2>", unsafe_allow_html=True)

try:
    st.sidebar.image(base64.b64decode(LOGO_BASE64), use_container_width=True)
except:
    pass

st.sidebar.markdown(f"👤 **{current_user['full_name']}** ({ROLE_LABELS.get(current_user['role'], current_user['role'])})")
if st.sidebar.button("🚪 ออกจากระบบ"):
    st.session_state.user = None
    st.rerun()
st.sidebar.markdown("---")

menu_options = ["📝 ฟอร์มตรวจเช็ก PM"]
if current_user["role"] in ("supervisor", "admin"):
    menu_options.append("📊 แดชบอร์ด & ตรวจสอบ")
if current_user["role"] == "admin":
    menu_options.append("🏭 จัดการเครื่องจักร/รายการตรวจ")
    menu_options.append("👥 จัดการผู้ใช้งานระบบ")
    menu_options.append("⚙️ ตั้งค่าระบบอีเมล")

menu = st.sidebar.radio("เมนูระบบ", menu_options)

if menu == "📝 ฟอร์มตรวจเช็ก PM":
    pm_form_page(current_user)
elif menu == "📊 แดชบอร์ด & ตรวจสอบ":
    dashboard_page(current_user)
elif menu == "🏭 จัดการเครื่องจักร/รายการตรวจ":
    admin_machines_page()
elif menu == "👥 จัดการผู้ใช้งานระบบ":
    admin_users_page()
elif menu == "⚙️ ตั้งค่าระบบอีเมล":
    admin_settings_page()