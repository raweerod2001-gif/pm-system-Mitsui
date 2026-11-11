"""
สคริปต์นี้ใช้เพิ่มเครื่องจักรจริงของ Mitsui High-Tec (Thailand) เข้าไปในฐานข้อมูลที่มีอยู่แล้ว
(pm_factory_v3.db) โดยไม่ลบหรือแก้ไขข้อมูลเดิมที่มีอยู่

วิธีใช้:
    1. วางไฟล์นี้ไว้โฟลเดอร์เดียวกับ app.py และ pm_factory_v3.db
    2. รันคำสั่ง: python seed_machines.py
    3. รันซ้ำได้หลายครั้งโดยไม่เป็นอันตราย (เครื่องที่มีรหัสซ้ำจะถูกข้ามไปเอง)

เพิ่มเครื่องจักรทั้งหมด 33 เครื่อง:
    - Uncoiler   UN-01 ถึง UN-11 (10 รายการเช็ก)
    - Leveller   LV-01 ถึง LV-11 (7 รายการเช็ก)
    - Press Machine MS-01 ถึง MS-11 (22 รายการเช็ก)
"""

import sqlite3

DB_PATH = "pm_factory_v3.db"

RADIO_OPTIONS = "ปกติ (OK)|ต้องแก้ไข (Need correction)|ไม่มีรายการนี้ (No have PM item)"
ABNORMAL_OPTIONS = "ต้องแก้ไข (Need correction)"

MACHINE_TYPES = [
    {
        "prefix": "UN",
        "category": "Uncoiler",
        "name_template": "Uncoiler {code}",
        "items": [
            "Check noise of Drum turning condition",
            "Check smoothly of dual drum head rotates",
            "Check structure, locked of main shaft & welding point",
            "Check Hydraulic pump and oil level",
            "Check control operation",
            "Check clutch & brake, fixed point of motor",
            "Check all electrical connection and equipment inside enclosure box (clean & secure)",
            "Grease-up dual drum head rotates 1 shot/point",
            "Check oil leak",
            "Check un-coiler close-open condition",
        ],
    },
    {
        "prefix": "LV",
        "category": "Leveller",
        "name_template": "Leveller {code}",
        "items": [
            "All clean",
            "Check motor current (by INV. Display)",
            "Roller checking by naked eye - must clean and move smoothly",
            "Check control operation (S-shaped looper)",
            "Check lubrication oil and grease 1 shot/point",
            "Check rotation roller",
            "Check emergency s/w condition",
        ],
    },
    {
        "prefix": "MS",
        "category": "Press Machine",
        "name_template": "Press Machine {code}",
        "items": [
            "Clean auto drain filter of pneumatic (water separator)",
            "Check belt main motor - visual check by naked eye",
            "Check coupling encoder - visual check by naked eye",
            "Check condition of magnetic contactor and overload",
            "Visual check wiring terminal, electric equipment inside enclosure panel",
            "Check clutch brake condition by two-hand control pedestal (normal working & secure)",
            "Check function of top stop of main press at low & hi-speed (must stop at accept zone)",
            "Check condition of front and rear guard (baring slide)",
            "Check operation of safety guard sensor",
            "Check cleanliness of motor pump",
            "Check oil level (Press, Feeder & MAC unit)",
            "Check leak point of oil (Feeder, Press & MAC unit)",
            "Check operation of oil cooler and clean-up air filter",
            "Grease-up feeder roller",
            "Check condition and operation function of feeder unit",
            "Check noise of main press",
            "Check condition/noise of enclosure cover of machine",
            "Check and clean inverter of main press",
            "Check condition of level sensor (cutting oil) and pump",
            "Check condition of scrap conveyor",
            "Check machine vibration (M/C bolt and nut stopper for die change)",
            "Check rotation condition of coupling, key, universal joint",
        ],
    },
]


def run():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    c = conn.cursor()

    added = []
    skipped = []

    for mtype in MACHINE_TYPES:
        for i in range(1, 12):
            code = f"{mtype['prefix']}-{i:02d}"
            c.execute("SELECT id FROM machines WHERE machine_code=?", (code,))
            if c.fetchone():
                skipped.append(code)
                continue

            name = mtype["name_template"].format(code=code)
            c.execute(
                "INSERT INTO machines (machine_code, machine_name, category) VALUES (?,?,?)",
                (code, name, mtype["category"]),
            )
            machine_id = c.lastrowid
            for order_no, item_name in enumerate(mtype["items"]):
                c.execute(
                    """INSERT INTO checklist_items
                       (machine_id, item_name, item_type, options, abnormal_options, unit, min_normal, max_normal, order_no)
                       VALUES (?,?,?,?,?,?,?,?,?)""",
                    (machine_id, item_name, "radio", RADIO_OPTIONS, ABNORMAL_OPTIONS, None, None, None, order_no),
                )
            added.append(code)

    conn.commit()
    conn.close()

    print(f"✅ เพิ่มเครื่องจักรใหม่ {len(added)} เครื่อง: {', '.join(added) if added else '-'}")
    if skipped:
        print(f"⏭️  ข้ามเครื่องที่มีอยู่แล้วในระบบ {len(skipped)} เครื่อง: {', '.join(skipped)}")


if __name__ == "__main__":
    run()
