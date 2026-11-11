import os
import qrcode

# ==========================================
# จุดที่ต้องแก้: เอาลิงก์หน้าเว็บจริงของคุณมาใส่ในเครื่องหมายคำพูดด้านล่าง
# ==========================================
BASE_URL = "https://pm-system-mht1-brk4tzlcpptrhsbevuyhmt.streamlit.app"  


# ==========================================
# รายชื่อรหัสเครื่องจักรทั้งหมด (ไม่ต้องแก้แล้ว ผมใส่ให้ครบ 50 เครื่อง)
# ==========================================
MACHINE_CODES = (
    [f"UN-{i:02d}" for i in range(1, 12)]        # Uncoiler 11 เครื่อง
    + [f"LV-{i:02d}" for i in range(1, 12)]      # Leveller 11 เครื่อง
    + [f"MS-{i:02d}" for i in range(1, 12)]      # Press Machine 11 เครื่อง
    + [f"AS-{i:02d}" for i in range(1, 18)]      # Auto Stack 17 เครื่อง
)

def run():
    # สร้างโฟลเดอร์ชื่อ qr_codes เอาไว้เก็บรูป
    os.makedirs("qr_codes", exist_ok=True)
    
    for code in MACHINE_CODES:
        # สร้าง URL สำหรับเครื่องนั้นๆ (เคล็ดลับอยู่ที่ ?machine=รหัสเครื่อง)
        url = f"{BASE_URL}/?machine={code}"
        
        # สร้างรูป QR Code
        img = qrcode.make(url)
        path = os.path.join("qr_codes", f"{code}.png")
        img.save(path)
        print(f"สร้างสำเร็จ: {code}.png -> ลิงก์: {url}")

    print(f"\n✅ เสร็จสิ้น! สร้าง QR Code ทั้งหมด {len(MACHINE_CODES)} ไฟล์ ไปดูในโฟลเดอร์ qr_codes ได้เลย")

if __name__ == "__main__":
    run()