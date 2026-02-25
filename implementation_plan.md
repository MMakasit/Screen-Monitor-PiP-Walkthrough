## Proposed Changes

### [Component] Core Application (Python + PyQt6)

#### [MODIFY] [main.py](file:///c:/Users/KCUF/Desktop/capturemonitorgame/main.py)
อัปเกรดระบบการแคปภาพให้รองรับการแคปหน้าต่างที่ถูกบัง (Covered) หรือพับอยู่ (Minimized):
- **UPDATE**: `CaptureThread`: ปรับปรุงฟังก์ชัน `capture_window_background` ให้จัดการกับ Stride (Scanline padding) และ Pixel Format ได้ถูกต้องเพื่อแก้ปัญหาภาพเบี้ยว (Skew) และสีเพี้ยน (Grayscale)
- **NEW**: เพิ่มการตรวจสอบว่าหน้าต่างเป็นแบบ Hardware Accelerated หรือไม่ และรองรับการดึงภาพผ่าน `BitBlt` ในกรณีที่ `PrintWindow` ให้ผลลัพธ์ไม่สมบูรณ์
- **FEATURE**: ปรับพิกัดการแคปให้แม่นยำขึ้นโดยใช้ `GetClientRect` แทน `GetWindowRect` ในบางกรณี

## Dependencies
- `PyQt6`: สำหรับ UI
- `mss`: สำหรับการแคปหน้าจอ (โหมด Manual Area)
- `Pillow`: สำหรับการจัดการรูปภาพ
- `pywin32`: สำหรับ Win32 API
- `ctypes`: สำหรับการเรียกโหมดขั้นสูงของ `PrintWindow`
- **NEW** `numpy`: สำหรับการประมวลผล Buffer ภาพที่รวดเร็วและแม่นยำ (ถ้าจำเป็นเพื่อแก้ปัญหา Stride)

## Deployment Plan

### การสร้างไฟล์ Executable (.exe)
- ใช้ `PyInstaller` เพื่อแพ็กเกจโปรแกรม:
  - `--onefile`: ให้เหลือไฟล์เดียว
  - `--windowed` / `--noconsole`: เพื่อไม่ให้มีหน้าต่าง Terminal ดำๆ ขึ้นมาระหว่างใช้งาน
  - รวม dependencies (PyQt6, mss, NumPy, pywin32, Pillow) เข้าไปในไฟล์เดียว

### การอัปโหลดขึ้น GitHub
- **Repo**: `https://github.com/MMakasit/Screen-Monitor-PiP-Walkthrough.git`
- **ขั้นตอน**:
  1. `git init`
  2. `git remote add origin ...`
  3. `git add .`
  4. ทำการ Commit และ Push ไปยังสาขา `main`
