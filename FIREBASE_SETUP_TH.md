# Laya Service Hub v1.14.1 Firebase Register Setup (TH)

โปรเจกต์นี้ถูกปรับให้พร้อมเชื่อม Firebase สำหรับ 4 ส่วนหลักแล้ว:
- Login ด้วย Firebase Authentication
- Tasks ด้วย Cloud Firestore
- MOD Reports ด้วย Cloud Firestore
- รูป/วิดีโอด้วย Cloud Storage

## 1) ค่าที่ต้องเอาจาก Firebase Console
ไปที่ **Project settings > Your apps > Web app** แล้วนำค่าเหล่านี้มาใส่ใน `firebase-config.js`
- apiKey
- authDomain
- projectId
- storageBucket
- messagingSenderId
- appId

## 2) เปิดบริการที่ต้องใช้
- Authentication > Sign-in method > เปิด **Email/Password**
- Firestore Database > Create database
- Storage > Get started

## 3) วิธีล็อกอินของระบบนี้
หน้า Login ยังให้พนักงานกรอก **Employee ID + Password** เหมือนเดิม
แต่ภายในระบบจะ map เป็นอีเมลแบบนี้:
- `11025` -> `11025@laya.local`
- `22018` -> `22018@laya.local`

เมื่อผู้ใช้กดสมัครสมาชิก ระบบจะสร้างบัญชี Email/Password ตามรูปแบบนี้ให้อัตโนมัติ
ถ้าเป็นบัญชีที่แอดมินสร้างไว้เองใน Firebase Authentication ก็ยังใช้งานได้เหมือนเดิม

## 4) Firestore Collections ที่ระบบใช้
- `users/{employeeId}`
- `tasks/{taskId}`
- `modReports/{reportId}`

## 5) โครงสร้าง users document ขั้นต่ำ
ตัวอย่าง `users/11025`
```json
{
  "employeeId": "11025",
  "name": "Noi",
  "role": "CGM",
  "department": "FO",
  "active": true
}
```

## 6) การสมัครสมาชิกพนักงาน
เวอร์ชันนี้เพิ่มปุ่ม **สมัครสมาชิกพนักงาน** ที่หน้า Login แล้ว

สิ่งที่สมัครเองได้:
- FO Staff
- HK Staff
- ENG Staff
- FB Staff

วิธีทำงาน:
- พนักงานกรอก **ชื่อ + Employee ID + แผนก + Password**
- ระบบจะสร้าง Firebase Authentication แบบ Email/Password ให้อัตโนมัติ โดยแปลงเป็น `employeeId@laya.local`
- จากนั้นระบบจะสร้างหรืออัปเดต `users/{employeeId}` ใน Firestore ให้อัตโนมัติ

ข้อจำกัด:
- Role ระดับหัวหน้างาน / ผู้จัดการ / MOD / CGM ควรสร้างโดยแอดมินหรือผ่าน Team Setup
- ถ้าจะให้พนักงานสมัครเองแบบไม่ต้องถูกเพิ่มในทีมก่อนก็ทำได้ เพราะระบบจะสร้าง profile staff ให้ตามแผนกที่เลือก

## 7) หมายเหตุเรื่องเพิ่ม/ลบทีมงาน
หน้า Team Setup ในแอปนี้ยังอัปเดตรายชื่อใน `users` collection ได้เหมือนเดิม
ผู้จัดการสามารถเพิ่มรายชื่อทีมไว้ก่อนได้ และให้พนักงานคนนั้นมากดสมัครสมาชิกเองภายหลัง

## 8) Seed demo data
ค่า `seedDemoDataIfEmpty: true` จะช่วยเติมตัวอย่าง Tasks / MOD Reports / roster ลง Firestore อัตโนมัติเมื่อข้อมูลยังว่าง
เหมาะสำหรับเริ่มทดสอบระบบ

## 9) Deployment
อัปโหลดไฟล์ทั้งหมดขึ้น GitHub Pages ได้ตามเดิม
แต่อย่าลืมแก้ `firebase-config.js` ให้เป็นค่าจริงก่อน
