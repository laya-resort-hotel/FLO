# Laya Service Hub v1.14 Firebase Core Setup (TH)

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

ดังนั้นใน Firebase Authentication ต้องมีบัญชี Email/Password ตามรูปแบบนี้

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

## 6) หมายเหตุเรื่องเพิ่ม/ลบทีมงาน
หน้า Team Setup ในแอปนี้จะอัปเดตรายชื่อใน `users` collection ได้
แต่การสร้าง **บัญชีล็อกอินจริง** ใน Firebase Authentication ยังต้องสร้างแยกอีกครั้งใน Firebase Console
หรือทำภายหลังด้วย Admin SDK / Cloud Functions

## 7) Seed demo data
ค่า `seedDemoDataIfEmpty: true` จะช่วยเติมตัวอย่าง Tasks / MOD Reports / roster ลง Firestore อัตโนมัติเมื่อข้อมูลยังว่าง
เหมาะสำหรับเริ่มทดสอบระบบ

## 8) Deployment
อัปโหลดไฟล์ทั้งหมดขึ้น GitHub Pages ได้ตามเดิม
แต่อย่าลืมแก้ `firebase-config.js` ให้เป็นค่าจริงก่อน
