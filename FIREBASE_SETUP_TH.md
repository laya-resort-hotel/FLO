# Laya Service Hub v1.15.3 Firebase Connected Setup (TH)

โปรเจกต์นี้ถูกใส่ค่า Firebase ของโปรเจกต์ **laya-flo-c4251** ให้แล้วในไฟล์ `firebase-config.js`
ตอนนี้ตัวแอปจะพยายามเชื่อม Firebase จริงทันทีเมื่อเปิดเว็บ

## Firebase ที่ไฟล์นี้ใช้
- projectId: `laya-flo-c4251`
- authDomain: `laya-flo-c4251.firebaseapp.com`
- storageBucket: `laya-flo-c4251.firebasestorage.app`

## ก่อนเริ่มเทส ต้องเปิดบริการนี้ใน Firebase Console
1. **Authentication > Sign-in method > Email/Password** = เปิด
2. **Firestore Database** = สร้างฐานข้อมูล
3. **Storage** = Get started

## Rules ที่ควรใช้กับ build นี้
Build นี้ใช้ collections แบบ clean start:
- `users/{employeeId}`
- `tasks_clean/{taskId}`
- `modReports_clean/{reportId}`

ให้นำไฟล์ต่อไปนี้ไปวางใน Firebase Console:
- `firestore.rules`
- `storage.rules`

## วิธีล็อกอินของระบบนี้
หน้า Login ให้กรอก **Employee ID + Password**
แต่ภายในระบบ Firebase Auth จะ map เป็นอีเมลแบบนี้:
- `11025` -> `11025@laya.local`
- `22018` -> `22018@laya.local`

## ถ้าจะทดสอบเร็วที่สุด
### ทางเลือก A: สมัครสมาชิกจากหน้าแอป
ใช้ปุ่ม **สมัครสมาชิกพนักงาน** แล้วกรอก:
- ชื่อ
- รหัสพนักงาน
- แผนก
- รหัสผ่าน

สมัครเองได้เฉพาะ Staff:
- FO Staff
- HK Staff
- ENG Staff
- FB Staff

### ทางเลือก B: สร้างบัญชีเองใน Firebase Console
1. ไปที่ **Authentication > Users > Add user**
2. Email ให้ใช้รูปแบบ `employeeId@laya.local`
3. ตั้งรหัสผ่าน
4. จากนั้นไปที่ Firestore เพิ่ม document ใน `users`

ตัวอย่าง `users/11026`
```json
{
  "employeeId": "11026",
  "name": "Pim",
  "role": "DOR",
  "department": "FO",
  "active": true
}
```

## Roles ที่ build นี้รองรับ
- CGM
- DOR
- FO Staff
- Housekeeping Manager
- HK Staff
- ENG Manager
- ENG Staff
- FB Manager
- FB Staff
- MOD

## หมายเหตุสำคัญ
- บัญชี Manager / DOR / MOD / CGM ควรสร้างผ่าน Firebase Console หรือ Firestore โดยผู้ดูแล
- `seedDemoDataIfEmpty: true` จะช่วยเติม roster ผู้ใช้เริ่มต้นลง Firestore ถ้า collection ยังว่าง
- Build นี้เริ่มงานจาก collection แบบสะอาด คือ `tasks_clean` และ `modReports_clean`
- ถ้าเปิดเว็บแล้วไม่เห็นผล ให้ hard refresh 1 ครั้ง
