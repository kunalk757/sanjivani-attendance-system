# SANJIVANI ATTENDANCE SYSTEM
*AI Face Recognition Attendance Management System*

---

## 📌 Project Overview
Sanjivani Attendance System is an advanced, AI-powered desktop attendance platform built with Python, OpenCV, Tkinter, and LBPH (Local Binary Patterns Histograms) face recognition. It provides real-time, proxy-proof face recognition attendance tracking, Excel & CSV report generation, admin security PIN protection, and automated email reporting.

---

## 🏛️ System Architecture & Workflow

```
Student Registration (app.py or register.py)
    │
    ├─► Enter Student Full Name
    ├─► Open Laptop/USB Webcam
    ├─► Detect Face via Haar Cascade
    ├─► Capture & Standardize 30 Face Samples
    └─► Store in students/<StudentName>/
            │
            ▼
Face Recognition Attendance (app.py or main.py)
    │
    ├─► Click "📸 START FACE ATTENDANCE"
    ├─► Train LBPH Recognizer from Registered Faces
    ├─► Live Camera Feed & Face Detection
    ├─► Real-time Face Recognition (< 70 Confidence = Match)
    ├─► Strict Unknown Face Rejection (>= 70 Confidence = Access Denied)
    ├─► Prevent Duplicate Attendance on the Same Day
    └─► Mark Present in Real Time!
            │
            ▼
Storage & Reporting
    ├─► attendance.xlsx (Excel formatted workbook)
    ├─► attendance.csv (Plain-text CSV log)
    └─► Email Report to rohitvinchu7754@gmail.com with Excel attachment
```

---

## 📁 File Structure

| File | Purpose |
|---|---|
| `app.py` | Main Desktop GUI Application (Attendance Dashboard, Direct Face Attendance, Student Registration, Security PIN lock) |
| `main.py` | Standalone CLI for AI Face Recognition Attendance |
| `register.py` | CLI for Student Registration & 30-sample face capture |
| `send_email.py` | SMTP email dispatcher for emailing `attendance.xlsx` to `rohitvinchu7754@gmail.com` |
| `students/` | Directory containing registered student face databases |
| `attendance.xlsx` | Excel attendance ledger |
| `attendance.csv` | CSV attendance ledger |
| `haarcascade_frontalface_default.xml` | OpenCV Haar Cascade for frontal face detection |

---

## 🚀 Quick Start Guide

### 1. Requirements & Package Installation
```powershell
py -m pip install "opencv-contrib-python<5" openpyxl pillow
```

### 2. Start Desktop Attendance Application
```powershell
py app.py
```

### 3. (Optional) Run CLI Face Attendance
```powershell
py main.py
```

### 4. Admin Security Passcode
To access the Student Registration tab in the GUI, enter admin PIN:
`7754`

---

## ⏱️ Attendance & Verification Features

### Feature 1: Direct AI Face Recognition Attendance
1. On the Attendance Dashboard, click **"📸 START FACE ATTENDANCE"**.
2. Look into the camera.
3. Upon matching with confidence < 70, the student is marked **Present** in `attendance.xlsx` and `attendance.csv`.
4. The dashboard statistics update immediately.

### Feature 2: Duplicate Attendance Prevention
1. If an already-marked student faces the camera again on the same day, the system displays **"(Already Marked Today)"** and prevents duplicate entries.

### Feature 3: Unknown Face Rejection
1. If an unregistered person faces the camera, the system detects confidence >= 70 and rejects the attempt: `⛔ UNKNOWN FACE REJECTED — Access Denied`.

### Feature 4: Student Registration
1. Navigate to **"👤 Student Registration"** tab (enter passcode `7754`).
2. Enter the student's full name and click **"📸 Start Face Capture"**.
3. The system captures 30 face samples and updates the directory.

---

## 📧 Email Report Configuration (SMTP)

To email the Excel attendance sheet to `rohitvinchu7754@gmail.com`:
```powershell
$env:SENDER_EMAIL="yourgmail@gmail.com"
$env:SENDER_PASSWORD="your-16-char-app-password"
py app.py
```
*(You can also input your credentials directly when clicking "📧 Email Report" or "🗑 Reset Log" in the Dashboard).*
