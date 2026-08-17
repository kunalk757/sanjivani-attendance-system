# 🎓 Sanjivani Attendance System
### *AI-Powered Face Recognition & Smart Attendance Management System*

[![Python Version](https://img.shields.io/badge/python-3.10%20%7C%203.11%20%7C%203.12%20%7C%203.13-blue.svg)](https://www.python.org/)
[![OpenCV](https://img.shields.io/badge/OpenCV-Contrib%204.x-brightgreen.svg)](https://opencv.org/)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%20%2F%2011-blue.svg)](https://www.microsoft.com/windows)
[![Release](https://img.shields.io/badge/release-v1.0.0-orange.svg)](https://github.com/kunalk757/sanjivani-attendance-system/releases)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

---

## 📌 Project Overview

**Sanjivani Attendance System** is an advanced, production-ready biometric attendance solution powered by Computer Vision and Artificial Intelligence. Built using **Python**, **OpenCV**, and **Tkinter**, the system utilizes Haar Cascade classifiers for ultra-fast face detection and **LBPH (Local Binary Patterns Histograms)** for real-time facial recognition.

Designed for educational institutions and corporate environments, it eliminates proxy attendance, automates record-keeping into structured **Excel (`.xlsx`)** and **CSV (`.csv`)** spreadsheets, secures student registration with an administrative PIN, and delivers instant automated attendance reports via **SMTP Email**.

---

## ✨ Key Features

- **⚡ Real-Time Face Recognition**: High-speed, accurate face detection and identity verification using LBPH algorithm.
- **🛡️ Strict Unknown Face Rejection**: Strict confidence thresholding (`< 70` match, `>= 70` Access Denied) prevents unauthorized or proxy attendance.
- **🚫 Duplicate Attendance Prevention**: Automatically detects if a student was already marked present today, preventing redundant entries.
- **👤 30-Sample Face Capture & Training**: Captures, crops, and standardizes 30 high-quality facial samples per student for robust recognition under varying lighting conditions.
- **📊 Dual-Format Storage**: Live synchronization between formatted Microsoft Excel workbooks (`attendance.xlsx`) and lightweight flat logs (`attendance.csv`).
- **🔒 PIN-Protected Admin Security**: Restricted access to student enrollment and registration to prevent unauthorized profile creation (Default PIN: `7754`).
- **📧 Automated SMTP Email Dispatch**: Export and dispatch daily attendance sheets with formatted Excel attachments directly to designated administrators.
- **🖥️ Executive Modern UI**: High-DPI Windows-aware GUI featuring an executive navy and royal blue palette, animated indicators, live video feed, and interactive data tables.
- **📦 Zero-Dependency Windows Executable**: Fully bundled, standalone Windows executable (`Sanjivani-Attendance-System.exe`) requiring no Python installation on client machines.

---

## 🏗️ System Architecture & Workflow

```
                        ┌─────────────────────────────────────────┐
                        │       Student Registration Panel        │
                        │    (PIN-Protected Admin Verification)   │
                        └────────────────────┬────────────────────┘
                                             │
                                             ▼
                        ┌─────────────────────────────────────────┐
                        │      Haar Cascade Face Detection        │
                        │     (Captures 30 Face Image Samples)    │
                        └────────────────────┬────────────────────┘
                                             │
                                             ▼
                        ┌─────────────────────────────────────────┐
                        │          Saved to `students/`           │
                        │        (Local Student Database)         │
                        └────────────────────┬────────────────────┘
                                             │
                                             ▼
  ┌─────────────────────────────────────────────────────────────────────────────────────┐
  │                           Live AI Face Attendance Engine                            │
  │  - LBPH Face Recognizer Trains on Student Face Data                                 │
  │  - Live Webcam Stream & Face Alignment                                              │
  │  - Confidence < 70  ──► Recognized Student ──► Check Duplicate ──► Mark Present     │
  │  - Confidence >= 70 ──► UNKNOWN FACE REJECTED (Access Denied)                       │
  └──────────────────────────────────┬──────────────────────────────────────────────────┘
                                     │
                                     ▼
                        ┌─────────────────────────────────────────┐
                        │           Storage & Reporting           │
                        │  - attendance.xlsx (Formatted Excel)    │
                        │  - attendance.csv (Log File)            │
                        │  - Automated SMTP Email Dispatcher      │
                        └─────────────────────────────────────────┘
```

---

## 🛠️ Technologies Used

| Technology / Library | Purpose |
|---|---|
| **Python 3.10+** | Core programming language and runtime environment |
| **OpenCV (`opencv-contrib-python`)** | Haar Cascade frontal face detection & LBPH face recognizer (`cv2.face`) |
| **NumPy** | High-performance multi-dimensional array operations and image matrix processing |
| **openpyxl** | Automated Excel workbook creation, styling, cell formatting, and table generation |
| **Tkinter / ttk** | Native desktop graphical user interface with High-DPI awareness |
| **PyInstaller** | Executable compilation and binary packaging for 64-bit Windows systems |
| **smtplib & email** | Secure TLS/SSL SMTP client for automated attendance report delivery |

---

## 🚀 Getting Started

You can run the Sanjivani Attendance System either using the **Standalone Windows Executable (No Python Required)** or by running **from Source Code**.

---

### Option A: Using the Windows Executable (`.exe`) — Recommended for End-Users

1. Navigate to the [Releases](https://github.com/kunalk757/sanjivani-attendance-system/releases) section.
2. Download **`Sanjivani-Attendance-System.exe`**.
3. Place `Sanjivani-Attendance-System.exe` in a clean dedicated folder (e.g., `C:\SanjivaniAttendance\`).
4. Double-click **`Sanjivani-Attendance-System.exe`** to run.
5. The application will automatically create and manage all required folders (`students/`, `attendance.xlsx`, `attendance.csv`).

> [!NOTE]
> No Python installation, external DLLs, or terminal commands are needed to run the standalone executable.

---

### Option B: Running from Source Code (For Developers)

#### 1. Prerequisites
- **Python 3.10 to 3.13** installed ([python.org](https://www.python.org/downloads/))
- **Git** installed on your system
- A functional USB Webcam or built-in Laptop Camera

#### 2. Clone the Repository
```powershell
git clone https://github.com/kunalk757/sanjivani-attendance-system.git
cd sanjivani-attendance-system
```

#### 3. Create and Activate Virtual Environment
```powershell
# Windows (PowerShell)
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

#### 4. Install Dependencies
```powershell
pip install --upgrade pip
pip install -r requirements.txt
```

#### 5. Launch the Desktop Application
```powershell
python app.py
```

*(Optional CLI alternative modes)*:
```powershell
# Run CLI Face Attendance directly
python main.py

# Run CLI Student Registration directly
python register.py
```

---

## 🔐 Administrative Controls & Configuration

### 1. Admin Security PIN
- The **Student Registration** tab and critical system reset functions are protected by a master administrator passcode.
- **Default PIN:** `7754`

### 2. Automated Email Configuration (SMTP)
To enable automated dispatch of the Excel attendance ledger to your administrative email:

**Set Environment Variables (PowerShell):**
```powershell
$env:SENDER_EMAIL = "your-institution-email@gmail.com"
$env:SENDER_PASSWORD = "your-16-character-app-password"
python app.py
```

*(Alternatively, you can enter credentials directly in the GUI prompt when clicking **"📧 Email Report"**).*

> [!TIP]
> For Gmail, generate a **16-character App Password** via your Google Account:
> `Manage your Google Account` ➔ `Security` ➔ `2-Step Verification` ➔ `App passwords`.

---

## 🔨 Building the Executable & Installer from Source

To compile the standalone Windows `.exe` and setup installer:

```powershell
# 1. Ensure Python dependencies are installed
pip install -r requirements.txt

# 2. Build the executable and Inno Setup installer
python build_exe.py
```

Or using PowerShell directly:
```powershell
.\build.ps1
```

The compiled binaries will be located at:
- Executable: `dist/Sanjivani Attendance.exe`
- Installer: `dist/Sanjivani-Attendance-Setup.exe`

---

## 📂 Project Structure

```
faceattendance/
├── .gitignore                           # Git ignore specifications for Python, build & local data
├── README.md                            # Comprehensive project documentation
├── requirements.txt                     # Production Python dependencies
├── app.spec                             # PyInstaller standalone build configuration
├── installer.iss                        # Inno Setup installer script with unified multi-resolution icon
├── create_icon.py                       # High-resolution multi-size icon generator (16-256px)
├── create_shortcut.py                   # Desktop & Start Menu shortcut creator with icon cache refresh
├── build_exe.py                         # Full build & packaging orchestration pipeline
├── build.ps1                            # PowerShell automated build script
├── build.bat                            # Windows batch build script
├── sanjivani.ico                        # Multi-resolution application icon (16, 32, 48, 64, 128, 256 px)
├── sanjivani.png                        # High-resolution application logo (512x512)
├── app.py                               # Main Desktop GUI application (Dashboard, Face Engine, Registration)
├── main.py                              # Standalone CLI Face Recognition Attendance script
├── register.py                          # Standalone CLI Student Face Registration script
├── send_email.py                        # SMTP email client and attendance report dispatcher
├── haarcascade_frontalface_default.xml  # Pre-trained OpenCV Haar Cascade frontal face model
├── dist/                                # Distribution folder containing final binaries
│   ├── Sanjivani Attendance.exe         # Standalone Windows executable
│   └── Sanjivani-Attendance-Setup.exe   # Professional Windows installer package
└── students/                            # Registered student image database (Generated at runtime)
```

---

## ⚠️ Important Notes & Troubleshooting

1. **Camera Permissions**: Ensure camera access is allowed for desktop applications in **Windows Settings ➔ Privacy & Security ➔ Camera**.
2. **DirectShow Backend**: The system employs OpenCV's DirectShow (`cv2.CAP_DSHOW`) capture backend on Windows to ensure immediate webcam initialization without Windows Media Foundation lag.
3. **Lighting & Face Angle**: For optimal registration, students should face the camera directly in a well-lit environment without heavy shadows or obstructions.
4. **Data Privacy**: The `.gitignore` configuration ensures local attendance records (`attendance.csv`, `attendance.xlsx`) and biometric face images (`students/`) remain private on your machine and are not published to version control.

---

## 📄 License

This project is open-source and available under the [MIT License](LICENSE).

---

## 👨‍💻 Author & Maintainer

- **GitHub**: [@kunalk757](https://github.com/kunalk757)
- **Repository**: [sanjivani-attendance-system](https://github.com/kunalk757/sanjivani-attendance-system)
