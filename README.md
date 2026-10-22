# 🏫 Government Primary School (GPS) Qala Qalagay — School Management & Examination System

[![Institution](https://img.shields.io/badge/Institution-GPS%20Qala%20Qalagay%20Swat-1a365d.svg)](https://github.com/Rashidzada/GPS-Qala-Qalagay-School-Management-System)
[![EMIS Code](https://img.shields.io/badge/EMIS%20Code-22093-0284c7.svg)](https://github.com/Rashidzada/GPS-Qala-Qalagay-School-Management-System)
[![Status](https://img.shields.io/badge/Status-Production%20Ready-success.svg)](https://github.com/Rashidzada/GPS-Qala-Qalagay-School-Management-System)
[![Coverage](https://img.shields.io/badge/Levels-Primary%20(Class%201st--5th)-emerald.svg)](https://github.com/Rashidzada/GPS-Qala-Qalagay-School-Management-System)
[![Django](https://img.shields.io/badge/Django-5.2+-092e20.svg?logo=django)](https://www.djangoproject.com/)
[![Database](https://img.shields.io/badge/Database-PostgreSQL-336791.svg?logo=postgresql)](https://www.postgresql.org/)
[![License](https://img.shields.io/badge/License-Proprietary-red.svg)](#)

A dedicated, production-grade **Student Information, Examination & Academic Management System** built specifically for **Government Primary School (GPS) Qala Qalagay, Tehsil Kabal, District Swat, Khyber Pakhtunkhwa** (EMIS Code: `22093`).

The system is streamlined specifically for Primary School administration (strictly Class 1st through 5th), focusing on student enrollment, government EMIS cross-referencing, multi-term examinations, gazettes, award sheets, Detailed Marks Certificates (DMCs), roll number slips, and exam fee registers.

---

## 📑 Table of Contents
1. [Core Features](#-core-features)
2. [Classes & Enrollment Structure](#-classes--enrollment-structure)
3. [Technology Stack](#-technology-stack)
4. [Quick Start (Windows 1-Click Launchers)](#-quick-start-windows-1-click-launchers)
5. [Manual Installation & Setup](#-manual-installation--setup)
6. [Initial Data & Fixtures](#-initial-data--fixtures)
7. [Default Credentials](#-default-credentials)
8. [Examination & Result Workflows](#-examination--result-workflows)
9. [Database Backup & Restore](#-database-backup--restore)
10. [Author & Support](#-author--support)

---

## 🏛️ Core Features

### 🎓 1. Student Information System (`students`)
* **Government EMIS Integration (EMIS: `22093`):** Pre-seeded with 209 active GPS primary students enriched with official government admission numbers, verified dates of birth (D.O.B), father CNICs, B-Forms, and parent contact numbers.
* **Primary Enrollment Structure:** Capped strictly to Primary Education (Class 1 to Class 5, Sections A & B).
* **Family / Household Linkage:** Automatically groups siblings under unified parent/guardian households.
* **Institutional Printables:**
  - Official School Leaving Certificate (SLC) with serial numbering (`GPS-SLC-<adm>`)
  - Character Certificate (`GPS-CC-<adm>`)
  - Official Date of Birth (D.O.B) School Certificate
  - Printable Student ID Cards with QR Verification
  - Blank Admission Forms matching KP Government standards

### 📝 2. Examination & Tabulation Management (`exams`)
* **Multi-Term Exam Scheduling:** Create and manage First Term, Mid Term, and Annual Examinations.
* **Dynamic Subject Award Sheets (F-2):** Single-subject award lists with max marks, passing criteria (30% KP Government standard), and check-marking sheets.
* **Spreadsheet-Style Live Marks Entry Grid:** Fast spreadsheet interface with instant auto-save and bulk Excel marks import/export.
* **Official Class Result Gazette (F-1):** Tabulation matrix showing per-subject scores, obtained totals, percentage, grade (A+, A, B, C, D, E, Fail), class positions (1st, 2nd, 3rd with tie resolution), and pass/fail summary.
* **Detailed Marks Certificate (DMC / Report Cards):** Clean, printable A4/half-page report cards for individual students and 1-click bulk DMC printing for entire classes.
* **Candidate Roll Number Slips:** Term examination admit cards with candidate details, exam rules, and candidate photo frames.
* **Exam Fee Collection Register:** Class-wise examination fee tracking, exemptions, and printable collection registers.

### 💾 3. Data Hub & Backup Engine (`school`)
* **Excel Data Hub (`/excel-hub/`):** 1-Click backup and restore of student registries, class structures, and examination marks in Excel format.
* **Automated Daily Backups:** Timestamped PostgreSQL `.sql` backups and multi-sheet Master Excel (`.xlsx`) backups.
* **Desktop Shortcuts & Automation:** Windows `.bat` launchers for automated provisioning, running, backup, and restore.

---

## 🏫 Classes & Enrollment Structure

| Class Level | Sections | Curriculum / Standard |
| :--- | :---: | :--- |
| **Class 1** | Section A, B | Pashto, Urdu, English, Mathematics, General Knowledge, Islamiyat |
| **Class 2** | Section A, B | Pashto, Urdu, English, Mathematics, General Knowledge, Islamiyat |
| **Class 3** | Section A, B | Pashto, Urdu, English, Mathematics, General Science, Islamiyat |
| **Class 4** | Section A, B | Pashto, Urdu, English, Mathematics, General Science, Islamiyat, Social Studies |
| **Class 5** | Section A, B | Pashto, Urdu, English, Mathematics, General Science, Islamiyat, Social Studies |

---

## 🛠️ Technology Stack
* **Backend:** Python 3.10+ / Django 5.2+
* **Database:** PostgreSQL 15+ (Production)
* **Frontend:** HTML5, Modern Vanilla JavaScript, Bootstrap 5, FontAwesome 6
* **Data Processing & Reporting:** `openpyxl`, `psycopg[binary]`, `reportlab`
* **Styling & Forms:** `django-crispy-forms`, `crispy-bootstrap5`

---

## 🚀 Quick Start (Windows 1-Click Launchers)

For non-technical deployment on Windows:

1. **Setup & Initialization:**
   Double-click `gps_setup.bat`. This script will:
   - Create a clean local Python virtual environment (`venv`).
   - Install all required Python packages.
   - Initialize the PostgreSQL database (`kms_db` or `gps_qala_db`).
   - Run database migrations.
   - Automatically load the 209 active GPS student records and exam data from `gps_initial_data.json`.
   - Create a desktop launcher shortcut.

2. **Launch Daily Server:**
   Double-click `gps_run.bat`. The server will boot at `http://127.0.0.1:8000/` and launch in your default web browser automatically.

3. **Automated Backup:**
   Double-click `gps_backup.bat` at any time to take an immediate full database and Excel backup.

4. **Restore Database:**
   Double-click `gps_restore.bat` to restore your data from any snapshot.

---

## 💻 Manual Installation & Setup

If setting up in a standard developer environment or on Linux / macOS:

```bash
# 1. Clone the repository
git clone https://github.com/Rashidzada/GPS-Qala-Qalagay-School-Management-System.git
cd GPS-Qala-Qalagay-School-Management-System

# 2. Create and activate a virtual environment
python -m venv venv

# Windows:
venv\Scripts\activate
# Linux/macOS:
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Configure environment / Database settings
# Ensure PostgreSQL is running on localhost:5432 with database 'kms_db' or configure .env

# 5. Run migrations
python manage.py migrate

# 6. Load official GPS Qala student data fixture
python manage.py loaddata gps_initial_data.json

# 7. Start the local development server
python manage.py runserver 127.0.0.1:8000
```

---

## 🔑 Default Credentials

| Role | Username | Password | Access Level |
| :--- | :--- | :--- | :--- |
| **System Administrator** | `admin` | `admin123` | Full Superuser Access |

*(After initial login, change password via User Settings).*

---

## 📄 Initial Data & Fixtures

The repository includes `gps_initial_data.json`:
- **Active Students:** 209 real enrolled students of GPS Qala Qalagay.
- **Classes:** 5 primary classes (Class 1st to 5th) with active sections.
- **Subjects:** Standard Khyber Pakhtunkhwa Primary Curriculum.
- **Active Exam:** Mid Term Examination with pre-computed subject marks for instant DMC and gazette preview.

---

## 👨‍💻 Author & Support

* **Developer:** **Rashid Zada**
* **GitHub:** [@Rashidzada](https://github.com/Rashidzada)
* **Email:** rashidzada6@gmail.com / developer@gpsqala.edu.pk
* **Target Institution:** Government Primary School Qala Qalagay, Swat, KP
* **EMIS Code:** 22093
