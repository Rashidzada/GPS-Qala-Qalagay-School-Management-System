# GPS Qala Qalagay School Management System - Client Deployment Guide
**Target Institution:** Government Primary School Qala Qalagay, Tehsil Kabal, District Swat, KP
**EMIS Code:** 22093
**Classes:** Primary Classes (Class 1st to Class 5th)

---

### 1. Requirements
* Windows 10 or 11 (64-bit)
* Python 3.10+ (recommended: Python 3.12 or 3.14)
* PostgreSQL 15+ (default credentials: user `postgres`, password `postgres`, port `5432`)

---

### 2. Quick Setup & Launch
1. Double-click `gps_setup.bat`. This automatically:
   * Provisions a dedicated virtual environment (`venv`).
   * Installs all required dependencies.
   * Creates the database and applies all migrations.
   * Loads the official GPS Qala student and exam data from `gps_initial_data.json`.
2. Double-click `gps_run.bat` to launch the server and open the application in your browser at `http://127.0.0.1:8000/`.
3. Login using:
   * **Username:** `admin`
   * **Password:** `admin123`

---

### 3. Backups & Restore
* **Create Backup:** Double-click `gps_backup.bat` to generate a full timestamped `.sql` dump and master Excel backup.
* **Restore Backup:** Double-click `gps_restore.bat` to restore your database.
