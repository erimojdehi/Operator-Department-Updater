
# Operator Department Updater

This project is a **Python automation system** developed for managing **operator and department records** in a fleet management environment. It integrates with **ARIS daily exports** and **FA DataLoader** to ensure that inactive or outdated operator records are detected, updated, and uploaded automatically.  

---

##  Features

- **Automated ARIS Parsing**
  - Reads raw `.txt` files exported daily from ARIS.
  - Extracts key operator data such as name, licence number, class, expiry date, status, department, and medical due date.

- **XML Generation**
  - Transforms parsed data into Excel 2003-compatible `.xml` format for seamless use with FA DataLoader.

- **Comparison & Change Detection**
  - Compares current operator data with previous day’s records.
  - Detects and reports all changes in licence class, status, expiry, department, or medical due dates.
  - Identifies and flags inactive operators for removal.

- **Automated Reporting**
  - Generates a **main HTML summary report** showing:
    - Detected changes.
    - Inactive operators flagged.
    - Licence expiries or medical due dates within the warning window.
  - Creates **individual HTML email reports** for each operator with relevant changes, formatted in single-row tables for clarity.

- **FA DataLoader Integration**
  - Automatically prepares and updates the `runfile.bat` script with the correct daily XML filename.
  - Executes FA DataLoader to upload changes directly into AssetWorks.
  - Extracts DataLoader confirmation from its generated `.txt` log file and embeds it into the HTML email report.

- **Retention & Cleanup**
  - Daily deletion of yesterday’s XML file.
  - Maintains only a single appended **program log file** with separators between runs.
  - Logs include total entries parsed, comparison results, operators uploaded, upload failures, and DataLoader messages.

- **Robust Logging & Error Handling**
  - Detailed program execution log saved to a designated `logs` folder.
  - Appends DataLoader’s `.txt` output log into the main log for a complete audit trail.

---

##  Security and Privacy

⚠️ **Important Note:**
This repository contains only the **calculation algorithms, reporting logic, and workflow automation code**.  
All **sensitive and private information** (e.g., internal usernames, passwords, server addresses, licence data, and operator details) has been completely removed for security reasons.  

---

##  Usage

1. Place daily ARIS export `.txt` files in the designated input folder.
2. Run the program manually or via Task Scheduler/automated runner.
3. Program will:
   - Parse and convert data into `.xml`.
   - Run DataLoader with the correct configuration.
   - Generate reports and logs in the configured directories.
4. Review HTML reports and appended log files for confirmation.

---

##  Future Enhancements

- Optional dashboard for reviewing daily runs.
- Configurable expiry/medical due thresholds.
- Multi-department reporting support.
- Email distribution integration directly via SMTP.
