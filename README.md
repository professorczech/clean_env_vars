# 🧹 Windows Environment Variable Cleaner (PowerShell)

This PowerShell script audits, cleans, and optimizes both **User** and **System** environment variables on Windows systems. It removes duplicate entries, validates paths, trims excess separators, and helps prevent errors caused by overly long `PATH` variables.

---

## ✅ Features

- 🔐 **Safe & Backed Up**: Automatically creates timestamped backups before any changes.
- 🧪 **Dry-Run Mode**: Preview what will change using the `-WhatIfChanges` switch.
- 🧼 **Cleans Entries**:
  - Removes duplicate or redundant values.
  - Trims trailing and leading semicolons.
  - Validates each path in variables like `PATH` with `Test-Path`.
  - Skips invalid or empty entries.
- 🔍 **Scope-Aware**: Processes both `User` and `System` registry paths separately.
- 🛑 **Length Guard**: Skips update if a cleaned `PATH` exceeds the 32,767 character Windows limit.
- 🔒 **Permission Check**: Skips System scope unless run with administrative privileges.
- 🔁 **Live Refresh**: Broadcasts changes to update the environment across the system.

---

## 📦 Download

1. Save the script as `env_cleanup.ps1`.
2. Save this `README.md` for documentation.
3. Run the script in PowerShell:

```powershell
# Preview changes without applying
.\env_cleanup.ps1 -WhatIfChanges

# Apply changes (requires admin for System env vars)
Start-Process powershell -Verb runAs -ArgumentList ".\env_cleanup.ps1"
```

---

## 📝 Example Output

```text
Processing User environment variables...
Updated User variable 'Path':
   Old: C:\Python\Scripts;C:\Python;C:\Python\Scripts;
   New: C:\Python\Scripts;C:\Python

Removed User variable 'OLD_SDK_PATH' (empty or invalid after cleaning).

Processing System environment variables...
[WARNING] Not running as administrator. Skipping System processing.
```

---

## 📁 Backups

Each time the script runs, it creates two files in your user profile directory:

```
C:\Users\<You>\EnvBackup_User_YYYYMMDD_HHMMSS.txt
C:\Users\<You>\EnvBackup_System_YYYYMMDD_HHMMSS.txt
```

---

## ⚠️ Requirements

- Windows 10 / 11
- PowerShell 5.1 or newer
- Admin privileges to modify System environment variables

---

## 🙋 FAQ

**Q: Can this script break things?**  
A: It’s very careful — it only removes paths that are:
- Invalid or non-existent
- Duplicated
- Empty or whitespace

**Q: What if I want to restore previous values?**  
A: Just open the `.txt` backups and manually paste the values into your system settings or registry.

**Q: Does it clean other environment variables beyond PATH?**  
A: Yes — it checks and cleans all variables in both scopes. PATH-specific logic applies only to list-style variables.

---

## 📖 License

MIT License

---

> Created with 🛠️ by professorczech
