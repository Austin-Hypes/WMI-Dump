# 🔍 WMI Info Dumper (HTML)

This project queries system-level WMI data using raw **COM** and **WMI** interfaces in **C++**, and formats the results into a clean, styled **HTML file**

---

## ✨ Features

- ✅ Queries key system classes from the `ROOT\\CIMV2` WMI namespace
- ✅ Outputs everything as a clean, formatted HTML file
- ✅ No dependencies — just use native Windows API + COM
- ✅ Useful for diagnostics, inventory tools, or forensic snapshots

---

## 📷 Example Output

The HTML file includes:

- Titled sections for each WMI class
- Tables of key-value property pairs
- Dark theme with monospaced layout for easy reading



---

## 🛠️ Build Instructions

### 🧰 Requirements

- Windows OS (any version with WMI)
- Visual Studio or MSVC (`cl.exe`)
- Admin privileges (recommended for full WMI access)
## Build with MSVC
Also can build the solution to generate executable
```bash
cl /EHsc /W4 /DUNICODE /D_UNICODE main.cpp /link wbemuuid.lib comsupp.lib
