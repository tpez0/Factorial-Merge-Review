# 🕒 Factorial Merge & Review

**Factorial Merge & Review** is a desktop application that allows HR and payroll teams to **import, merge, compare, and verify** employee timesheets exported from [FactorialHR](https://factorialhr.com/).

Designed to simplify administrative tasks, it automates data processing and provides visual summaries to support payroll validation and attendance tracking.

---

## ✅ Key Features

- 📥 **Import multiple Excel files** from FactorialHR
- 🧮 **Automatically calculate**:
  - Total hours
  - Night shifts
  - Sunday work
- 🔍 **Compare two Excel exports** to detect differences
- 📊 **Generate summary reports** per employee
- 🚨 **Highlight discrepancies** between expected and actual work hours
- 🎛️ User-friendly GUI with tabbed interface (no Excel skills required)

---

## 🖥️ User Interface

The app provides a clean, tabbed interface with three main sections:

1. **Process Timesheets**: Merge and normalize exported files
2. **Totals Counter**: Calculate total, night and Sunday hours
3. **Compare Sheets**: Detect cell-by-cell differences across two timesheet versions

All outputs are saved in Excel format, compatible with further reporting or audits.

---

## 📦 Installation

### 🔧 Requirements
- Python 3.8+
- [pip](https://pip.pypa.io/)
- Windows (recommended) or macOS/Linux with Tkinter support

### 📥 Dependencies
Installed automatically at first launch:
- `openpyxl`
- `pillow`

---

## 🚀 Getting Started

1. Clone the repository:
   ```bash
   git clone https://github.com/your-org/factorial-merge-review.git
   cd factorial-merge-review
