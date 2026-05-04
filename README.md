# Assignment 1 – Singlish to Sinhala Transliteration Test Automation

> **Course:** IT Project Management & Quality Assurance
> **Student ID:** IT23685116
> **Target URL:** [pixelssuite.com/chat-translator](https://www.pixelssuite.com/chat-translator)

---

## Overview

This project contains **50 automated negative test cases** for validating the edge-case handling of the PixelsSuite Chat Translator web application — a Singlish → Sinhala transliteration tool.

All 50 test cases are **negative / boundary tests** designed to verify:
- Graceful handling of invalid or empty inputs
- Correct passthrough of non-Singlish content
- Stability under stress and malformed inputs
- Security against injection attacks (XSS, SQL, HTML)
- Proper behaviour with special characters, emojis, and foreign scripts

---

## Project Structure

```
ITPM-assignment-IT23685116-/
├── test_automation.py               # Playwright automation script
├── setup_testcases.py               # Script to populate the Excel test cases
├── Assignment 1 - Test cases.xlsx   # Test cases with results
└── README.md
```

---

## Setup Instructions

### 1. Install Python Dependencies

```powershell
pip install playwright openpyxl
python -m playwright install chromium
```

### 2. Populate the Excel File

```powershell
python setup_testcases.py
```

> ⚠️ Make sure the Excel file is **closed** before running this script.

### 3. Run the Automation

```powershell
python test_automation.py `
  --excel "Assignment 1 - Test cases.xlsx" `
  --url "https://www.pixelssuite.com/chat-translator" `
  --input-col "Input" `
  --expected-col "Expected output" `
  --actual-col "Actual output" `
  --status-col "Status" `
  --wait-ms 6000 `
  --type-delay-ms 60 `
  --slow-mo-ms 150 `
  --save-every 1
```

The script will:
- Open the PixelsSuite Chat Translator in a Chromium browser
- Automatically type each Singlish input from the Excel file
- Capture the Sinhala output
- Write **Actual output** and **Status (PASS/FAIL)** back to the Excel file

---

## Test Cases Summary

| Category | Count |
|---|---|
| Empty / Whitespace | 4 |
| Special Characters | 6 |
| Very Long Input | 6 |
| Script Injection (XSS/SQL/HTML) | 6 |
| Non-Singlish Languages | 6 |
| Gibberish / Keyboard Mash | 5 |
| Mixed Scripts (Sinhala + Singlish) | 5 |
| Numbers Only | 5 |
| Encoding / Unicode Edge Cases | 4 |
| Formatting / Control Characters | 3 |
| **Total** | **50** |

### Input Length Distribution

| Length Type | Count |
|---|---|
| Short | 20 |
| Medium | 19 |
| Long | 11 |

---

## Excel File Columns

| Column | Description | Filled By |
|---|---|---|
| Test Case ID | Neg_0001 … Neg_0050 | Setup script |
| Input length type | Short / Medium / Long | Setup script |
| Input | Singlish / edge-case text | Setup script |
| Expected output | Expected Sinhala / passthrough | Setup script |
| Actual output | Actual output from website | **Automation** |
| Status | PASS / FAIL / UI Error | **Automation** |
| Singlish input types covered | Category of test case | Setup script |
| Evidence or rationale | Why this input was chosen | Setup script |