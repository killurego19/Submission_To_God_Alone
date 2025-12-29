# Tamil Quran Text Corrector

### بِسْمِ اللَّهِ الرَّحْمَٰنِ الرَّحِيمِ
**In the Name of Allah, the Most Gracious, the Most Merciful**

---

## Overview

**Tamil Quran Text Corrector** is a Python desktop application designed to fix common spacing issues in Tamil text, specifically created for correcting Tamil translations of the Holy Quran stored in Excel files.

When Tamil text is copied, converted, or processed through various systems, it often develops spacing problems that break the character combinations, making the text difficult to read or completely meaningless. This tool automatically detects and fixes these issues.

**Developed by:** Social Media Team, Masjid India, Chennai

---

## Table of Contents

1. [Features](#features)
2. [System Requirements](#system-requirements)
3. [Installation](#installation)
4. [Quick Start Guide](#quick-start-guide)
5. [Detailed Usage Guide](#detailed-usage-guide)
6. [What Issues Does It Fix?](#what-issues-does-it-fix)
7. [Understanding Tamil Text Issues](#understanding-tamil-text-issues)
8. [Troubleshooting](#troubleshooting)
9. [FAQ](#faq)
10. [Technical Details](#technical-details)

---

## Features

✅ **Surah & Verse Location Display** - Shows exact Quran location (Surah X, Verse Y) for each error  
✅ **Automatic Column Detection** - Auto-detects Surah, Verse, and Tamil columns  
✅ **Preview Before Fix** - Scan and preview all issues before making changes  
✅ **Before/After Comparison** - See exactly what will be changed for each error  
✅ **Export Error Report** - Save all errors to a separate Excel file for review  
✅ **Batch Processing** - Process thousands of rows in seconds  
✅ **Non-Destructive** - Creates a new corrected file, original remains unchanged  
✅ **Color-Coded Output** - Easy to read error report with syntax highlighting  
✅ **User-Friendly GUI** - No command line knowledge required  
✅ **Progress Tracking** - Visual progress bar during processing  

---

## System Requirements

| Requirement | Details |
|-------------|---------|
| **Operating System** | Windows 7/8/10/11, macOS, or Linux |
| **Python Version** | Python 3.7 or higher |
| **RAM** | Minimum 4GB (8GB recommended for large files) |
| **Disk Space** | 100MB free space |
| **Excel File Format** | .xlsx or .xls |

---

## Installation

### Step 1: Install Python (if not already installed)

1. Download Python from: https://www.python.org/downloads/
2. During installation, **CHECK** the box that says **"Add Python to PATH"**
3. Click "Install Now"

To verify installation, open Command Prompt (Windows) or Terminal (Mac/Linux) and type:
```bash
python --version
```
You should see something like: `Python 3.11.4`

### Step 2: Install Required Libraries

Open Command Prompt (Windows) or Terminal (Mac/Linux) and run:

```bash
pip install pandas openpyxl
```

**Alternative commands if the above doesn't work:**

```bash
# For Windows
python -m pip install pandas openpyxl

# For Mac/Linux
pip3 install pandas openpyxl

# If you have Anaconda
conda install pandas openpyxl
```

### Step 3: Download the Application

Save the `tamil_quran_fixer_v2.py` file to a folder on your computer, for example:
- Windows: `C:\Users\YourName\Documents\TamilFixer\`
- Mac: `/Users/YourName/Documents/TamilFixer/`

---

## Quick Start Guide

### 5-Step Quick Process:

1. **Double-click** `tamil_quran_fixer_v3.py` to launch the application
2. Click **"Browse..."** and select your Excel file (columns will be auto-detected)
3. Verify the **Surah**, **Verse**, and **Tamil** columns are correctly selected
4. Click **"🔍 Find Errors"** to see all issues with their Surah:Verse locations
5. Click **"✓ Fix All Errors"** to fix the issues and save the corrected file

---

## Detailed Usage Guide

### Launching the Application

**Method 1: Double-Click**
- Navigate to the folder containing `tamil_quran_fixer_v2.py`
- Double-click the file to open it

**Method 2: Command Line**
```bash
cd path/to/folder
python tamil_quran_fixer_v2.py
```

### Application Interface

When the application opens, you'll see:

```
┌─────────────────────────────────────────────────────────────┐
│           Tamil Quran Text Corrector                        │
│         Bismillah hir Rahman nir Raheem                     │
├─────────────────────────────────────────────────────────────┤
│  This tool fixes common Tamil text spacing issues:          │
│  • Removes unwanted spaces before vowel signs (matras)      │
│  • Fixes broken character combinations                      │
│  • Normalizes multiple spaces                               │
│  • Removes zero-width characters                            │
├─────────────────────────────────────────────────────────────┤
│  Select Excel File                                          │
│  [No file selected                    ] [Browse...]         │
├─────────────────────────────────────────────────────────────┤
│  Options                                                    │
│  Tamil Column (leave empty for auto-detect):                │
│  [                                        ]                 │
├─────────────────────────────────────────────────────────────┤
│  Ready                                                      │
│  [═══════════════════════════════════════]                  │
├─────────────────────────────────────────────────────────────┤
│  [Process & Fix Tamil Text] [Preview Issues]    [Clear Log] │
├─────────────────────────────────────────────────────────────┤
│  Results / Log                                              │
│  ┌────────────────────────────────────────────────────────┐ │
│  │ Application ready. Select an Excel file to begin.     │ │
│  │                                                        │ │
│  └────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
```

### Step-by-Step Instructions

#### Step 1: Select Your Excel File

1. Click the **"Browse..."** button
2. Navigate to your Excel file location
3. Select the file (e.g., `Quran_Tamil_English_Arabic.xlsx`)
4. Click **"Open"**

The file name will appear in the selection box, and the log will show:
```
✓ Selected: Quran_Tamil_English_Arabic.xlsx
  Full path: C:\Users\Sam\Documents\Quran_Tamil_English_Arabic.xlsx
```

#### Step 2: Preview Issues (Recommended)

Before making any changes, it's recommended to preview what issues exist:

1. Click **"Preview Issues"** button
2. Wait for the scan to complete
3. Review the log output

**Sample Preview Output:**
```
==================================================
SCANNING FOR TAMIL SPACING ISSUES...
==================================================

Reading file: Quran_Tamil_English_Arabic.xlsx
✓ Loaded 6236 rows, 5 columns

Columns in file: ['Surah', 'Ayah', 'Arabic', 'English', 'Tamil']

✓ Found Tamil column by name: 'Tamil'

Total Tamil columns: 1
--------------------------------------------------

Scanning column: 'Tamil'...
  Row 15: Space before matra: ...அரு ள்...
    Before: அல்லாஹ்வின் அரு ள் மிக்க...
    After:  அல்லாஹ்வின் அருள் மிக்க...
  Row 23: Space before matra: ...கரு ணை...
    Before: அவனது கரு ணை யாலே...
    After:  அவனது கருணையாலே...
  → Found 847 rows with issues in 'Tamil'

==================================================
SCAN SUMMARY
==================================================
Total rows scanned: 6236
Rows with issues: 847
Tamil columns: Tamil

→ Click 'Process & Fix Tamil Text' to fix these issues.
```

#### Step 3: Specify Column (Optional)

If your file has multiple Tamil columns and you only want to fix a specific one:

1. Type the exact column name in the **"Tamil Column"** field
2. Example: `Tamil` or `Tamil_Translation`

**Leave empty** to auto-detect and process all Tamil columns.

#### Step 4: Process and Fix

1. Click **"Process & Fix Tamil Text"** button
2. Watch the progress bar as it processes
3. Wait for the completion message

**Sample Processing Output:**
```
==================================================
PROCESSING FILE...
==================================================

✓ Loaded 6236 rows
✓ Tamil columns to process: ['Tamil']

Processing column: 'Tamil'...
  → Fixed 847 rows in 'Tamil'

Saving to: Quran_Tamil_English_Arabic_corrected.xlsx...

==================================================
✓ PROCESSING COMPLETE - Alhamdulillah!
==================================================
Total rows processed: 6236
Total fixes applied: 847
Output file: C:\Users\Sam\Documents\Quran_Tamil_English_Arabic_corrected.xlsx
```

#### Step 5: Locate Your Corrected File

The corrected file is saved in the **same folder** as your original file with `_corrected` added to the name:

| Original File | Corrected File |
|--------------|----------------|
| `Quran.xlsx` | `Quran_corrected.xlsx` |
| `Tamil_Translation.xlsx` | `Tamil_Translation_corrected.xlsx` |

---

## What Issues Does It Fix?

### 1. Space Before Vowel Signs (Matras)

Tamil vowel signs must attach to consonants. Spaces break the combination.

| Issue | Before | After |
|-------|--------|-------|
| Space before ா | க ா | கா |
| Space before ி | க ி | கி |
| Space before ீ | க ீ | கீ |
| Space before ு | க ு | கு |
| Space before ூ | க ூ | கூ |
| Space before ெ | க ெ | கெ |
| Space before ே | க ே | கே |
| Space before ை | க ை | கை |
| Space before ொ | க ொ | கொ |
| Space before ோ | க ோ | கோ |
| Space before ௌ | க ௌ | கௌ |

### 2. Space Before Pulli (Virama ்)

The pulli (்) removes the inherent vowel from a consonant.

| Before | After |
|--------|-------|
| க ் | க் |
| ள ் | ள் |

### 3. Multiple Consecutive Spaces

Normalizes multiple spaces to single space.

| Before | After |
|--------|-------|
| அல்லாஹ்   அருள் | அல்லாஹ் அருள் |

### 4. Zero-Width Characters

Removes invisible characters that can cause display issues:
- Zero-width space (U+200B)
- Zero-width non-joiner (U+200C)
- Zero-width joiner (U+200D)
- Byte order mark (U+FEFF)

### Real-World Example

**Before (Broken):**
```
அல்லா ஹ் வி ன் அரு ள் மி க்க கரு ணை யா ளனா ன
```

**After (Fixed):**
```
அல்லாஹ்வின் அருள் மிக்க கருணையாளனான
```

---

## Understanding Tamil Text Issues

### Why Do These Issues Occur?

1. **Copy-Paste from PDFs** - PDF text extraction often breaks Tamil character combinations
2. **Font Conversion** - Converting between different Tamil fonts can introduce spaces
3. **OCR Errors** - Optical Character Recognition may not properly combine characters
4. **Encoding Issues** - Converting between different text encodings
5. **Manual Typing Errors** - Accidentally hitting space before completing a character

### Tamil Script Basics

Tamil is an abugida script where:
- **Consonants** (க, ங, ச, ஞ, ட, ண, த, ந, ப, ம, etc.) have an inherent 'a' vowel
- **Vowel signs** (matras) modify the consonant's vowel sound
- **Pulli** (்) removes the inherent vowel

A consonant + vowel sign must be written together WITHOUT space:
- க + ா = கா (ka)
- க + ி = கி (ki)
- க + ் = க் (k, with no vowel)

---

## Troubleshooting

### Error: "Missing optional dependency 'openpyxl'"

**Solution:** Install the required library:
```bash
pip install openpyxl
```

### Error: "No module named 'pandas'"

**Solution:** Install pandas:
```bash
pip install pandas
```

### Error: "No Tamil columns detected"

**Possible causes:**
1. The file doesn't contain Tamil text
2. Tamil text uses a non-Unicode encoding

**Solution:**
- Check if your Excel file actually contains Tamil Unicode text
- Manually specify the column name in the "Tamil Column" field

### Application doesn't open when double-clicked

**Solution:** Run from command line to see the error:
```bash
cd path/to/folder
python tamil_quran_fixer_v2.py
```

### Preview or Process button doesn't respond

**Solution:**
1. Check the Results/Log area for error messages
2. Make sure you've selected a valid Excel file
3. Try running from command line to see detailed errors

### Output file is identical to input

**Possible causes:**
1. No Tamil spacing issues exist in the file
2. Wrong column was processed

**Solution:**
- Use "Preview Issues" first to confirm issues exist
- Manually specify the correct Tamil column name

### Large file taking too long

**For files with 10,000+ rows:**
- Processing may take 1-2 minutes
- Watch the progress bar and status updates
- Don't close the application during processing

---

## FAQ

**Q: Will this modify my original file?**  
A: No! The application creates a NEW file with `_corrected` suffix. Your original file remains unchanged.

**Q: Can I process multiple files at once?**  
A: Currently, the application processes one file at a time. Process each file separately.

**Q: Does it work with .xls (older Excel) files?**  
A: Yes, both .xlsx and .xls formats are supported.

**Q: Can I use this for non-Quran Tamil text?**  
A: Absolutely! This tool works with any Tamil text in Excel files.

**Q: What if my Tamil column has a different name?**  
A: The app auto-detects columns containing Tamil text. You can also manually specify the column name.

**Q: Is my data safe/private?**  
A: Yes, all processing happens locally on your computer. No data is sent anywhere.

**Q: Can I undo the changes?**  
A: Since the original file is preserved, you can always go back to it.

---

## Technical Details

### Tamil Unicode Range

The application processes characters in the Tamil Unicode block (U+0B80 to U+0BFF):

| Character Type | Unicode Range | Examples |
|---------------|---------------|----------|
| Vowels | U+0B85 - U+0B94 | அ, ஆ, இ, ஈ |
| Consonants | U+0B95 - U+0BB9 | க, ங, ச, ஞ |
| Vowel Signs | U+0BBE - U+0BCC | ா, ி, ீ, ு |
| Pulli | U+0BCD | ் |
| Anusvara | U+0B82 | ஂ |

### Files Included

| File | Description |
|------|-------------|
| `tamil_quran_fixer_v3.py` | Main GUI application (Latest - with Surah/Verse location) |
| `tamil_quran_fixer_v2.py` | Previous GUI version |
| `tamil_fixer_cli.py` | Command-line version |
| `requirements.txt` | Python dependencies |
| `README.md` | This documentation |

### Command-Line Version

For advanced users or batch processing:

```bash
# Basic usage
python tamil_fixer_cli.py input.xlsx

# Specify output file
python tamil_fixer_cli.py input.xlsx output.xlsx
```

---

## Support

For issues or feature requests, please contact:

**Spiral Nineteen IT Technologies LLC**

---

## Version History

| Version | Date | Changes |
|---------|------|---------|
| 3.0 | 2024 | Added Surah/Verse location display, column auto-detection dropdowns, error export feature, color-coded output |
| 2.0 | 2024 | Improved error handling, detailed logging, before/after preview |
| 1.0 | 2024 | Initial release |

---

## License

This software is provided free for personal and educational use, especially for Islamic educational purposes.

---

**JazakAllah Khair for using Tamil Quran Text Corrector!**

May Allah accept this effort and make it beneficial for the Ummah.

