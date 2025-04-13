# 📊 Excel Conditional Formatting Comparison Script

This Node.js script compares Conditional Formatting Rules between two Excel worksheets or files. It’s especially useful for validating formatting consistency in automated reporting, templates, or generated Excel sheets.

## ✨ Features

- Comparing the differences texts
- Comparing the differences in conditional formatting rules
- Comparing the differences in cell styles
- Comparing the differences in multiple sheets

## Prerequisites

- Make sure that the Sheet of 2 files are in the same order (can enhance later to compare by sheet name)
- Make sure that the conditional formatting rules are in the same order

## 📦 Installation

```bash
npm install
```

Copy the 2 excels files and paste them to ./excel folder. If you want to compare, for example, file A to file B, please type

```bash
node excel.js A.xlsx B.xlsx
```

The result will be in the `differences.xlsx`
