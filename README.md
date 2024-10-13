
# FixFaultyLineBreaks Macro

## Overview
This VBA macro cleans up faulty line breaks in specified columns of your Excel worksheet, which often happens because of automated manipulation of data or insertion from other types of files. It preserves correct line breaks between sentences, logs all changes in a ChangeLog table, and provides processing statistics in a Statistics table.

## Setup Instructions

### Backup Your Workbook
**Important:** Always create a backup of your Excel workbook before running macros to prevent accidental data loss.

### Enable Macros
1. Open your Excel workbook.
2. If prompted with a security warning about macros, choose to enable macros.

### Insert the Macro
1. Press `ALT + F11` to open the VBA Editor.
2. In the VBA Editor, go to **Insert > Module** to add a new module.
3. Copy and paste the VBA macro code into the module.

### Customize Configuration

- **Worksheet Name:**
```vba
Set ws = ThisWorkbook.Sheets("Sheet1") ' <-- Change "Sheet1" to your actual sheet name
```
Replace `"Sheet1"` with the name of your data sheet.

- **Target Columns:**
```vba
targetColumns = Array("F", "H") ' <-- Modify with your target columns
```
Update the array with the column letters you want to process.

- **Sentence Terminators (Optional):**
```vba
sentenceTerminators = ".!?"
```
Add or remove punctuation marks as needed.

## Running the Macro

### Save Your Workbook
Save your workbook as a macro-enabled file (`.xlsm`) to retain the macro.

### Execute the Macro
1. Press `ALT + F8` to open the Macro dialog box.
2. Select **FixFaultyLineBreaksPreservingSentenceBreaks**.
3. Click **Run**.

## Post-Execution

- **ChangeLog Sheet:**
  - **Name:** `ChangeLog`
  - **Content:** Logs each change with details like timestamp, sheet name, column, row, original content, and new content.
  - **Format:** Excel Table named `ChangeLogTable`.

- **Statistics Sheet:**
  - **Name:** `Statistics`
  - **Content:** Summarizes processing metrics such as total cells processed, total modified, and per-column statistics.
  - **Format:** Excel Table named `StatisticsTable`.

## Notes

- **Text Wrapping:** The macro disables text wrapping in all relevant sheets to maintain single-line text entries.
- **Table Styles:** Default styles are applied to tables. You can customize these within the VBA code if desired.
- **Customization:** To process additional columns or add more metrics, modify the `targetColumns` array and the metrics array in the VBA code accordingly.

## Support
For any issues or questions regarding this macro, please refer to VBA resources or contact me directly.
