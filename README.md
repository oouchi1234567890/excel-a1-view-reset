# excel-a1-view-reset

An Excel VBA tool that selects cell A1 and resets the view on each visible worksheet in Excel workbooks in a specified folder.

## Features

- Finds `.xls`, `.xlsx`, `.xlsm`, and `.xlsb` files directly in the specified folder.
- Shows the total file count and up to 30 filenames before processing.
- Starts only after confirmation. Choosing **No** or closing the prompt leaves the target workbooks unopened.
- Processes every matching file after confirmation, including files beyond the 30-name preview.
- Visits visible worksheets from right to left, selects A1, and scrolls to the top-left without changing Freeze Panes.
- Leaves the leftmost visible worksheet selected, saves over the original file, and closes it.
- Continues if an individual file cannot be processed, then reports successes and failures.
- Skips hidden worksheets, Excel lock files (`~$...`), and the workbook containing this macro.

## Before you begin

This tool overwrites the original workbooks. Back them up and test with copies first. If you are using a company computer, follow your organization's policies and contact your administrator if macros are restricted.

### Prepare a macro-enabled workbook

This repository contains the VBA source module (`ResetView.bas`), not a prebuilt macro-enabled workbook.

1. Create a new Excel workbook and save it as an **Excel Macro-Enabled Workbook** (`.xlsm`).
2. Create a worksheet named `Control`.
3. Enter the target folder path in cell `B2` on `Control`.
4. Open the Visual Basic Editor (`Alt`+`F11`), choose **File → Import File**, and import `ResetView.bas`.
5. Add a Form Control button to the `Control` sheet and assign it to `RunA1ViewReset`. You can label it **Run** (or **実行** in Japanese).
6. Save the workbook.

### Enable macros safely

1. Open the macro-enabled workbook and inspect its source and contents so you know you trust it.
2. If Excel opens it in Protected View or warns that macros are disabled, review the file and use **Enable Editing** and **Enable Content** only as appropriate.
3. Enabling editing does not necessarily enable macros. Depending on Excel and your organization's settings, files downloaded from the internet may have macros blocked and **Enable Content** may not appear. Do not bypass such a warning; consult Microsoft's guidance or your administrator.

If your organization permits it and normal macro enabling does not work, you may place the checked and trusted tool workbook in a dedicated folder and register only that folder as a Trusted Location through **File → Options → Trust Center → Trust Center Settings → Trusted Locations**. Files in a Trusted Location bypass usual security checks, so do not register broad folders such as Desktop or Downloads and keep only trusted files there. The folder containing the workbooks to process does not need to be a Trusted Location. **Enabling all macros is not recommended.**

## Use

1. Put the Excel workbooks to process in the target folder.
2. Enter that folder's path in `Control!B2`.
3. Click the button assigned to `RunA1ViewReset`.
4. Review the total and filename preview. Choose **Yes** to continue or **No** to cancel.
5. Review the completion summary and check the resulting workbooks.

The search is limited to files immediately inside the selected folder; subfolders are not searched. Password-protected, read-only, damaged, or otherwise inaccessible workbooks may fail. Failures are reported, and processing continues with the remaining files.

## Testing status

Basic operation and cancellation before processing have been tested with the Japanese version in Excel 2013 on Windows. The English VBA module in this repository has not been run, and the behavior has not been tested in an English-language Excel installation or with every workbook format and condition.
