# Creating-a-Macro
Begin recording the macro to format the worksheets.
    Select the North American worksheet.
    Select View→Macros→Record Macro.
    In the Record Macro dialog box, in the Macro name text box, type FormatSheet and press Tab.
    In the Shortcut key text box, press Shift+F to produce the shortcut key Ctrl+Shift+F.
    The Ctrl+ is preconfigured.
    In the Store macro in field, verify that This Workbook is selected.
    Select the Description text box, type Format worksheets and then select OK.
    Observe that the status bar displays a Stop button, indicating that a macro is currently recording.
Format the title of the worksheet.
    Select cell A1.
    On the Home tab, select the Font size drop-down arrow, and select 14.
    Select Bold.
    Select Italic.
Format the expense and quarterly headings in column A and row 5.
    Select the range A5:A10, press and hold Ctrl, and select the range B5:E5.
    On the Home tab, select Bold, Italic, and Underline.
Format the quarterly values and totals with currency formatting.
    Select the range B6:E10.
    On the Home tab, select the Number Format dialog box launcher.
    In the Format Cells dialog box, select the Currency category. Set the Decimal places to zero (0) and select OK.
    Select cell A2.
Stop recording the macro.
    Select View→Macros→Stop Recording.
    Note: You may also stop recording the macro by selecting the Stop Recording button on the status bar.
Use the Macro dialog box to run the macro on the European worksheet.
    Select the European worksheet.
    Select View→Macros→View Macros.
    In the Macro dialog box, verify that the FormatSheet macro is selected and select Run.
    The formatting is applied to the current worksheet.
Use the assigned key combination to run the macro on the Australian worksheet.
    Select the Australian worksheet.
    Press Ctrl+Shift+F.
    Observe that the macro runs and the formatting is applied to the Australian worksheet.
Create a macro that clears all formatting.
    Select the North American worksheet.
    Select the View→Macros→Record Macro.
    In the Record Macro dialog box, in the Macro name text box, type ClearFormats and press Tab.
    In the Shortcut key text box, press Shift+C to produce the shortcut key Ctrl+Shift+C.
    In the Store macro in field, verify that This Workbook is selected.
    Select the Description text box, then type Clear worksheet formatting and select OK.
Record the steps to clear the formatting.
    Select the Select All icon in the upper-left corner of the worksheet.
    On the Home tab, select the Clear drop-down arrow, and then select Clear Formats.
    Press Ctrl+Home.
    Select View→Macros→Stop Recording.
Save the workbook as a macro-enabled workbook.
    Select File→Save As.
    From the Save as type drop-down list, select Excel Macro-Enabled Workbook (*.xlsm).
