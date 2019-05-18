This  Project contains macros (VBA) to help automate various tasks on Microsoft Excel.
1) GroupAndFilterData.excel.macro.bas:
    a-  This macro groups data in active worksheet in excel.
        NOTE: You need to update the column Letter that needs to be used to group and split data into separated sheets.
    b-  Based on the column letter, this macro fetches out unique data in that column, Repeat steps (c) to (e) for each unique value:
    c-  Creates worksheet with the name same as the unique value (fetched in step b)
    d-  Copies the grouped data to the new worksheet.
    e-  Applies filter on the data of worksheet populated on step (d).
        NOTE: You need to update the filter column letter and filter value.



Instructions on how to use the macros:
-   Add Developer Tab: First of all, you need to enable Developer option in excel.
    (a)-    Launch Microsoft Excel and open a spreadsheet containing a macro that you want to import into another spreadsheet.
    (b)-    Skip to the next section if you see the Developer tab on the Excel ribbon; if not, click "File" and then click "Options."
    (c)-    Click "Customize Ribbon" and move to the "Main Tabs" box, which contains a list of tabs to connect to Excel.
    (d)-    Place a check mark next to "Developer" and click "OK" to add the Developer tab to the ribbon

-   Enable Macro Security.
    (a)-    Click the "Developer" tab, then click "Macro Security" to view the Trust Center window that contains security settings.
    (b)-    Click the "Enable all macros (not recommended, potentially dangerous code can run)" radio button to select it. Selecting this option allows you to enable all macros in the worksheet temporarily.
    (c)-    Click "OK" to close the Trust Center window and return to the main Excel window containing your spreadsheet.

-   Import Macro Code Into Spreadsheet.
    (a)-    Click the "Developer" tab and then click the "Macros" button. Excel displays the Macro dialog window that contains a list of the spreadsheet's macros.
    (b)-    Click on "View Code" option.
            This will open "Microsoft Visual Basic for Applications" window with code for all the macros that already exists in spreadsheet.

            NOTE: Alternatively you can use Alt+F11 shortcut to open "Microsoft Visual Basic for Applications" window directly.
    (c)-    Click on "File" -> "Import File" submenu to import macro in your spreadsheet.
    (d)-    Search for the ".bas" file you cloned from this project.
    (e)-    Close "Microsoft Visual Basic for Applications" window.

-   Run Macro
    (a)-    Click the "Developer" tab and then click the "Macros" button. Excel displays the Macro dialog window that contains a list of the spreadsheet's macros.
    (b)-    Select the imported macro from the list -> Click on "Run" button.
    (c)-    Relax till the macro is executed.

Hurray!!!


For any query, please connect at kapil.bhagat29@gmail.com