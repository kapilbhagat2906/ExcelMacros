Sub GroupAndfilterDataIntoSeparateWorksheets()
'
' GroupAndfilterDataIntoSeparateWorksheets Macro
'
' Keyboard Shortcut: Ctrl+Shift+W
'
    Dim InitialDataSheet As Worksheet
    Dim InitialDataWorkBook As Workbook

    RowCount = Application.WorksheetFunction.CountA(Range("A:A"))
    LastColumnIndex = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    uniqueDataColumnIndex = LastColumnIndex + 5
    UniqueDataColumnLetter = Split(Cells(1, uniqueDataColumnIndex).Address, "$")(1)

    Set InitialDataSheet = ActiveSheet
    Set InitialDataWorkBook = ActiveWorkbook

    ActiveSheet.Range("H2:H" & RowCount).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range(UniqueDataColumnLetter & "1"), Unique:=True
    LastRowIndex = ActiveSheet.Cells(ActiveSheet.Rows.Count, UniqueDataColumnLetter).End(xlUp).Row
    ActiveSheet.Range(UniqueDataColumnLetter & "1:" & UniqueDataColumnLetter & LastRowIndex).RemoveDuplicates Columns:=Array(1), Header:=xlNo
    LastRowIndex = ActiveSheet.Cells(ActiveSheet.Rows.Count, UniqueDataColumnLetter).End(xlUp).Row
    UniqueValues = ActiveSheet.Range(UniqueDataColumnLetter & "1:" & UniqueDataColumnLetter & LastRowIndex)
    ActiveSheet.Columns(UniqueDataColumnLetter).ClearContents
    For Each uniqueValue In UniqueValues
        Call CreateGroupedDataSheet(uniqueValue, RowCount, InitialDataSheet, InitialDataWorkBook)
        Application.Wait (Now + TimeValue("0:00:01"))
        Next

End Sub

Private Sub CreateGroupedDataSheet(uniqueValue, RowCount, InitialDataSheet, InitialDataWorkBook)
    Dim DataWorksheet As Worksheet
    FilterColumnLetter = "H"
    InitialDataSheet.Select
    ActiveSheet.Range(FilterColumnLetter & "1:" & FilterColumnLetter & RowCount).AutoFilter Field:=1, Criteria1:=uniqueValue
    LastColumnIndex = ActiveSheet.Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
    LastColumnLetter = Split(Cells(1, LastColumnIndex).Address, "$")(1)
    LastRowIndex = ActiveSheet.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row

    Range("A1:" & LastColumnLetter & LastRowIndex).SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Set DataWorksheet = InitialDataWorkBook.Sheets.Add(After:=InitialDataWorkBook.Sheets(InitialDataWorkBook.Sheets.Count))
    DataWorksheet.Name = uniqueValue
    DataWorksheet.Select
    ActiveSheet.Paste
    Call WorksheetApplyFilter(DataWorksheet)
End Sub

Private Sub WorksheetApplyFilter(DataWorksheet)
    LastRowIndex = DataWorksheet.Cells(DataWorksheet.Rows.Count, 1).End(xlUp).Row
    FilterColumnLetter = "F"
    DataWorksheet.Range(FilterColumnLetter & "1:" & FilterColumnLetter & LastRowIndex).AutoFilter Field:=1, Criteria1:="YES"
End Sub



