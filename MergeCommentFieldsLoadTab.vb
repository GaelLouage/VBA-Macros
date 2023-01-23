Imports System.Net.Mime.MediaTypeNames
Imports System.Xml

Sub CopyValuesInLoadTab()
    '
    ' CopyValuesInLoadTab Macro
    '

    '
End Sub
Sub CopyDateInInvullenTab()
    '
    '
    ' FixDateInSheetInvullenTSCodes Macro
    ' will remove the function behind the Date (calc) column from Invullen TS codes
    '

    '
    On Error GoTo CouldNotFindSheet :  
    Set sheetToEdit = Sheets("Invullen TS codes")
    sheetToEdit.Select

    'Remove functions behind cells
    Range("M2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("M2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Exit Sub

    'Error message
CouldNotFindSheet:
    MsgBox "Could not find the sheet 'Invullen TS codes'. Make sure that you have a sheet called Invullen TS codes."
    Exit Sub

End Sub
Sub MergeCommentFieldsLoadTab()

    ' Merge data for a user when there are multiple rows with the same TS code, workday, workMonth and workYear
    '
    Dim currentRowUserName As Range
    Dim currentRowTsCode As Range
    Dim currentRowDay As Range
    Dim currentRowMonth As Range
    Dim currentRowYear As Range
    Dim currentRowHoursWorked As Range
    Dim currentRowComment As Range

    Dim nextRowUserName As Range
    Dim nextRowTsCode As Range
    Dim nextRowDay As Range
    Dim nextRowMonth As Range
    Dim nextRowYear As Range
    Dim nextRowHoursWorked As Range
    Dim nextRowComment As Range

    Dim tempUserName As Range
    Dim tempTsCode As Range
    Dim tempDay As Range
    Dim tempMonth As Range
    Dim tempYear As Range
    Dim tempHoursWorked As Double
    Dim tempComments As String

    Dim rowsProcessed As Integer
    Dim rowIndex As Integer


    On Error GoTo CouldNotFindSheetLoad
    Dim sheetToEdit As Worksheet
    Set sheetToEdit = Sheets("Load")
          
    sheetToEdit.Select


    'Sort data
    Range("A2").Select
    Selection.AutoFilter

    ActiveWorkbook.Worksheets("Load").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Load").AutoFilter.Sort.SortFields.Add2 Key:=Range(
        "E1:E1048576"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=
        xlSortNormal
    With ActiveWorkbook.Worksheets("Load").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Load").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Load").AutoFilter.Sort.SortFields.Add2 Key:=Range(
        "D1:D1048576"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=
        xlSortNormal
    With ActiveWorkbook.Worksheets("Load").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Load").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Load").AutoFilter.Sort.SortFields.Add2 Key:=Range(
        "C1:C1048576"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=
        xlSortNormal
    With ActiveWorkbook.Worksheets("Load").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Load").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Load").AutoFilter.Sort.SortFields.Add2 Key:=Range(
        "B1:B1048576"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=
        xlSortNormal
    With ActiveWorkbook.Worksheets("Load").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Load").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Load").AutoFilter.Sort.SortFields.Add2 Key:=Range(
        "A1:A1048576"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=
        xlSortNormal
    With ActiveWorkbook.Worksheets("Load").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    'turn off filter tool
    Selection.AutoFilter

    'Process data
    On Error Resume Next
    Range("A2").Activate
    Application.ScreenUpdating = False

    rowsProcessed = 0
    rowIndex = 2

    Do Until IsEmpty(activeCell)
        Set currentRowUserName = activeCell
        Set currentRowTsCode = activeCell.Offset(0, 1)
        Set currentRowDay = activeCell.Offset(0, 2)
        Set currentRowMonth = activeCell.Offset(0, 3)
        Set currentRowYear = activeCell.Offset(0, 4)
        Set currentRowHoursWorked = activeCell.Offset(0, 5)
        Set currentRowComment = activeCell.Offset(0, 9)

        'Is there's no next row
        If (IsEmpty(activeCell.Offset(1, 0))) Then
            If (rowsProcessed > 0) Then 'Means we'll have to merge the same rows
                Call MergeRows(activeCell, rowIndex, rowsProcessed, tempHoursWorked, tempComments)
                rowsProcessed = 0
            End If
        Else
            Set nextRowUserName = activeCell.Offset(1, 0)
            Set nextRowTsCode = activeCell.Offset(1, 1)
            Set nextRowDay = activeCell.Offset(1, 2)
            Set nextRowMonth = activeCell.Offset(1, 3)
            Set nextRowYear = activeCell.Offset(1, 4)
            Set nextRowHoursWorked = activeCell.Offset(1, 5)
            Set nextRowComment = activeCell.Offset(1, 9)
               

            If (currentRowUserName = nextRowUserName _
                And currentRowTsCode = nextRowTsCode _
                And currentRowDay = nextRowDay _
                And currentRowMonth = nextRowMonth _
                And currentRowYear = nextRowYear) Then

                'First time, add add both rows
                If (rowsProcessed = 0) Then
                    tempHoursWorked = currentRowHoursWorked + nextRowHoursWorked
                    tempComments = currentRowComment & Chr(10) & nextRowComment
                Else 'Add only data from next row
                    tempHoursWorked = tempHoursWorked + nextRowHoursWorked
                    tempComments = tempComments & Chr(10) & nextRowComment
                End If



                rowsProcessed = rowsProcessed + 1
            Else
                If (rowsProcessed > 0) Then 'Means we'll have to merge the same rows
                    Call MergeRows(activeCell, rowIndex, rowsProcessed, tempHoursWorked, tempComments)
                End If

                rowsProcessed = 0
            End If
        End If


        'Increase loop
        activeCell.Offset(1, 0).Select
        rowIndex = rowIndex + 1
    Loop


    Application.ScreenUpdating = False
    Sheets("Startup").Select
    Exit Sub

    'Error message
CouldNotFindSheetLoad:
    MsgBox "Could not find the sheet 'Load'. Make sure that you have a sheet called Load."
    Exit Sub
End Sub


Function MergeRows(activeCell As Range, rowIndex As Integer, rowsProcessed As Integer, hours As Double, comments As String)
    'delete previous lines
    Rows(rowIndex - 1 & ":" & rowIndex - rowsProcessed).Select
    Selection.Delete Shift:=xlUp

    'decrease row index since the rows are deleted
    rowIndex = rowIndex - rowsProcessed

    'update current line with new merged data
    activeCell.Offset(0, 5).Value = hours
    activeCell.Offset(0, 9).Value = comments
End Function


