Imports System.Net.Mime.MediaTypeNames

Sub RemoveFunctionsBehindCellsAndRemoveEmptyLines()
    '
    ' RemoveFunctionsAndEmptyLines Macro
    ' This macro will remove functions behind cells and remove empty lines (#N/A)
    '
    On Error GoTo CouldNotFindSheet
    Dim sheetToEdit As Worksheet
Set sheetToEdit = Sheets("Load")

Dim isNotAvailable As Boolean
    isNotAvailable = False

    sheetToEdit.Select

    'Remove functions behind cells
    sheetToEdit.Cells.Copy
    sheetToEdit.Cells.PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False

    'Remove empty rows
    Dim emptyCell As Range
Set emptyCell = sheetToEdit.Cells.Find(What:="#N/A", LookIn:=xlFormulas)

If (Not emptyCell Is Nothing) Then
        lr = sheetToEdit.Cells(sheetToEdit.Rows.Count, "A").End(xlUp).Row 'find last row
        For i = 2 To lr 'loop thru backwards, finish at 2 for headers
            If sheetToEdit.Cells(i, "A").Text = "#N/A" And sheetToEdit.Cells(i, "B").Text <> "0" Then
                MsgBox "Pas op, niet alle namen zijn gekend blijkbaar. Cell A positie " & i & ". Maak een nieuwe user aan."
isNotAvailable = True
                'go to the cell where name is not available (#N/A)
                Range("A" & i).Select


                Exit For

            End If

        Next i
    End If

    If (isNotAvailable = False) Then
        emptyCell.Activate
        lr = Cells(Rows.Count, "A").End(xlUp).Row 'find last row
        For i = lr To 2 Step -1 'loop thru backwards, finish at 2 for headers
            If Cells(i, "A").Text = "#N/A" And Cells(i, "B").Text = "0" Then Rows(i).EntireRow.Delete
        Next i

    End If




    'Put sheet on the first position of sheets
    Worksheets("Load").Move Before:=Worksheets(1)

If isNotAvailable = False Then Sheets("Startup").Select
    Exit Sub

CouldNotFindSheet:
    MsgBox "Could not find the sheet 'Load'. Make sure that your sheet is called Load."

End Sub

