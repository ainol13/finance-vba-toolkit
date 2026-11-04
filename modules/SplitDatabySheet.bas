Attribute VB_Name = "SplitDatabySheet"
Option Explicit

Public Sub SplitDataIntoSheetsBySelectedColumn()

    Dim sourceWS As Worksheet
    Dim targetWS As Worksheet
    Dim logWS As Worksheet
    Dim workbookObject As Workbook

    Dim columnInput As Variant
    Dim groupingColumn As Long
    Dim lastRow As Long
    Dim lastColumn As Long

    Dim dataRange As Range
    Dim visibleDataRange As Range
    Dim lastUsedCell As Range
    Dim cell As Range

    Dim uniqueValues As Object
    Dim sheetNamesUsed As Object
    Dim blankGroupRows As Collection

    Dim groupValue As Variant
    Dim newSheetName As String
    Dim logSheetName As String

    Dim columnNumber As Long
    Dim sourceRowNumber As Variant
    Dim logRow As Long

    Dim originalCalculation As XlCalculation
    Dim sheetsCreated As Long
    Dim loggedRows As Long
    Dim processCompleted As Boolean

    Set sourceWS = ActiveSheet
    Set workbookObject = sourceWS.Parent

    'Ask the user which column should be used to group the data.
    columnInput = Application.InputBox( _
        Prompt:="Enter the column to use for grouping the data." & vbCrLf & _
                "For example: A, B, C, or AA", _
        title:="Select Grouping Column", _
        Default:="A", _
        Type:=2)

    'Exit if the user clicks Cancel.
    If VarType(columnInput) = vbBoolean Then
        If columnInput = False Then Exit Sub
    End If

    columnInput = UCase$(Trim$(CStr(columnInput)))

    If Len(columnInput) = 0 Then
        MsgBox "Please enter a valid column letter.", _
               vbExclamation, _
               "Invalid Column"
        Exit Sub
    End If

    'Convert the entered column letter into a column number.
    On Error Resume Next
    groupingColumn = sourceWS.Range(columnInput & "1").Column
    On Error GoTo ErrorHandler

    If groupingColumn = 0 Then
        MsgBox "The column entered is not valid.", _
               vbExclamation, _
               "Invalid Column"
        Exit Sub
    End If

    'Find the final used row anywhere in the source worksheet.
    Set lastUsedCell = sourceWS.Cells.Find( _
        What:="*", _
        After:=sourceWS.Range("A1"), _
        LookAt:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious)

    If lastUsedCell Is Nothing Then
        MsgBox "The source worksheet is empty.", _
               vbInformation, _
               "No Data"
        Exit Sub
    End If

    lastRow = lastUsedCell.row

    If lastRow < 2 Then
        MsgBox "No data was found below the header row.", _
               vbInformation, _
               "No Data"
        Exit Sub
    End If

    'Find the final used column anywhere in the source worksheet.
    Set lastUsedCell = sourceWS.Cells.Find( _
        What:="*", _
        After:=sourceWS.Range("A1"), _
        LookAt:=xlPart, _
        LookIn:=xlFormulas, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious)

    If lastUsedCell Is Nothing Then
        MsgBox "The source worksheet is empty.", _
               vbInformation, _
               "No Data"
        Exit Sub
    End If

    lastColumn = lastUsedCell.Column

    'Confirm that the selected grouping column is within
    'the source data range.
    If groupingColumn > lastColumn Then
        MsgBox "The selected grouping column is outside " & _
               "the source data range.", _
               vbExclamation, _
               "Invalid Grouping Column"
        Exit Sub
    End If

    'The complete source range, including the header.
    Set dataRange = sourceWS.Range( _
        sourceWS.Cells(1, 1), _
        sourceWS.Cells(lastRow, lastColumn))

    Set uniqueValues = CreateObject("Scripting.Dictionary")
    uniqueValues.CompareMode = vbTextCompare

    Set sheetNamesUsed = CreateObject("Scripting.Dictionary")
    sheetNamesUsed.CompareMode = vbTextCompare

    'Store source row numbers that have a blank grouping cell
    'but contain information elsewhere in the row.
    Set blankGroupRows = New Collection

    'Review every source row beginning from row 2.
    For Each cell In sourceWS.Range( _
        sourceWS.Cells(2, groupingColumn), _
        sourceWS.Cells(lastRow, groupingColumn))

        groupValue = Trim$(CStr(cell.value))

        If Len(groupValue) > 0 Then

            'Store each unique grouping value.
            If Not uniqueValues.exists(groupValue) Then
                uniqueValues.Add groupValue, groupValue
            End If

        Else

            'Only log the row if another column contains data.
            If RowHasOtherData( _
                sourceWS, _
                cell.row, _
                groupingColumn, _
                lastColumn) Then

                blankGroupRows.Add cell.row

            End If

        End If

    Next cell

    If uniqueValues.Count = 0 And blankGroupRows.Count = 0 Then
        MsgBox "No grouping values or review items were found.", _
               vbInformation, _
               "Nothing to Process"
        Exit Sub
    End If

    originalCalculation = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    'Remove any existing filter from the source worksheet.
    If sourceWS.AutoFilterMode Then
        sourceWS.AutoFilterMode = False
    End If

    '------------------------------------------------------
    'Create the review log when blank grouping cells exist.
    '------------------------------------------------------
    If blankGroupRows.Count > 0 Then

        logSheetName = GetUniqueSheetName( _
            "Split Log", _
            workbookObject, _
            sheetNamesUsed)

        Set logWS = workbookObject.Worksheets.Add( _
            After:=workbookObject.Worksheets( _
                workbookObject.Worksheets.Count))

        logWS.name = logSheetName

        CreateBlankGroupLogHeader _
            sourceWS:=sourceWS, _
            logWS:=logWS, _
            lastColumn:=lastColumn

        logRow = 2

        For Each sourceRowNumber In blankGroupRows

            'Record details about the source row.
            logWS.Cells(logRow, 1).value = sourceWS.name
            logWS.Cells(logRow, 2).value = CLng(sourceRowNumber)
            logWS.Cells(logRow, 3).value = _
                "Grouping cell is blank, but other data exists in the row."

            'Copy the complete source row, including its formatting.
            sourceWS.Range( _
                sourceWS.Cells(CLng(sourceRowNumber), 1), _
                sourceWS.Cells(CLng(sourceRowNumber), lastColumn)).Copy

            logWS.Cells(logRow, 4).PasteSpecial Paste:=xlPasteAll

            Application.CutCopyMode = False

            'Preserve the original source row height.
            logWS.rows(logRow).RowHeight = _
                sourceWS.rows(CLng(sourceRowNumber)).RowHeight

            logRow = logRow + 1
            loggedRows = loggedRows + 1

        Next sourceRowNumber

        'Add a filter to the complete log header.
        logWS.Range( _
            logWS.Cells(1, 1), _
            logWS.Cells(1, lastColumn + 3)).AutoFilter

        'Freeze the log header.
        logWS.Activate
        logWS.Range("A2").Select
        ActiveWindow.FreezePanes = True

        'Keep the log columns readable.
        logWS.Columns(1).ColumnWidth = 20
        logWS.Columns(2).ColumnWidth = 12
        logWS.Columns(3).ColumnWidth = 55

        logWS.Columns(1).HorizontalAlignment = xlLeft
        logWS.Columns(2).HorizontalAlignment = xlCenter
        logWS.Columns(3).WrapText = True

    End If

    '------------------------------------------------------
    'Create one worksheet for each unique grouping value.
    '------------------------------------------------------
    For Each groupValue In uniqueValues.keys

        newSheetName = GetUniqueSheetName( _
            CleanSheetName(CStr(groupValue)), _
            workbookObject, _
            sheetNamesUsed)

        Set targetWS = workbookObject.Worksheets.Add( _
            After:=workbookObject.Worksheets( _
                workbookObject.Worksheets.Count))

        targetWS.name = newSheetName

        'Copy the source header with all formatting.
        sourceWS.Range( _
            sourceWS.Cells(1, 1), _
            sourceWS.Cells(1, lastColumn)).Copy

        targetWS.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
        targetWS.Cells(1, 1).PasteSpecial Paste:=xlPasteColumnWidths

        Application.CutCopyMode = False

        'Preserve the exact source header row height.
        targetWS.rows(1).RowHeight = sourceWS.rows(1).RowHeight

        'Preserve the exact width of every source column.
        For columnNumber = 1 To lastColumn

            targetWS.Columns(columnNumber).ColumnWidth = _
                sourceWS.Columns(columnNumber).ColumnWidth

            targetWS.Columns(columnNumber).Hidden = _
                sourceWS.Columns(columnNumber).Hidden

        Next columnNumber

        'Filter the source data using the current grouping value.
        dataRange.AutoFilter _
            Field:=groupingColumn - dataRange.Column + 1, _
            Criteria1:=EscapeFilterCriteria(CStr(groupValue))

        Set visibleDataRange = Nothing

        On Error Resume Next

        'Copy only visible data rows.
        'The header has already been copied separately.
        Set visibleDataRange = sourceWS.Range( _
            sourceWS.Cells(2, 1), _
            sourceWS.Cells(lastRow, lastColumn)) _
            .SpecialCells(xlCellTypeVisible)

        On Error GoTo ErrorHandler

        If Not visibleDataRange Is Nothing Then

            visibleDataRange.Copy

            targetWS.Cells(2, 1).PasteSpecial Paste:=xlPasteAll

            Application.CutCopyMode = False

            'Add filter controls to the copied header.
            targetWS.Range( _
                targetWS.Cells(1, 1), _
                targetWS.Cells(1, lastColumn)).AutoFilter

            sheetsCreated = sheetsCreated + 1

        End If

    Next groupValue

    processCompleted = True

CleanExit:

    On Error Resume Next

    If sourceWS.AutoFilterMode Then
        sourceWS.AutoFilterMode = False
    End If

    sourceWS.Activate

    Application.CutCopyMode = False
    Application.Calculation = originalCalculation
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    On Error GoTo 0

    If processCompleted Then

        Dim completionMessage As String

        completionMessage = _
            sheetsCreated & " grouped worksheet(s) were created."

        If loggedRows > 0 Then
            completionMessage = completionMessage & vbCrLf & vbCrLf & _
                loggedRows & " row(s) with a blank grouping cell " & _
                "were added to the '" & logSheetName & _
                "' worksheet for review."
        Else
            completionMessage = completionMessage & vbCrLf & vbCrLf & _
                "No rows with blank grouping cells required review."
        End If

        MsgBox completionMessage, _
               vbInformation, _
               "Split Complete"

    End If

    Exit Sub

ErrorHandler:

    MsgBox "The data could not be split into worksheets." & _
           vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, _
           "Split Data Error"

    Resume CleanExit

End Sub


Private Sub CreateBlankGroupLogHeader( _
    ByVal sourceWS As Worksheet, _
    ByVal logWS As Worksheet, _
    ByVal lastColumn As Long)

    Dim columnNumber As Long

    'Create the log information headers.
    With logWS

        .Cells(1, 1).value = "Source Worksheet"
        .Cells(1, 2).value = "Source Row"
        .Cells(1, 3).value = "Review Reason"

        With .Range(.Cells(1, 1), .Cells(1, 3))

            .Font.Bold = True
            .Font.color = RGB(255, 255, 255)

            .Interior.color = RGB(192, 80, 77)

            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True

        End With

    End With

    'Copy the original source headers and their formatting.
    sourceWS.Range( _
        sourceWS.Cells(1, 1), _
        sourceWS.Cells(1, lastColumn)).Copy

    logWS.Cells(1, 4).PasteSpecial Paste:=xlPasteAll

    Application.CutCopyMode = False

    'Preserve source column widths in the copied-data section.
    For columnNumber = 1 To lastColumn

        logWS.Columns(columnNumber + 3).ColumnWidth = _
            sourceWS.Columns(columnNumber).ColumnWidth

        logWS.Columns(columnNumber + 3).Hidden = _
            sourceWS.Columns(columnNumber).Hidden

    Next columnNumber

    logWS.rows(1).RowHeight = sourceWS.rows(1).RowHeight

End Sub


Private Function RowHasOtherData( _
    ByVal sourceWS As Worksheet, _
    ByVal rowNumber As Long, _
    ByVal groupingColumn As Long, _
    ByVal lastColumn As Long) As Boolean

    Dim columnNumber As Long
    Dim cellValue As Variant

    For columnNumber = 1 To lastColumn

        'Do not check the grouping column itself.
        If columnNumber <> groupingColumn Then

            cellValue = sourceWS.Cells( _
                rowNumber, columnNumber).value

            'A formula is treated as information even when it
            'currently displays an empty string.
            If sourceWS.Cells(rowNumber, columnNumber).HasFormula Then
                RowHasOtherData = True
                Exit Function
            End If

            'Check whether the cell contains a visible value.
            If IsError(cellValue) Then

                RowHasOtherData = True
                Exit Function

            ElseIf Len(Trim$(CStr(cellValue))) > 0 Then

                RowHasOtherData = True
                Exit Function

            End If

        End If

    Next columnNumber

    RowHasOtherData = False

End Function


Private Function CleanSheetName(ByVal value As String) As String

    Dim invalidCharacters As Variant
    Dim item As Variant

    invalidCharacters = Array("\", "/", ":", "*", "?", "[", "]")

    value = Trim$(value)

    For Each item In invalidCharacters
        value = Replace(value, CStr(item), "_")
    Next item

    'Excel worksheet names cannot exceed 31 characters.
    value = Left$(value, 31)

    If Len(value) = 0 Then
        value = "Blank"
    End If

    CleanSheetName = value

End Function


Private Function GetUniqueSheetName( _
    ByVal proposedName As String, _
    ByVal workbookObject As Workbook, _
    ByVal namesCreated As Object) As String

    Dim finalName As String
    Dim suffix As String
    Dim counter As Long

    finalName = Left$(proposedName, 31)
    counter = 1

    Do While SheetExists(finalName, workbookObject) _
          Or namesCreated.exists(finalName)

        counter = counter + 1
        suffix = "_" & counter

        finalName = Left$( _
            proposedName, _
            31 - Len(suffix)) & suffix

    Loop

    namesCreated.Add finalName, True
    GetUniqueSheetName = finalName

End Function


Private Function SheetExists( _
    ByVal sheetName As String, _
    ByVal workbookObject As Workbook) As Boolean

    Dim worksheetObject As Worksheet

    On Error Resume Next
    Set worksheetObject = workbookObject.Worksheets(sheetName)
    SheetExists = Not worksheetObject Is Nothing
    On Error GoTo 0

End Function


Private Function EscapeFilterCriteria(ByVal value As String) As String

    'Ensure wildcard characters are treated as normal characters.
    value = Replace(value, "~", "~~")
    value = Replace(value, "*", "~*")
    value = Replace(value, "?", "~?")

    EscapeFilterCriteria = "=" & value

End Function

