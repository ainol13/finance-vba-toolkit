Attribute VB_Name = "modXLookup"
Option Explicit

'===============================================================================
' LOOKUP PROFILE CONFIGURATION
'
' Each profile defines:
'   - The master workbook and worksheet
'   - The master lookup and return columns
'   - The active worksheet lookup and output columns
'   - The first data row in each worksheet
'===============================================================================
Public Const PROFILE_MAX As Long = 10
Public Type LookupProfile
    name As String

    MasterPath As String
    MasterSheet As String
    MasterKeyCol As String
    MasterRetFirstCol As String
    MasterRetLastCol As String
    MasterStartRow As Long

    ActiveKeyCol As String
    ActiveOutFirstCol As String
    ActiveOutLastCol As String
    ActiveStartRow As Long
End Type

Private Const MISSING_SHEET_NAME As String = "_MissingFromMaster"

'===============================================================================
' ENTRY POINT
'
' Displays the profile selector, then either:
'   - Runs the lookup using the selected profile, or
'   - Opens the selected master workbook
'===============================================================================
Public Sub XLOOKUP_VBA()
    Dim p As LookupProfile
    Dim idx As Long
    Dim action As String

    idx = ShowMasterSelector(action)
    If idx = 0 Then Exit Sub

    p = GetProfile(idx)

    If action = "OPENFILE" Then
        OpenMasterFile p.MasterPath
        Exit Sub
    End If

    RunLookupUsingProfile p
End Sub



'===============================================================================
' PROFILE DEFINITIONS
'
' These profiles use fictional paths and names for documentation purposes.
' Replace them with approved locations in private production code.
'
' PROFILE_MAX must be greater than or equal to the highest profile index.
'===============================================================================

Public Function GetProfile(ByVal idx As Long) As LookupProfile
    Dim p As LookupProfile

    Select Case idx

        Case 1
            With p
                .name = "Example Lookup A"
                .MasterPath = _
                    "\\SERVER\SHARE\ReferenceData\ExampleMasterA.xlsx"
                .MasterSheet = "LookupData"
                .MasterKeyCol = "A"
                .MasterRetFirstCol = "C"
                .MasterRetLastCol = "D"
                .MasterStartRow = 2

                .ActiveKeyCol = "C"
                .ActiveOutFirstCol = "E"
                .ActiveOutLastCol = "F"
                .ActiveStartRow = 2
            End With

        Case 2
            With p
                .name = "Example Lookup B"
                .MasterPath = _
                    "C:\ReferenceData\ExampleMasterB.xlsx"
                .MasterSheet = "ReferenceList"
                .MasterKeyCol = "A"
                .MasterRetFirstCol = "C"
                .MasterRetLastCol = "C"
                .MasterStartRow = 2

                .ActiveKeyCol = "C"
                .ActiveOutFirstCol = "G"
                .ActiveOutLastCol = "G"
                .ActiveStartRow = 2
            End With

        Case 3
            With p
                .name = "Example Lookup C"
                .MasterPath = _
                    "\\SERVER\SHARE\ReferenceData\ExampleMasterC.xlsx"
                .MasterSheet = "SiteData"
                .MasterKeyCol = "B"
                .MasterRetFirstCol = "D"
                .MasterRetLastCol = "F"
                .MasterStartRow = 2

                .ActiveKeyCol = "D"
                .ActiveOutFirstCol = "H"
                .ActiveOutLastCol = "J"
                .ActiveStartRow = 3
            End With

        Case Else
            Err.Raise _
                vbObjectError + 999, _
                "GetProfile", _
                "Invalid profile index."
    End Select

    GetProfile = p
End Function

Public Function ShowMasterSelector(Optional ByRef action As String = "CANCEL") As Long
    Dim f As frmMasterSelector
    Set f = New frmMasterSelector

    f.Show vbModal

    ShowMasterSelector = f.SelectedIndex
    action = f.SelectedAction

    Unload f
    Set f = Nothing
End Function

'===============================================================================
' LOOKUP PROCESS
'
' Loads the selected master data into memory, performs the lookup, writes the
' results to the active worksheet, and reports keys not found in the master.
'===============================================================================
Private Sub RunLookupUsingProfile(ByRef p As LookupProfile)

    Dim wbMaster As Workbook
    Dim wsMaster As Worksheet
    Dim wbActive As Workbook
    Dim wsActive As Worksheet

    Dim dict As Object
    Dim dictMissing As Object

    Dim lastRowMaster As Long
    Dim lastRowActive As Long

    Dim keysMaster As Variant
    Dim valsMaster As Variant
    Dim keysActive As Variant
    Dim outArr() As Variant

    Dim i As Long
    Dim j As Long
    Dim k As String
    Dim tmp As Variant

    Dim masterWasAlreadyOpen As Boolean
    Dim masterOpenedByMacro As Boolean
    Dim keepMasterOpen As Boolean

    Dim prevCalc As XlCalculation
    Dim prevScreenUpdating As Boolean
    Dim prevEvents As Boolean
    Dim prevAlerts As Boolean
    Dim prevAskLinks As Boolean
    Dim appStateSaved As Boolean

    Dim retCols As Long
    Dim outCols As Long

    Dim rngHdrDest As Range
    Dim rngHdrSrc As Range
    Dim col As Long
    Dim destRow As Long
    Dim srcRow As Long

    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo CleanFail
    
    'Capture the workbook and worksheet that will receive the lookup results.
    Set wbActive = ActiveWorkbook
    Set wsActive = ActiveSheet

    'Save the current Excel application settings before disabling them.
    prevCalc = Application.Calculation
    prevScreenUpdating = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevAlerts = Application.DisplayAlerts
    prevAskLinks = Application.AskToUpdateLinks
    appStateSaved = True

    'Improve performance and prevent prompts while the lookup is running.
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    Application.Calculation = xlCalculationManual

    'Confirm that the master return range and active output range have
    'the same number of columns.
    retCols = ColumnSpanCount( _
        p.MasterRetFirstCol, _
        p.MasterRetLastCol _
    )

    outCols = ColumnSpanCount( _
        p.ActiveOutFirstCol, _
        p.ActiveOutLastCol _
    )

    If retCols <> outCols Then
        Err.Raise vbObjectError + 200, , _
            "The output-column configuration is invalid for '" & _
            p.name & "'." & vbCrLf & vbCrLf & _
            "Master return columns: " & _
            p.MasterRetFirstCol & ":" & p.MasterRetLastCol & _
            " (" & retCols & ")" & vbCrLf & _
            "Active output columns: " & _
            p.ActiveOutFirstCol & ":" & p.ActiveOutLastCol & _
            " (" & outCols & ")" & vbCrLf & vbCrLf & _
            "Both ranges must contain the same number of columns."
    End If
    
    'Reuse the master workbook when it is already open.
    Set wbMaster = GetWorkbookIfOpen(p.MasterPath)

    If wbMaster Is Nothing Then
        masterWasAlreadyOpen = False
        masterOpenedByMacro = True

        'Open the master in the background as read-only.
        Set wbMaster = Workbooks.Open( _
            Filename:=p.MasterPath, _
            ReadOnly:=True, _
            UpdateLinks:=0, _
            AddToMru:=False _
        )

        If wbMaster.Windows.Count > 0 Then
            wbMaster.Windows(1).Visible = False
        End If
    Else
        masterWasAlreadyOpen = True
        masterOpenedByMacro = False
    End If

    Set wsMaster = wbMaster.Worksheets(p.MasterSheet)

    'Copy each master header only when the corresponding destination
    'header cell is blank.
    If p.MasterStartRow > 1 And p.ActiveStartRow > 1 Then

        destRow = p.ActiveStartRow - 1
        srcRow = p.MasterStartRow - 1

        'Load the master keys and return values into memory.
        Set rngHdrDest = wsActive.Range( _
            p.ActiveOutFirstCol & destRow & ":" & _
            p.ActiveOutLastCol & destRow _
        )

        Set rngHdrSrc = wsMaster.Range( _
            p.MasterRetFirstCol & srcRow & ":" & _
            p.MasterRetLastCol & srcRow _
        )

        For col = 1 To rngHdrDest.Columns.Count
            If Len(Trim$(CStr(rngHdrDest.Cells(1, col).Value))) = 0 Then
                rngHdrDest.Cells(1, col).Value = _
                    rngHdrSrc.Cells(1, col).Value
            End If
        Next col
    End If

    'Create a dictionary for fast exact-match lookups.
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    lastRowMaster = wsMaster.Cells( _
        wsMaster.Rows.Count, _
        p.MasterKeyCol _
    ).End(xlUp).row

    If lastRowMaster < p.MasterStartRow Then
        Err.Raise vbObjectError + 100, , _
            "No lookup data was found in master worksheet '" & _
            p.MasterSheet & "'." & vbCrLf & _
            "Expected data to begin on row " & p.MasterStartRow & "."
    End If


    'Always return two-dimensional arrays, including for one-cell ranges.
    keysMaster = RangeTo2DArray( _
        wsMaster.Range( _
            p.MasterKeyCol & p.MasterStartRow & ":" & _
            p.MasterKeyCol & lastRowMaster _
        ) _
    )

    valsMaster = RangeTo2DArray( _
        wsMaster.Range( _
            p.MasterRetFirstCol & p.MasterStartRow & ":" & _
            p.MasterRetLastCol & lastRowMaster _
        ) _
    )

    For i = 1 To UBound(keysMaster, 1)

        If IsError(keysMaster(i, 1)) Then
            k = vbNullString
        Else
            k = CStr(keysMaster(i, 1))
        End If

        If Len(k) > 0 Then
            dict(k) = GetRowVector(valsMaster, i, retCols)
        End If
    Next i

    'Find the last active key row.
    lastRowActive = wsActive.Cells( _
        wsActive.Rows.Count, _
        p.ActiveKeyCol _
    ).End(xlUp).row

    If lastRowActive < p.ActiveStartRow Then
        GoTo CleanExit
    End If

    'Always return a two-dimensional array for active keys.
    keysActive = RangeTo2DArray( _
        wsActive.Range( _
            p.ActiveKeyCol & p.ActiveStartRow & ":" & _
            p.ActiveKeyCol & lastRowActive _
        ) _
    )

    ReDim outArr( _
        1 To UBound(keysActive, 1), _
        1 To retCols _
    )

    Set dictMissing = CreateObject("Scripting.Dictionary")
    dictMissing.CompareMode = vbBinaryCompare

    'Perform the lookup.
    For i = 1 To UBound(keysActive, 1)

        If IsError(keysActive(i, 1)) Then
            k = vbNullString
        Else
            k = CStr(keysActive(i, 1))
        End If

        If Len(k) = 0 Then

            For j = 1 To retCols
                outArr(i, j) = vbNullString
            Next j

        ElseIf dict.exists(k) Then

            tmp = dict(k)

            For j = 1 To retCols
                outArr(i, j) = tmp(j)
            Next j

        Else

            For j = 1 To retCols
                outArr(i, j) = CVErr(xlErrNA)
            Next j

            If Not dictMissing.exists(k) Then
                dictMissing.Add k, True
            End If

        End If
    Next i

    'Write all lookup results in one operation.
    wsActive.Range( _
        p.ActiveOutFirstCol & p.ActiveStartRow & ":" & _
        p.ActiveOutLastCol & lastRowActive _
    ).Value = outArr

    'Handle missing keys.
    If dictMissing.Count > 0 Then

        WriteMissingKeysSheet wbActive, dictMissing, p.name

        Application.ScreenUpdating = True

        If Not masterWasAlreadyOpen Then

            'Close the hidden read-only instance.
            wbMaster.Close SaveChanges:=False
            Set wbMaster = Nothing
            masterOpenedByMacro = False

            'Reopen the master normally for manual editing.
            Set wbMaster = Workbooks.Open( _
                Filename:=p.MasterPath, _
                ReadOnly:=False, _
                UpdateLinks:=0, _
                AddToMru:=False _
            )

            'This workbook is intentionally left open.
            keepMasterOpen = True
            ShowWorkbook wbMaster

            If wbMaster.ReadOnly Then
                MsgBox _
                    "The master workbook was opened as read-only." & vbCrLf & vbCrLf & _
                    "Another user may currently have the file open, or you may not " & _
                    "have permission to edit it." & vbCrLf & vbCrLf & _
                    "The macro did not save or modify the master workbook.", _
                    vbExclamation + vbOKOnly, _
                    "Master Workbook Is Read-Only"
            End If

        Else

            keepMasterOpen = True
            ShowWorkbook wbMaster

            If wbMaster.ReadOnly Then
                MsgBox _
                    "The master workbook is already open as read-only." & _
                    vbCrLf & vbCrLf & _
                    "To edit it, close the workbook and reopen it when an editable " & _
                    "copy is available." & vbCrLf & vbCrLf & _
                    "The macro did not save or modify the master workbook.", _
                    vbExclamation + vbOKOnly, _
                    "Master Workbook Is Read-Only"
            End If

        End If

        MsgBox _
            Format$(dictMissing.Count, "#,##0") & _
            IIf(dictMissing.Count = 1, _
                " key was", _
                " keys were") & _
            " not found in the selected master." & vbCrLf & vbCrLf & _
            "Profile: " & p.name & vbCrLf & _
            "Missing-key list: " & MISSING_SHEET_NAME & vbCrLf & vbCrLf & _
            "The master workbook has been opened so the missing records can " & _
            "be reviewed or added." & vbCrLf & vbCrLf & _
            "Any changes must be saved manually.", _
            vbInformation + vbOKOnly, _
            "Lookup Completed with Missing Keys"

        GoTo CleanExit
    End If

CleanExit:
'Close the master only when it was opened by this procedure and
'has not intentionally been left open for manual maintenance.

    On Error Resume Next

    'Centralized workbook cleanup.
    'This also closes the master when there are no active data rows.
    If masterOpenedByMacro Then
        If Not keepMasterOpen Then
            If Not wbMaster Is Nothing Then
                wbMaster.Close SaveChanges:=False
            End If
        End If
    End If

    Set wsMaster = Nothing
    Set wbMaster = Nothing

    If appStateSaved Then
        Application.Calculation = prevCalc
        Application.ScreenUpdating = prevScreenUpdating
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        Application.AskToUpdateLinks = prevAskLinks
    End If

    On Error GoTo 0
    Exit Sub

CleanFail:
    'Preserve the original error before attempting any cleanup.
    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next

    'Close only the master workbook opened by this procedure.
    'Never close a workbook that was already open before the lookup started.
    If masterOpenedByMacro Then
        If Not wbMaster Is Nothing Then
            wbMaster.Close SaveChanges:=False
        End If
    End If

    Set wsMaster = Nothing
    Set wbMaster = Nothing

    'Restore the Excel application settings before displaying the error.
    If appStateSaved Then
        Application.Calculation = prevCalc
        Application.ScreenUpdating = prevScreenUpdating
        Application.EnableEvents = prevEvents
        Application.DisplayAlerts = prevAlerts
        Application.AskToUpdateLinks = prevAskLinks
    End If

    On Error GoTo 0

    MsgBox _
        "The lookup could not be completed." & vbCrLf & vbCrLf & _
        "Profile: " & p.name & vbCrLf & _
        "Error " & errorNumber & ": " & errorDescription & vbCrLf & vbCrLf & _
        "Excel settings have been restored." & vbCrLf & _
        "No files were saved by the macro.", _
        vbCritical + vbOKOnly, _
        "Lookup Failed"

End Sub


'===============================================================================
' SUPPORTING PROCEDURES
'===============================================================================
Private Function GetWorkbookIfOpen(ByVal fullPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set GetWorkbookIfOpen = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub ShowWorkbook(ByVal wb As Workbook)
    On Error Resume Next
    If wb.Windows.Count = 0 Then wb.Activate
    If wb.Windows.Count > 0 Then
        wb.Windows(1).Visible = True
        wb.Windows(1).WindowState = xlNormal
    End If
    wb.Activate
    Application.GoTo wb.Worksheets(1).Range("A1"), True
    On Error GoTo 0
End Sub

Private Function ColumnSpanCount(ByVal firstCol As String, ByVal lastCol As String) As Long
    ColumnSpanCount = ColLetterToNum(lastCol) - ColLetterToNum(firstCol) + 1
End Function

Private Function ColLetterToNum(ByVal colLetters As String) As Long
    Dim i As Long, n As Long
    colLetters = UCase$(Trim$(colLetters))
    For i = 1 To Len(colLetters)
        n = n * 26 + (Asc(Mid$(colLetters, i, 1)) - Asc("A") + 1)
    Next i
    ColLetterToNum = n
End Function

Private Function GetRowVector(ByVal arr2D As Variant, ByVal rowIndex As Long, ByVal retCols As Long) As Variant
    Dim v() As Variant, j As Long
    ReDim v(1 To retCols)
    For j = 1 To retCols
        v(j) = arr2D(rowIndex, j)
    Next j
    GetRowVector = v
End Function

Private Sub WriteMissingKeysSheet(ByVal wb As Workbook, ByVal dictMissing As Object, ByVal profileName As String)
    Dim ws As Worksheet
    Dim k As Variant
    Dim r As Long

    On Error Resume Next
    Set ws = wb.Worksheets(MISSING_SHEET_NAME)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.name = MISSING_SHEET_NAME

    ws.Range("A1").Value = "Missing keys (not found in master): " & profileName
    ws.Range("A1").Font.Bold = True

    r = 2
    For Each k In dictMissing.keys
        ws.Cells(r, 1).Value = k
        r = r + 1
    Next k

    ws.Columns("A:A").AutoFit
End Sub

Private Sub OpenMasterFile(ByVal masterFullPath As String)
    Dim wb As Workbook

    On Error GoTo OpenFail

    'Validate the configured path before attempting to open the workbook.
    If Len(Trim$(masterFullPath)) = 0 Then
        MsgBox _
            "No master workbook path has been configured for this profile.", _
            vbExclamation + vbOKOnly, _
            "Master Path Not Configured"
        Exit Sub
    End If

    'If the workbook is already open, display and activate it.
    Set wb = GetWorkbookIfOpen(masterFullPath)

    If Not wb Is Nothing Then
        ShowWorkbook wb

        If wb.ReadOnly Then
            MsgBox _
                "The master workbook is already open as read-only." & _
                vbCrLf & vbCrLf & _
                "Another user may currently have the file open, or you may " & _
                "not have permission to edit it." & vbCrLf & vbCrLf & _
                "To edit the workbook, reopen it when an editable copy is " & _
                "available.", _
                vbExclamation + vbOKOnly, _
                "Master Workbook Is Read-Only"
        End If

        Exit Sub
    End If

    'Open the workbook for editing. Excel may open it as read-only if the
    'file is locked, already in use, or protected by access permissions.
    Set wb = Workbooks.Open( _
        Filename:=masterFullPath, _
        ReadOnly:=False, _
        UpdateLinks:=0, _
        AddToMru:=False _
    )

    ShowWorkbook wb

    If wb.ReadOnly Then
        MsgBox _
            "The master workbook was opened as read-only." & _
            vbCrLf & vbCrLf & _
            "Another user may currently have the file open, or you may " & _
            "not have permission to edit it." & vbCrLf & vbCrLf & _
            "To edit the workbook, try again when an editable copy is " & _
            "available.", _
            vbExclamation + vbOKOnly, _
            "Master Workbook Is Read-Only"
    End If

    Exit Sub

OpenFail:
    MsgBox _
        "The master workbook could not be opened." & vbCrLf & vbCrLf & _
        "File:" & vbCrLf & masterFullPath & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbCritical + vbOKOnly, _
        "Unable to Open Master Workbook"
End Sub

Private Function RangeTo2DArray(ByVal rng As Range) As Variant
    Dim sourceValue As Variant
    Dim result() As Variant
    Dim r As Long
    Dim c As Long

    sourceValue = rng.Value

    ReDim result(1 To rng.Rows.Count, 1 To rng.Columns.Count)

    If rng.Cells.CountLarge = 1 Then
        'A one-cell range returns a scalar, so convert it to a 2D array.
        result(1, 1) = sourceValue
    Else
        For r = 1 To rng.Rows.Count
            For c = 1 To rng.Columns.Count
                result(r, c) = sourceValue(r, c)
            Next c
        Next r
    End If

    RangeTo2DArray = result
End Function


