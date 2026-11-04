VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMasterSelector 
   Caption         =   "Select Master File"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9810.001
   OleObjectBlob   =   "frmMasterSelector.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMasterSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public SelectedIndex As Long
Public SelectedAction As String   ' "LOOKUP" | "OPENFILE" | "CANCEL"



Private Sub UserForm_Initialize()
    Dim i As Long
    Dim p As LookupProfile
    Dim row As Long

    SelectedIndex = 0
    SelectedAction = "CANCEL"

    With lstProfiles
        .Clear
        .ColumnCount = 3
        .ColumnWidths = "260 pt;90 pt;0 pt"  'Name | Output | (hidden idx)
    End With

    For i = 1 To PROFILE_MAX
        Err.Clear
        On Error Resume Next
        p = GetProfile(i)
        If Err.Number = 0 Then
            If Len(p.name) > 0 Then
                lstProfiles.AddItem p.name
                row = lstProfiles.ListCount - 1
                lstProfiles.List(row, 1) = p.ActiveOutFirstCol & ":" & p.ActiveOutLastCol
                lstProfiles.List(row, 2) = CStr(i) 'hidden index
            End If
        End If
        On Error GoTo 0
    Next i

    If lstProfiles.ListCount = 0 Then
        MsgBox _
            "No lookup profiles are available." & vbCrLf & vbCrLf & _
            "Add a profile in GetProfile and ensure PROFILE_MAX includes " & _
            "its profile index.", _
            vbExclamation + vbOKOnly, _
            "No Lookup Profiles"
        SelectedIndex = 0
        Me.Hide
        Exit Sub
    End If

    lstProfiles.ListIndex = 0
End Sub

Private Sub cmdOK_Click()
    If lstProfiles.ListIndex < 0 Then
        MsgBox _
            "Select a lookup profile before continuing.", _
            vbExclamation + vbOKOnly, _
            "Profile Required"
        Exit Sub
    End If

    SelectedIndex = CLng(lstProfiles.List(lstProfiles.ListIndex, 2))
    SelectedAction = "LOOKUP"
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    SelectedIndex = 0
    SelectedAction = "CANCEL"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Treat closing the form with the title-bar X as cancellation.
    SelectedIndex = 0
    SelectedAction = "CANCEL"
End Sub

Private Sub lstProfiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOK_Click
End Sub


Private Sub cmdOpenFile_Click()
    If lstProfiles.ListIndex < 0 Then
        MsgBox _
            "Select a lookup profile before continuing.", _
            vbExclamation + vbOKOnly, _
            "Profile Required"
        Exit Sub
    End If

    SelectedIndex = CLng(lstProfiles.List(lstProfiles.ListIndex, 2))
    SelectedAction = "OPENFILE"
    Me.Hide
End Sub
