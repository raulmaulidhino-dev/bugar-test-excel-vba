VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PFT_Form2 
   Caption         =   "Physical Fitness Test Form (2 of 3)"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10155
   OleObjectBlob   =   "PFT_Form2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PFT_Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-----------------------------------------------------------------------------------
' Keypress Rules in Each TextBox Input
Private Sub Run1200mMinuteCombo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
Private Sub Run1200mSecondCombo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
Private Sub Run60mComboInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub SitUpInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
Private Sub HandEyeCoordinationInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
Private Sub HexagonalAgilityInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub StorkBalanceInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

'------------------------------------------------------------------------------------
' Previous Button
Private Sub PreviousButton_Click()
    PFT_Form2.Hide
    InputTKJ1.Show
End Sub

' Cancel Button
Private Sub CancelButton_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to cancel this test? All data you have filled in will be lost and cannot be recovered!", vbYesNo + vbCritical, "Confirmation")
    If response = vbYes Then
        Unload Me
    End If
End Sub

'Tombol Lanjut
Private Sub NextButton_Click()
    Dim Run1200mTime As Integer
    Dim totalDb As Integer
    totalDb = ThisWorkbook.Worksheets("Tools").Range("totalDatabase")
    
    If Run1200mMinuteCombo.Value = "" Then
        MsgBox "There are still empty input(s)!", vbExclamation + vbOKOnly
    ElseIf Run1200mSecondCombo.Value = "" Then
        MsgBox "There are still empty input(s)!", vbExclamation + vbOKOnly
    ElseIf Run60mComboInput.Value = "" Then
        MsgBox "There are still empty input(s)!", vbExclamation + vbOKOnly
    ElseIf SitUpInput.Value = "" Then
        MsgBox "There are still empty input(s)!", vbExclamation + vbOKOnly
    ElseIf HandEyeCoordinationInput.Value = "" Then
        MsgBox "There are still empty input(s)!", vbExclamation + vbOKOnly
    ElseIf HexagonalAgilityInput.Value = "" Then
        MsgBox "There are still empty input(s)!", vbExclamation + vbOKOnly
    ElseIf StorkBalanceInput.Value = "" Then
        MsgBox "There are still empty input(s)!", vbExclamation + vbOKOnly
    Else
        Dim run1200mTimeValue As Date
        Dim formattedTime As String
        
        run1200mTimeValue = TimeSerial(0, Run1200mMinuteCombo.Value, Run1200mSecondCombo.Value)
        formattedTime = Format(run1200mTimeValue, "hh:mm:ss")
        
        ThisWorkbook.Worksheets("Result").Range("Run1200mNumberOutput") = formattedTime
        
        ThisWorkbook.Worksheets("Result").Range("Run60mNumberOutput") = Round(Run60mComboInput.Value, 1)
        ThisWorkbook.Worksheets("Result").Range("SitUpNumberOutput") = SitUpInput.Value
        ThisWorkbook.Worksheets("Result").Range("HandEyeCoordinationNumberOutput") = HandEyeCoordinationInput.Value
        ThisWorkbook.Worksheets("Result").Range("HexagonalAgilityNumberOutput") = Round(HexagonalAgilityInput.Value, 1)
        ThisWorkbook.Worksheets("Result").Range("StorkBalanceNumberOutput") = StorkBalanceInput.Value
        
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb, 7) = formattedTime
        
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb, 10) = Round(Run60mComboInput.Value, 1)
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb, 16) = SitUpInput.Value
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb, 22) = HandEyeCoordinationInput.Value
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb, 13) = Round(HexagonalAgilityInput.Value, 1)
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb, 19) = StorkBalanceInput.Value
        
        PFT_Form2.Hide
        PFT_Form3.Show
    End If
End Sub

Private Sub UserForm_Initialize()
    Run60mComboInput.MaxLength = 4
    SitUpInput.MaxLength = 2
    HandEyeCoordinationInput.MaxLength = 2
    HexagonalAgilityInput.MaxLength = 4
    StorkBalanceInput.MaxLength = 2
End Sub
