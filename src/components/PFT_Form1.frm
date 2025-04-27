VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PFT_Form1 
   Caption         =   "Physical Fitness Test Form (1 of 3)"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10155
   OleObjectBlob   =   "PFT_Form1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PFT_Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Dynamic Date Combo
Private Sub DateCombo_Change()
    Application.EnableEvents = False
    
    If Me.MonthCombo.Value = 4 Or Me.MonthCombo.Value = 6 Or Me.MonthCombo.Value = 9 And Me.MonthCombo.Value = 11 Then
        DateCombo.RowSource = "Tools!C3:C32"
    ElseIf Me.MonthCombo.Value = 2 Then
        If Me.YearCombo.Value Mod 4 = 0 Then
            DateCombo.RowSource = "Tools!C3:C31"
        Else
            DateCombo.RowSource = "Tools!C3:C30"
        End If
    Else
        DateCombo.RowSource = "Tools!C3:C33"
    End If
    
    Application.EnableEvents = True
End Sub

' Keypress Rules in Each TextBox Input
Private Sub NameInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 65 And KeyAscii <= 90) And Not (KeyAscii >= 97 And KeyAscii <= 122) And Not (KeyAscii = 32) And Not (KeyAscii = 39) And Not (KeyAscii >= 44 And KeyAscii <= 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub GenderInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 65 And KeyAscii <= 90) And Not (KeyAscii >= 97 And KeyAscii <= 122) And Not (KeyAscii = 45) Then
        KeyAscii = 0
    End If
End Sub
Private Sub CityOfBirthInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 65 And KeyAscii <= 90) And Not (KeyAscii >= 97 And KeyAscii <= 122) And Not (KeyAscii = 32) And Not (KeyAscii = 39) And Not (KeyAscii >= 44 And KeyAscii <= 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub BodyWeightInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
Private Sub BodyHeightInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

' Cancel Button
Private Sub CancelButton_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to cancel this test? All data you have filled in will be lost and cannot be recovered!", vbYesNo + vbCritical, "Confirmation")
    If response = vbYes Then
        Unload Me
    End If
End Sub

' Next Button
Private Sub NextButton_Click()
    Dim totalDb As Integer
    totalDb = ThisWorkbook.Worksheets("Tools").Range("totalDatabase")
    
    ' Name Input Check
    If NameInput = "" Then
        MsgBox "Name cannot be empty!", vbExclamation + vbOKOnly, "Warning"
    ' Gender Input Check
    ElseIf GenderInput <> "Male" And GenderInput <> "Female" Then
        MsgBox "Gender is invalid!", vbExclamation + vbOKOnly, "Warning"
    ' Date of Birth Check
    ElseIf CityOfBirthInput = "" Then
        MsgBox "Do not forget to fill your city of birth!", vbExclamation + vbOKOnly, "Warning"
    ' Body Weight Check
    ElseIf BodyWeightInput = 0 Then
        MsgBox "Body Weight cannot be 0 kg!", vbExclamation + vbOKOnly, "Warning"
    ' Body Height Check
    ElseIf BodyWeightInput = 0 Then
        MsgBox "Body Height cannot be 0 cm!", vbExclamation + vbOKOnly, "Warning"
    ' Execute if all data are already valid
    Else
        ThisWorkbook.Worksheets("Result").Range("NameOutput").Value = UCase(NameInput.Value)
        ThisWorkbook.Worksheets("Result").Range("GenderOutput").Value = GenderInput.Value
        ThisWorkbook.Worksheets("Result").Range("CityOfBirthOutput").Value = UCase(CityOfBirthInput.Value)
        ThisWorkbook.Worksheets("Result").Range("BodyWeightOutput").Value = BodyWeightInput.Value
        ThisWorkbook.Worksheets("Result").Range("BodyHeightOutput").Value = BodyHeightInput.Value

        Dim formattedDate As String
        formattedDate = Format(DateSerial(YearCombo.Value, MonthCombo.Value, DateCombo.Value), "DD/MM/YYYY")
        ThisWorkbook.Worksheets("Result").Range("DateOfBirthOutput").MergeArea.Value = formattedDate
        
        
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb + 1, 0) = UCase(NameInput.Value)
        
        If UCase(GenderInput.Value) = "MALE" Then
            ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb + 1, 1) = "M"
        ElseIf UCase(GenderInput.Value) = "FEMALE" Then
            ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb + 1, 1) = "F"
        Else
            ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb + 1, 1) = "Error"
        End If
        
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb + 1, 2) = BodyWeightInput.Value
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb + 1, 3) = BodyHeightInput.Value
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb + 1, 4) = formattedDate
        ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb + 1, 5) = UCase(CityOfBirthInput.Value)
        

        PFT_Form1.Hide
        PFT_Form2.Show
    End If
End Sub

' Hide Button
Private Sub HideButton_Click()
    PFT_Form1.Hide
End Sub

Private Sub UserForm_Initialize()
    BodyWeightInput.MaxLength = 3
    BodyHeightInput.MaxLength = 3
End Sub
