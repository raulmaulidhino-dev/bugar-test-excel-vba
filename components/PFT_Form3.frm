VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PFT_Form3 
   Caption         =   "Physical Fitness Test Form (3 of 3)"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10155
   OleObjectBlob   =   "PFT_Form3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PFT_Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Keypress Rules in Each TextBox Input
Private Sub HorizontalJumpInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub SitNReachFlexibilityInput_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub

' Previous Button
Private Sub PreviousButton_Click()
    PFT_Form3.Hide
    PFT_Form2.Show
End Sub

' Cancel Button
Private Sub CancelButton_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Are you sure you want to cancel this test? All data you have filled in will be lost and cannot be recovered!", vbYesNo + vbCritical, "Confirmation")
    If response = vbYes Then
        Unload Me
    End If
End Sub

' Submit Button
Private Sub SubmitButton_Click()
    Dim totalDb As Integer
    totalDb = ThisWorkbook.Worksheets("Tools").Range("totalDatabase")
    If HorizontalJumpInput.Value = "" Then
        MsgBox ("There are still empty input(s)!")
    ElseIf SitNReachFlexibilityInput.Value = "" Then
        MsgBox ("There are still empty input(s)!")
    Else
        Dim response As VbMsgBoxResult
        response = MsgBox("Are you sure you have entered the data correctly and validly? You can tap recheck if you feel there is any incorrect or invalid data", vbYesNo + vbQuestion, "Reconfirm")
        If response = vbYes Then
            ThisWorkbook.Worksheets("Result").Range("HorizontalJumpNumberOutput") = HorizontalJumpInput.Value
            ThisWorkbook.Worksheets("Result").Range("SitNReachFlexibilityNumberOutput") = SitNReachFlexibilityInput.Value
            ThisWorkbook.Worksheets("Result").Activate
            ThisWorkbook.Worksheets("Result").Range("AZ1").Select
            
            ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb, 25) = HorizontalJumpInput.Value
            ThisWorkbook.Worksheets("Database").Range("nameColumn").Offset(totalDb, 28) = SitNReachFlexibilityInput.Value
            
            
            Unload Me
            Unload PFT_Form2
            Unload PFT_Form1
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    HorizontalJumpInput.MaxLength = 4
    SitNReachFlexibilityInput.MaxLength = 3
End Sub
