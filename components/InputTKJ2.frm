VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputTKJ2 
   Caption         =   "Input Kebugaran Tes Jasmani (2 of 3)"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10155
   OleObjectBlob   =   "InputTKJ2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputTKJ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------
'Aturan Keypress untuk tiap TextBox Input
Private Sub ComboMenitLari1200_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
Private Sub ComboDetikLari1200_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
Private Sub InputLari60_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub InputSitUp_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
Private Sub InputHandEyeCoor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
Private Sub InputHexAgil_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub InputStorkBalance_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
'------------------------------------------------------------------------------------
'Tombol Sebelumnya
Private Sub TombolSebelumnya_Click()
    InputTKJ2.Hide
    InputTKJ1.Show
End Sub
'Tombol Batal
Private Sub TombolBatal_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Apakah anda yakin ingin membatalkan tes ini? Semua data yang telah anda isi akan hilang dan tidak dapat dipulihkan!", vbYesNo + vbCritical, "Konfirmasi")
    If response = vbYes Then
        Unload Me
    End If
End Sub
'Tombol Lanjut
Private Sub TombolLanjut_Click()
    Dim WaktuLari1200 As Integer
    Dim totalDb As Integer
    totalDb = ThisWorkbook.Worksheets("Tools").Range("totalDatabase")
    
    If ComboMenitLari1200.Value = "" Then
        MsgBox "Masih ada input yang kosong", vbExclamation + vbOKOnly
    ElseIf ComboDetikLari1200.Value = "" Then
        MsgBox "Masih ada input yang kosong", vbExclamation + vbOKOnly
    ElseIf InputLari60.Value = "" Then
        MsgBox "Masih ada input yang kosong", vbExclamation + vbOKOnly
    ElseIf InputSitUp.Value = "" Then
        MsgBox "Masih ada input yang kosong", vbExclamation + vbOKOnly
    ElseIf InputHandEyeCoor.Value = "" Then
        MsgBox "Masih ada input yang kosong", vbExclamation + vbOKOnly
    ElseIf InputHexAgil.Value = "" Then
        MsgBox "Masih ada input yang kosong", vbExclamation + vbOKOnly
    ElseIf InputStorkBalance.Value = "" Then
        MsgBox "Masih ada input yang kosong", vbExclamation + vbOKOnly
    Else
        Dim nilaiWaktuLari1200 As Date
        Dim waktuTerformat As String
        
        nilaiWaktuLari1200 = TimeSerial(0, ComboMenitLari1200.Value, ComboDetikLari1200.Value)
        waktuTerformat = Format(nilaiWaktuLari1200, "hh:mm:ss")
        
        ThisWorkbook.Worksheets("Hasil").Range("OutputAngkaLari1200") = waktuTerformat
        
        ThisWorkbook.Worksheets("Hasil").Range("OutputAngkaLari60") = Round(InputLari60.Value, 1)
        ThisWorkbook.Worksheets("Hasil").Range("OutputAngkaSitUp") = InputSitUp.Value
        ThisWorkbook.Worksheets("Hasil").Range("OutputAngkaHandEyeCoor") = InputHandEyeCoor.Value
        ThisWorkbook.Worksheets("Hasil").Range("OutputAngkaHexAgil") = Round(InputHexAgil.Value, 1)
        ThisWorkbook.Worksheets("Hasil").Range("OutputAngkaStorkBalance") = InputStorkBalance.Value
        
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb, 7) = waktuTerformat
        
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb, 10) = Round(InputLari60.Value, 1)
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb, 16) = InputSitUp.Value
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb, 22) = InputHandEyeCoor.Value
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb, 13) = Round(InputHexAgil.Value, 1)
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb, 19) = InputStorkBalance.Value
        
        InputTKJ2.Hide
        InputTKJ3.Show
    End If
End Sub
Private Sub UserForm_Initialize()
    InputLari60.MaxLength = 4
    InputSitUp.MaxLength = 2
    InputHandEyeCoor.MaxLength = 2
    InputHexAgil.MaxLength = 4
    InputStorkBalance.MaxLength = 2
End Sub
