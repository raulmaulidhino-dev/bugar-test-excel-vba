VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputTKJ3 
   Caption         =   "Input Tes Kebugaran (3 of 3)"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10155
   OleObjectBlob   =   "InputTKJ3.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputTKJ3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Aturan Keypress untuk tiap TextBox Input
Private Sub InputHJump_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) And Not (KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub InputSitNReach_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
'Tombol Sebelumnya
Private Sub TombolSebelumnya_Click()
    InputTKJ3.Hide
    InputTKJ2.Show
End Sub
'Tombol Batal
Private Sub TombolBatal_Click()
    Dim response As VbMsgBoxResult
    response = MsgBox("Apakah anda yakin ingin membatalkan tes ini? Semua data yang telah anda isi akan hilang dan tidak dapat dipulihkan!", vbYesNo + vbCritical, "Konfirmasi")
    If response = vbYes Then
        Unload Me
    End If
End Sub
'Tombol Selesai
Private Sub TombolSelesai_Click()
    Dim totalDb As Integer
    totalDb = ThisWorkbook.Worksheets("Tools").Range("totalDatabase")
    If InputHJump.Value = "" Then
        MsgBox ("Masih ada input yang kosong")
    ElseIf InputSitNReach.Value = "" Then
        MsgBox ("Masih ada input yang kosong")
    Else
        Dim response As VbMsgBoxResult
        response = MsgBox("Apakah anda yakin sudah mengisi data dengan benar dan valid? Anda bisa mengecek kembali bila merasa ada data input yang salah/tidak valid", vbYesNo + vbQuestion, "Konfirmasi Ulang")
        If response = vbYes Then
            ThisWorkbook.Worksheets("Hasil").Range("OutputAngkaHorizontalJump") = InputHJump.Value
            ThisWorkbook.Worksheets("Hasil").Range("OutputAngkaSitNReach") = InputSitNReach.Value
            ThisWorkbook.Worksheets("Hasil").Activate
            ThisWorkbook.Worksheets("Hasil").Range("AZ1").Select
            
            ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb, 25) = InputHJump.Value
            ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb, 28) = InputSitNReach.Value
            
            
            Unload Me
            Unload InputTKJ2
            Unload InputTKJ1
        End If
    End If
End Sub
Private Sub UserForm_Initialize()
    InputHJump.MaxLength = 4
    InputSitNReach.MaxLength = 3
End Sub
