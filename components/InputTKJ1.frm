VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} InputTKJ1 
   Caption         =   "Input Tes Kebugaran Jasmani (1 of 3)"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10155
   OleObjectBlob   =   "InputTKJ1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "InputTKJ1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Combo Tanggal Dinamis
Private Sub ComboTanggal_Change()
    Application.EnableEvents = False
    
    If Me.ComboBulan.Value = 4 Or Me.ComboBulan.Value = 6 Or Me.ComboBulan.Value = 9 And Me.ComboBulan.Value = 11 Then
        ComboTanggal.RowSource = "Tools!C3:C32"
    ElseIf Me.ComboBulan.Value = 2 Then
        If Me.ComboTahun.Value Mod 4 = 0 Then
            ComboTanggal.RowSource = "Tools!C3:C31"
        Else
            ComboTanggal.RowSource = "Tools!C3:C30"
        End If
    Else
        ComboTanggal.RowSource = "Tools!C3:C33"
    End If
    
    Application.EnableEvents = True
End Sub
'Aturan KeyPress Tiap TextBox Input
Private Sub InputNama_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 65 And KeyAscii <= 90) And Not (KeyAscii >= 97 And KeyAscii <= 122) And Not (KeyAscii = 32) And Not (KeyAscii = 39) And Not (KeyAscii >= 44 And KeyAscii <= 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub InputJenisKelamin_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 65 And KeyAscii <= 90) And Not (KeyAscii >= 97 And KeyAscii <= 122) And Not (KeyAscii = 45) Then
        KeyAscii = 0
    End If
End Sub
Private Sub InputKotaKelahiran_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 65 And KeyAscii <= 90) And Not (KeyAscii >= 97 And KeyAscii <= 122) And Not (KeyAscii = 32) And Not (KeyAscii = 39) And Not (KeyAscii >= 44 And KeyAscii <= 46) Then
        KeyAscii = 0
    End If
End Sub
Private Sub InputBeratBadan_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
End Sub
Private Sub InputTinggiBadan_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
        KeyAscii = 0
    End If
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
    Dim totalDb As Integer
    totalDb = ThisWorkbook.Worksheets("Tools").Range("totalDatabase")
    
    'Cek Input Nama
    If InputNama = "" Then
        MsgBox "Input nama tidak boleh kosong!", vbExclamation + vbOKOnly, "Peringatan"
    'Cek Input Jenis Kelamin
    ElseIf InputJenisKelamin <> "Laki-laki" And InputJenisKelamin <> "Perempuan" Then
        MsgBox "Input Jenis Kelamin tidak valid!", vbExclamation + vbOKOnly, "Peringatan"
    'Cek Input Tempat Lahir
    ElseIf InputKotaKelahiran = "" Then
        MsgBox "Jangan lupa isi tempat lahir anda!", vbExclamation + vbOKOnly, "Peringatan"
    'Cek Input BB
    ElseIf InputBeratBadan = 0 Then
        MsgBox "Berat Badan tidak boleh 0 kg!", vbExclamation + vbOKOnly, "Peringatan"
    'Cek Input TB
    ElseIf InputBeratBadan = 0 Then
        MsgBox "Tinggi Badan tidak boleh 0 cm!", vbExclamation + vbOKOnly, "Peringatan"
    'Eksekusi jika semuanya sudah valid
    Else
        ThisWorkbook.Worksheets("Hasil").Range("OutputNama").Value = UCase(InputNama.Value)
        ThisWorkbook.Worksheets("Hasil").Range("OutputJenisKelamin").Value = InputJenisKelamin.Value
        ThisWorkbook.Worksheets("Hasil").Range("OutputKotaKelahiran").Value = UCase(InputKotaKelahiran.Value)
        ThisWorkbook.Worksheets("Hasil").Range("OutputAngkaBeratBadan").Value = InputBeratBadan.Value
        ThisWorkbook.Worksheets("Hasil").Range("OutputAngkaTinggiBadan").Value = InputTinggiBadan.Value

        Dim formattedDate As String
        formattedDate = Format(DateSerial(ComboTahun.Value, ComboBulan.Value, ComboTanggal.Value), "DD/MM/YYYY")
        ThisWorkbook.Worksheets("Hasil").Range("OutputTanggalLahir").MergeArea.Value = formattedDate
        
        
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb + 1, 0) = UCase(InputNama.Value)
        
        If UCase(InputJenisKelamin.Value) = "LAKI-LAKI" Then
            ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb + 1, 1) = "L"
        ElseIf UCase(InputJenisKelamin.Value) = "PEREMPUAN" Then
            ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb + 1, 1) = "P"
        Else
            ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb + 1, 1) = "Err"
        End If
        
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb + 1, 2) = InputBeratBadan.Value
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb + 1, 3) = InputTinggiBadan.Value
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb + 1, 4) = formattedDate
        ThisWorkbook.Worksheets("Database").Range("kolomNama").Offset(totalDb + 1, 5) = UCase(InputKotaKelahiran.Value)
        

        InputTKJ1.Hide
        InputTKJ2.Show
    End If
End Sub
'Tombol Sembunyikan
Private Sub TombolSembunyikan_Click()
    InputTKJ1.Hide
End Sub
Private Sub UserForm_Initialize()
    InputBeratBadan.MaxLength = 3
    InputTinggiBadan.MaxLength = 3
End Sub
