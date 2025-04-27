Attribute VB_Name = "ComponentImporter"
Sub ImportAllVBAComponents()
    Dim importPath As String
    Dim fileName As String

    ' Folder where your exported modules are
    importPath = ThisWorkbook.Path & "\components\"
    
    If Dir(importPath, vbDirectory) = "" Then
        MsgBox "The folder path is not found!", vbExclamation
        Exit Sub
    End If

    ' Import each file
    fileName = Dir(importPath & "*.*")
    Do While fileName <> ""
        If (LCase(Right(fileName, 4)) = ".bas") Or _
           (LCase(Right(fileName, 4)) = ".cls") Or _
           (LCase(Right(fileName, 4)) = ".frm") Then
           
           ThisWorkbook.VBProject.VBComponents.Import importPath & fileName
        End If
        fileName = Dir
    Loop

    MsgBox "Import completed!" & vbCrLf & "Files imported from: " & importPath, vbInformation
End Sub



