Attribute VB_Name = "ComponentExporter"
Sub ExportAllVBAComponents()
    Dim vbComp As Object
    Dim exportPath As String

    ' Where you want to save exported files
    exportPath = ThisWorkbook.Path & "\exported_components\" ' You can change the folder path based on your needs

    ' Create the folder if it doesn't exist
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    ' Loop through each component
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case 1, 2, 3 ' 1=Module, 2=Class module, 3=UserForm
                vbComp.Export exportPath & vbComp.Name & GetProperExtension(vbComp.Type)
        End Select
    Next vbComp

    MsgBox "Export completed!" & vbCrLf & "Files are saved in: " & exportPath, vbInformation
End Sub

Private Function GetProperExtension(compType As Long) As String
    Select Case compType
        Case 1 ' Module
            GetProperExtension = ".bas"
        Case 2 ' Class Module
            GetProperExtension = ".cls"
        Case 3 ' UserForm
            GetProperExtension = ".frm"
        Case Else
            GetProperExtension = ""
    End Select
End Function


