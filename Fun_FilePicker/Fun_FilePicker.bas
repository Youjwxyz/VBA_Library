Attribute VB_Name = "Fun_FilePicker"
Option Explicit

'pickToHaveFile
'pickFolderToHaveFiles

Private Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("FilesSearcher", "FolderPicker", "StrPather", "FilePicker")
Dim classFolder As String
classFolder = "C:\Business\Macros\VBA_Class"
Dim one As Variant
Dim com As Variant
Dim comFound As Boolean
For Each one In classArr
    comFound = False
    For Each com In ThisWorkbook.VBProject.VBComponents
        If com.Name = one Then
            comFound = True
            Exit For
        End If
    Next com
    If Not comFound Then
        ThisWorkbook.VBProject.VBComponents.Import classFolder & "\" & one & ".cls"
    End If
Next one
End Sub

Public Function pickToHaveFile(Optional ByVal fileProperty As String = "FileFullName") As String
Dim FP As New FilePicker
FP.pickFile
Dim SP As New StrPather
SP.fileFullPath = FP.filePath
Select Case fileProperty
Case "FileName"
    pickToHaveFile = SP.fileNameWithExt
Case "FileFullName"
    pickToHaveFile = SP.fileFullPath
Case "FileDirectory"
    pickToHaveFile = SP.path
Case "FileNameWithoutExtension"
    pickToHaveFile = SP.fileNameWithoutExt
Case "FileExtension"
    pickToHaveFile = SP.extensionName
Case Else
    MsgBox "Incorrect fileProperty which should be one of these: " & Chr(10) & _
        "- FileName" & Chr(10) & _
        "- FileFullName" & Chr(10) & _
        "- FileDirectory" & Chr(10) & _
        "- FileNameWithoutExtension" & Chr(10) & _
        "- FileExtension"
End Select
Set SP = Nothing
Set FP = Nothing
End Function

Public Function pickFolderToHaveFiles(ByVal fileExtension As String, _
    Optional ByVal fileProperty As String = "FileFullName") As Variant
Dim FS As New FilesSearcher
Dim FP As New FolderPicker
FP.pickFolder
FS.folder = FP.folderPath
FS.extension = fileExtension
FS.searchFile
Dim i As Long
Dim j As Long
Dim SP As New StrPather
Dim Temp() As String
i = UBound(FS.fileList)
If i > -1 Then
    ReDim Temp(i) As String
    For j = 0 To i
        SP.fileFullPath = FS.fileList(j)
        Select Case fileProperty
        Case "FileName"
            Temp(j) = SP.fileNameWithExt
        Case "FileFullName"
            Temp(j) = SP.fileFullPath
        Case "FileDirectory"
            Temp(j) = SP.path
        Case "FileNameWithoutExtension"
            Temp(j) = SP.fileNameWithoutExt
        Case "FileExtension"
            Temp(j) = SP.extensionName
        Case Else
            MsgBox "Incorrect fileProperty which should be one of these: " & Chr(10) & _
                "- FileName" & Chr(10) & _
                "- FileFullName" & Chr(10) & _
                "- FileDirectory" & Chr(10) & _
                "- FileNameWithoutExtension" & Chr(10) & _
                "- FileExtension"
            Exit For
        End Select
    Next j
End If
pickFolderToHaveFiles = Temp
Erase Temp
Set SP = Nothing
Set FS = Nothing
Set FP = Nothing
End Function

