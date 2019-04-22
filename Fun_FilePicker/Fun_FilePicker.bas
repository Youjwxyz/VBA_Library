Attribute VB_Name = "Fun_FilePicker"
Option Explicit

'pickToHaveFileName
'pickToHaveFileFullName
'pickToHaveFileDirectory
'pickFolderToHaveFileNames
'pickFolderToHaveFileFullNames
'pickFolderToHaveFileDirectories

Public Sub importFunctionRelatedClass()
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

Public Function pickToHaveFileName() As String
pickToHaveFileName = pickToHaveFile("FileName")
End Function

Public Function pickToHaveFileFullName() As String
pickToHaveFileFullName = pickToHaveFile("FileFullName")
End Function

Public Function pickToHaveFileDirectory() As String
pickToHaveFileDirectory = pickToHaveFile("FileDirectory")
End Function

Private Function pickToHaveFile(ByVal fileProperty As String) As String
Dim FP As New FilePicker
FP.pickFile
Dim SP As New StrPather
SP.fileFullPath = FP.file
Select Case fileProperty
Case "FileName"
    pickToHaveFile = SP.fileNameWithExt
Case "FileFullName"
    pickToHaveFile = SP.fileFullPath
Case "FileDicrectory"
    pickToHaveFile = SP.path
End Select
Set SP = Nothing
Set FP = Nothing
End Function

Public Function pickFolderToHaveFileNames(ByVal fileExtension As String) As Variant
pickFolderToHaveFileNames = pickFolderToHaveFiles(fileExtension, "FileName")
End Function

Public Function pickFolderToHaveFileFullNames(ByVal fileExtension As String) As Variant
pickFolderToHaveFileFullNames = pickFolderToHaveFiles(fileExtension, "FileFullName")
End Function

Public Function pickFolderToHaveFileDirectories(ByVal fileExtension As String) As Variant
pickFolderToHaveFileDirectories = pickFolderToHaveFiles(fileExtension, "FileDicrectory")
End Function

Private Function pickFolderToHaveFiles(ByVal fileExtension As String, _
    ByVal fileProperty As String) As Variant
Dim FS As New FilesSearcher
Dim FP As New FolderPicker
FP.pickFolder
FS.folder = FP.folder
FS.extension = fileExtension
FS.searchFile
Dim i As Long
Dim j As Long
Dim SP As New StrPather
Dim Temp() As String
i = UBound(FS.files)
If i > -1 Then
    ReDim Temp(i) As String
    For j = 0 To i
        SP.fileFullPath = FS.files(j)
        Select Case fileProperty
        Case "FileName"
            Temp(j) = SP.fileNameWithExt
        Case "FileFullName"
            Temp(j) = SP.fileFullPath
        Case "FileDicrectory"
            Temp(j) = SP.path
        End Select
    Next j
End If
pickFolderToHaveFiles = Temp
Erase Temp
Set SP = Nothing
Set FS = Nothing
Set FP = Nothing
End Function

