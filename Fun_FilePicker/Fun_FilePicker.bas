Attribute VB_Name = "Fun_FilePicker"
Option Explicit

'pickToHaveFile
'pickToHaveFolder
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
If FP.filePath <> "" Then
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
        pickToHaveFile = ""
    End Select
    Set SP = Nothing
Else
    pickToHaveFile = ""
End If
Set FP = Nothing
End Function

Public Function pickToHaveFolder() As String
Dim FP As New FolderPicker
FP.pickFolder
If FP.folderPath <> "" Then
    pickToHaveFolder = FP.folderPath
Else
    pickToHaveFolder = ""
End If
Set FP = Nothing
End Function

Public Function pickFolderToHaveFiles(ByVal fileExtension As String, _
    Optional ByVal fileProperty As String = "FileFullName") As Variant
Dim FP As New FolderPicker
FP.pickFolder
If FP.folderPath <> "" Then
    Dim FS As New FilesSearcher
    FS.folder = FP.folderPath
    FS.extension = fileExtension
    FS.searchFile
    If TypeName(FS.fileList) = "String()" And nonEmptyOneDArray(FS.fileList) = True Then
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
                    Temp(j) = ""
                    Exit For
                End Select
            Next j
            pickFolderToHaveFiles = Temp
            Erase Temp
        Else
            pickFolderToHaveFiles = Empty
        End If
        Set SP = Nothing
    Else
        pickFolderToHaveFiles = Empty
    End If
    Set FS = Nothing
Else
    pickFolderToHaveFiles = Empty
End If
Set FP = Nothing
End Function

Private Function nonEmptyOneDArray(ByRef sourceOneDArray As Variant) As Boolean
On Error Resume Next
If UBound(sourceOneDArray) < LBound(sourceOneDArray) Then
    nonEmptyOneDArray = False
Else
    nonEmptyOneDArray = True
End If
End Function
