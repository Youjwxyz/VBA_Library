
Option Explicit

Public Function FilesFromPicker(Optional ByVal TitleStr As String, _
    Optional ByVal MultiSelection As Boolean) As Variant
'Select single or multile files in pop up dialog.
'Related file(s) data will be stored in a 1-demention array.
'If no file is selected, array(0) = "" will be returned.
Dim TempStrs() As String
Dim i, j As Long
Dim FSO As FileDialog
Set FSO = Application.FileDialog(msoFileDialogFilePicker)
FSO.Title = TitleStr
If MultiSelection = True Then
    FSO.AllowMultiSelect = True
Else
    FSO.AllowMultiSelect = False
End If
If FSO.Show = -1 Then
    If MultiSelection = True Then
        i = FSO.SelectedItems.Count
        ReDim TempStrs(i - 1) As String
        For j = 1 To i
            TempStrs(j - 1) = FSO.SelectedItems(j)
        Next j
    Else
        ReDim TempStrs(0) As String
        TempStrs(0) = FSO.SelectedItems(1)
    End If
Else
    ReDim TempStrs(0) As String
    TempStrs(0) = ""
End If
Set FSO = Nothing
FilesFromPicker = TempStrs
Erase TempStrs
End Function

Public Function FilesInFolder(Optional ByVal TitleStr As String, _
    Optional ByVal Extension As String) As Variant
'Select a folder in pop up dialog, and look for files with specified extension.
'Related file(s) data will be stored in a 1-demention array.
'If no file is selected, array(0) = "" will be returned.
Dim TempStrs() As String
Dim FileStr As String
Dim i As Long
Dim FSO As FileDialog
Set FSO = Application.FileDialog(msoFileDialogFolderPicker)
FSO.Title = TitleStr
If FSO.Show = -1 Then
    FileStr = Dir(FSO.SelectedItems(1) & "\*." & Extension & "*", vbNormal)
    Do Until FileStr = ""
        i = i + 1
        ReDim Preserve TempStrs(i - 1) As String
        TempStrs(i - 1) = FileStr
        FileStr = Dir
    Loop
Else
    ReDim TempStrs(0) As String
    TempStrs(0) = ""
End If
Set FSO = Nothing
FilesInFolder = TempStrs
Erase TempStrs
End Function
