VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FolderPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public folder As String

Public Sub pickFolder()
Dim FSO As FileDialog
Set FSO = Application.FileDialog(msoFileDialogFolderPicker)
If FSO.Show = -1 Then
    folder = FSO.SelectedItems(1)
Else
    Call resetClass
End If
Set FSO = Nothing
End Sub

Private Sub resetClass()
folder = ""
End Sub