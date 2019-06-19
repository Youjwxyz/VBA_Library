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

Private folder As String

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

Property Get folderPath() As String
If folder <> "" Then
    folderPath = folder
Else
    Call resetClass("No folder is picked.")
    folderPath = folder
End If
End Property

Private Sub resetClass(Optional ByVal errorMessage As String = "Unknown Error")
folder = ""
MsgBox errorMessage
End Sub
