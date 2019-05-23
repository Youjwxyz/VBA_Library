VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private file As String

Public Sub pickFile()
Dim FSO As FileDialog
Set FSO = Application.FileDialog(msoFileDialogFilePicker)
FSO.AllowMultiSelect = False
If FSO.Show = -1 Then
    file = FSO.SelectedItems(1)
Else
    Call resetClass
End If
Set FSO = Nothing
End Sub

Property Get filePath() As String
If file <> "" Then
    filePath = file
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Private Sub resetClass()
file = ""
End Sub
