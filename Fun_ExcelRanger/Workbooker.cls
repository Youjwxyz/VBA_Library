VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Workbooker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public updateLink As Boolean
Public readOnly As Boolean
Public password As String
Public ignoreReadOnly As Boolean
Private bookFullPath As String
Private bookObject As Excel.Workbook

Property Set workbookObject(ByRef sourceBook As Excel.Workbook)
If TypeName(sourceBook) = "Workbook" Then
    Set bookObject = sourceBook
    Call synchronizeClass
Else
    Call resetClass
End If
End Property

Property Let workbookFullPath(ByVal fileFullPath As String)
fileFullPath = Trim(fileFullPath)
If verifyBookFullPath(fileFullPath) = True Then
    bookFullPath = fileFullPath
    Set bookObject = Nothing
    Dim one As Excel.Workbook
    For Each one In Workbooks
        If UCase(one.FullName) = UCase(bookFullPath) Then
            Set bookObject = one
            Call synchronizeClass
            Exit For
        End If
    Next one
    If TypeName(bookObject) <> "Workbook" Then
        If Dir(bookFullPath) <> "" Then
            If password = "" Then
                Set one = Workbooks.Open(bookFullPath, updateLink, readOnly, , , , ignoreReadOnly)
            Else
                Set one = Workbooks.Open(bookFullPath, updateLink, readOnly, , password, , ignoreReadOnly)
            End If
        Else
            Set one = Workbooks.Add
            Application.DisplayAlerts = False
            one.SaveAs bookFullPath
            Application.DisplayAlerts = True
        End If
        Set bookObject = one
        Call synchronizeClass
    End If
Else
    Call resetClass
End If
End Property

Property Get workbookObject() As Excel.Workbook
If TypeName(bookObject) = "Workbook" Then
    Set workbookObject = bookObject
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Private Sub resetClass()
updateLink = False
readOnly = False
password = ""
ignoreReadOnly = False
bookFullPath = ""
Set bookObject = Nothing
End Sub

Private Sub synchronizeClass()
If TypeName(bookObject) = "Workbook" Then
    bookFullPath = bookObject.FullName
End If
End Sub

Private Function verifyBookFullPath(ByVal fileFullPath As String) As Boolean
fileFullPath = Trim(fileFullPath)
If fileFullPath <> "" Then
    Dim i As Long
    i = InStrRev(fileFullPath, ".")
    If i > 0 And UCase(Mid(fileFullPath, i + 1, 3)) = "XLS" Then
        verifyBookFullPath = True
    End If
End If
End Function
