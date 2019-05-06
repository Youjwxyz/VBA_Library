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

Property Let workbookFullPath(ByVal fileFullPath As String)
If fileFullPath <> "" Then
    Dim i As Long
    i = InStrRev(fileFullPath, ".")
    If i > 0 And UCase(Mid(fileFullPath, i + 1, 3)) = "XLS" Then
        bookFullPath = fileFullPath
        Dim one As Excel.Workbook
        For Each one In Workbooks
            If UCase(one.FullName) = UCase(bookFullPath) Then
                Set bookObject = one
                Exit For
            End If
        Next one
        Set one = Nothing
        If TypeName(bookObject) = "Workbook" Then
            If UCase(bookObject.FullName) <> UCase(bookFullPath) Then
                Call resetClass
            End If
        Else
            If Dir(bookFullPath) <> "" Then
                If password = "" Then
                    Set bookObject = Workbooks.Open(bookFullPath, updateLink, readOnly, , , , ignoreReadOnly)
                Else
                    Set bookObject = Workbooks.Open(bookFullPath, updateLink, readOnly, , password, , ignoreReadOnly)
                End If
            Else
                Set bookObject = Workbooks.Add
                Application.DisplayAlerts = False
                bookObject.SaveAs bookFullPath
                Application.DisplayAlerts = True
            End If
        End If
    Else
        Call resetClass
    End If
Else
    Call resetClass
End If
End Property

Property Set workbookObject(ByRef sourceBook As Excel.Workbook)
If TypeName(sourceBook) = "Workbook" Then
    Set bookObject = sourceBook
    bookFullPath = bookObject.FullName
Else
    Call resetClass
End If
End Property

Property Get workbookObject() As Excel.Workbook
If TypeName(bookObject) = "Workbook" Then
    Set workbookObject = bookObject
Else
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
