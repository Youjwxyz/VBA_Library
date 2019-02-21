
Option Explicit

'Below varaibles are used in development phase, using Word ojbect reference.
'Private WordApp As Word.Application
'Private WordDoc As Word.Document
'Private TempDoc As Word.Document
'Private TempApp As Word.Application

Private WordApp As Object
Private WordDoc As Object
Private TempDoc As Object
Private TempApp As Object
Private ErrMessage As String
Private WordName As String
Private WordFullPath As String


Private Sub Class_Initialize()

Call AttainWordApp

End Sub

Private Sub AttainWordApp()

On Error Resume Next
Set TempApp = GetObject(, "Word.Application")
If Err.Number <> 0 Then
    Err.Clear
    Set TempApp = CreateObject("Word.Application")
    If Err.Number <> 0 Then
        Err.Clear
        Set TempApp = Nothing
        ErrMessage = "Word application is not installed"
    End If
End If
If Not TempApp Is Nothing Then
    Set WordApp = TempApp
    WordApp.Visible = True
End If

End Sub

Private Sub AttainWordDoc()

Set WordDoc = Nothing
For Each TempDoc In WordApp.Documents
    If TempDoc.Name = WordName Then
        Set WordDoc = TempDoc
        Exit For
    End If
Next TempDoc

End Sub


Public Property Get App() As Object

If ErrMessage = "" Then
    Set App = WordApp
Else
    Err.Raise 1, , ErrMessage
End If

End Property

Public Property Get Doc() As Object

If ErrMessage = "" Then
    Set Doc = WordDoc
Else
    Err.Raise 1, , ErrMessage
End If

End Property

Public Property Let Name(ByVal FileName As String)

WordName = FileName
Call AttainWordDoc
If WordDoc Is Nothing Then
    ErrMessage = "Word document is not found with provided name"
    WordFullPath = ""
Else
    WordFullPath = WordDoc.Path & "\" & WordName
End If

End Property

Public Property Get Name() As String

Name = WordName

End Property






Public Property Let FullPath(ByVal FileFullPath As String)

On Error Resume Next
WordFullPath = FileFullPath
WordName = Mid(WordFullPath, InStrRev(WordFullPath, "\") + 1)
Call AttainWordDoc
If WordDoc Is Nothing Then
    WordApp.Documents.Open WordFullPath
    If Err.Number <> 0 Then
        Err.Clear
        ErrMessage = "Word document fail to open with provided file full path"
    Else
        Call AttainWordDoc
    End If
Else
    If WordDoc.Path & "\" & WordName <> WordFullPath Then
        Set WordDoc = Nothing
        ErrMessage = "Word document could not be open because of name duplication"
    End If
End If

End Property

Public Property Get FullPath() As String

FullPath = WordFullPath

End Property







