Option Explicit

'Below varaibles are used in development phase, using Word ojbect reference.
'Private WordRange As Word.Range
Private WordRange As Object
Private LineText As String
Private ListItem As Boolean
Private ErrMessage As String


Private Sub Class_Initialize()

Call RefreshData

End Sub

Private Sub RefreshData()

If WordRange Is Nothing Then
    LineText = ""
    ListItem = False
    ErrMessage = "Current word line is not defined yet"
Else
    LineText = WordRange.Text
    ErrMessage = ""
    If WordRange.ListParagraphs.Count > 0 Then
        ListItem = True
    Else
        ListItem = False
    End If
End If

End Sub

Public Property Set CurrentLine(ByRef SourceWordRange As Object)

SourceWordRange.Move wdParagraph, -1
SourceWordRange.MoveEnd wdParagraph, 1
Set WordRange = SourceWordRange
Call RefreshData

End Property

Public Sub LastLine(Optional ByVal SkipTable As Boolean)

If Not WordRange Is Nothing Then
    WordRange.Move wdParagraph, -2
    WordRange.MoveEnd wdParagraph, 1
    If SkipTable Then
        Do Until WordRange.Tables.Count = 0
            WordRange.Move wdTable, -1
            WordRange.Move wdParagraph, -1
            WordRange.MoveEnd wdParagraph, 1
        Loop
    End If
End If
Call RefreshData

End Sub

Public Property Get Text() As String

If WordRange Is Nothing Then
    Err.Raise 1, , ErrMessage
Else
    Text = LineText
End If

End Property

Public Property Get IsList() As Boolean

If WordRange Is Nothing Then
    Err.Raise 1, , ErrMessage
Else
    IsList = ListItem
End If

End Property
