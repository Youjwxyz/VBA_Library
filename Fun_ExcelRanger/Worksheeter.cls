VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Worksheeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public bookObject As Excel.Workbook
Public removeExistingSheet As Boolean
Public createIfNotFound As Boolean
Private sheetName As String
Private sheetObject As Excel.Worksheet

Property Set worksheetObject(ByRef sourceSheet As Excel.Worksheet)
If TypeName(sourceSheet) = "Worksheet" Then
    Set sheetObject = sourceSheet
    Call synchronizeClass
Else
    Call resetClass
End If
End Property

Property Let worksheetName(ByVal nameOfSheet As String)
If TypeName(bookObject) = "Workbook" And nameOfSheet <> "" Then
    sheetName = nameOfSheet
    Set sheetObject = Nothing
    Dim one As Excel.Worksheet
    For Each one In bookObject.Worksheets
        If UCase(one.Name) = UCase(sheetName) Then
            If removeExistingSheet = True Then
                Application.DisplayAlerts = False
                    one.Delete
                    Set one = bookObject.Worksheets.Add
                    one.Name = sheetName
                Application.DisplayAlerts = True
            End If
            Set sheetObject = one
            Call synchronizeClass
            Exit For
        End If
    Next one
    If TypeName(sheetObject) <> "Worksheet" Then
        If createIfNotFound = True Then
            Set one = bookObject.Worksheets.Add
            one.Name = sheetName
            Set sheetObject = one
            Call synchronizeClass
        Else
            Call resetClass
        End If
    End If
Else
    Call resetClass
End If
End Property

Property Get worksheetObject() As Excel.Worksheet
If TypeName(sheetObject) = "Worksheet" Then
    Set worksheetObject = sheetObject
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Private Sub resetClass()
Set bookObject = Nothing
removeExistingSheet = False
createIfNotFound = True
sheetName = ""
Set sheetObject = Nothing
End Sub

Private Sub synchronizeClass()
If TypeName(sheetObject) = "Worksheet" Then
    Set bookObject = sheetObject.Parent
    sheetName = sheetObject.Name
End If
End Sub
