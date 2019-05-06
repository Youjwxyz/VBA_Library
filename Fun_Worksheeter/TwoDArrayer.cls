VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TwoDArrayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sourceData As Variant
Public parentWorkbook As Excel.Workbook
Public worksheetName As String

Public Sub toWorksheet()
If isTwoDArray(sourceData) And TypeName(parentWorkbook) = "Workbook" And worksheetName <> "" Then
    Dim i As Long
    Dim j As Long
    i = UBound(sourceData, 1)
    j = UBound(sourceData, 2)
    Call resetWorksheet
    With parentWorkbook.Worksheets(worksheetName)
        .Range(.Cells(1, 1), .Cells(i, j)).Value = sourceData
    End With
Else
    Call resetClass
End If
End Sub

Public Sub fromWorksheet()
If TypeName(parentWorkbook) = "Workbook" And worksheetName <> "" Then
    Dim i As Long
    Dim j As Long
    With parentWorkbook.Worksheets(worksheetName)
        i = .UsedRange.SpecialCells(xlCellTypeLastCell).Row
        j = .UsedRange.SpecialCells(xlCellTypeLastCell).Column
        If i = 1 And j = 1 Then
            Dim temp(1 To 1, 1 To 1) As Variant
            temp(1, 1) = .Cells(1, 1).Value
            sourceData = temp
            Erase temp
        Else
            sourceData = .Range(.Cells(1, 1), .Cells(i, j)).Value
        End If
    End With
Else
    Call resetClass
End If
End Sub

Private Function resetWorksheet()
If TypeName(parentWorkbook) = "Workbook" And worksheetName <> "" Then
    Dim one As Excel.Worksheet
    Application.DisplayAlerts = False
    For Each one In parentWorkbook.Worksheets
        If one.Name = worksheetName Then
            one.Delete
        End If
    Next one
    Application.DisplayAlerts = True
    Set one = parentWorkbook.Worksheets.Add(Before:=parentWorkbook.Worksheets(1))
    one.Name = worksheetName
    Set one = Nothing
Else
    Call resetClass
End If
End Function

Private Function isTwoDArray(ByRef sourceArr As Variant) As Boolean
On Error Resume Next
Dim i As Long
Dim j As Long
For i = 1 To 9
    j = -1
    j = UBound(sourceArr, i)
    If j = -1 Then
        If i = 3 Then
            isTwoDArray = True
        End If
        Exit For
    End If
Next i
End Function

Private Function resetClass()
sourceData = Empty
Set parentWorkbook = Nothing
worksheetName = ""
End Function

