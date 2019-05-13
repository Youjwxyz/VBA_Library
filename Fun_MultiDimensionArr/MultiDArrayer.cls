VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiDArrayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rowDemension As Excel.Range
Public colDemension As Excel.Range
Public titleArr As Variant
Public outputSheet As String
Private dataArray As Variant

Public Sub generateDataArray()
If TypeName(rowDemension) = "Range" And TypeName(colDemension) = "Range" And _
    rowDemension.Parent.Name = colDemension.Parent.Name And _
    rowDemension.Parent.Parent.FullName = colDemension.Parent.Parent.FullName Then
        Dim h As Long
        Dim i As Long
        Dim j As Long
        Dim l As Long
        Dim m As Long
        Dim n As Long
        Dim a As Long
        Dim b As Long
        Dim c As Long
        h = rowDemension.Cells(1, 1).Row
        i = rowDemension.Rows.Count
        j = rowDemension.Columns.Count
        l = colDemension.Cells(1, 1).Column
        m = colDemension.Rows.Count
        n = colDemension.Columns.Count
        Dim tempArray() As Variant
        ReDim tempArray(1 To i * n, 1 To j + m + 1) As Variant
        For a = 1 To i
            For b = 1 To n
                For c = 1 To j + m + 1
                    Select Case c
                    Case Is <= j
                        tempArray((a - 1) * n + b, c) = rowDemension.Cells(a, c).Value
                    Case Is <= j + m
                        tempArray((a - 1) * n + b, c) = colDemension.Cells(c - j, b).Value
                    Case Is > j + m
                        tempArray((a - 1) * n + b, c) = rowDemension.Parent.Cells(h + a - 1, l + b - 1).Value
                    End Select
                Next c
            Next b
        Next a
        dataArray = tempArray
        Erase tempArray
Else
    Call resetClass
End If
End Sub

Public Sub outputDataArray()
If TypeName(titleArr) = "String()" And TypeName(dataArray) = "Variant()" And outputSheet <> "" Then
    Dim a As Long
    Dim i As Long
    Dim j As Long
    a = UBound(titleArr)
    i = UBound(dataArray, 1)
    j = UBound(dataArray, 2)
    If i > 0 And j > 0 And a = j - 1 Then
        Call resetOutputSheet
        With ThisWorkbook.Worksheets(outputSheet)
            Dim k As Long
            For k = 1 To j
                .Cells(1, k).Value = titleArr(k - 1)
            Next k
            .Range("A2").Resize(i, j).Value = dataArray
        End With
    Else
        Call resetClass
    End If
Else
    Call resetClass
End If
End Sub

Private Function resetOutputSheet()
If outputSheet <> "" Then
    Dim one As Excel.Worksheet
    Application.DisplayAlerts = False
    For Each one In ThisWorkbook.Worksheets
        If UCase(one.Name) = UCase(outputSheet) Then
            one.Delete
        End If
    Next one
    Application.DisplayAlerts = True
    Set one = ThisWorkbook.Worksheets.Add(before:=ThisWorkbook.Worksheets(1))
    one.Name = outputSheet
Else
    Call resetClass
End If
End Function

Private Sub resetClass()
Set rowDemension = Nothing
Set colDemension = Nothing
outputSheet = ""
titleArr = Empty
dataArray = Empty
End Sub

