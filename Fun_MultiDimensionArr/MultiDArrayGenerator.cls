VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MultiDArrayGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rowDemension As Excel.Range
Public colDemension As Excel.Range
Public titleArr As Variant
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
        Dim A As Long
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
        For A = 1 To i
            For b = 1 To n
                For c = 1 To j + m + 1
                    Select Case c
                    Case Is <= j
                        tempArray((A - 1) * n + b, c) = rowDemension.Cells(A, c).Value
                    Case Is <= j + m
                        tempArray((A - 1) * n + b, c) = colDemension.Cells(c - j, b).Value
                    Case Is > j + m
                        tempArray((A - 1) * n + b, c) = rowDemension.Parent.Cells(h + A - 1, l + b - 1).Value
                    End Select
                Next c
            Next b
        Next A
        dataArray = addTitleToDataArray(titleArr, tempArray)
        Erase tempArray
Else
    Call resetClass
End If
End Sub

Property Get outputArray() As Variant
If TypeName(dataArray) = "Variant()" Then
    outputArray = dataArray
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Private Sub resetClass()
Set rowDemension = Nothing
Set colDemension = Nothing
titleArr = Empty
dataArray = Empty
End Sub

Private Function addTitleToDataArray(ByRef sourceTitleArray As Variant, _
    ByRef sourceDataArray As Variant) As Variant
If TypeName(sourceTitleArray) = "String()" And TypeName(sourceDataArray) = "Variant()" Then
    Dim A As Long
    Dim i As Long
    Dim j As Long
    A = UBound(sourceTitleArray)
    i = UBound(sourceDataArray, 1)
    j = UBound(sourceDataArray, 2)
    If i > 0 And j > 0 And A = j - 1 Then
        Dim finalArray() As Variant
        ReDim finalArray(1 To i + 1, 1 To j) As Variant
        Dim m As Long
        Dim n As Long
        For m = 1 To i + 1
            For n = 1 To j
                If m = 1 Then
                    finalArray(m, n) = sourceTitleArray(n - 1)
                Else
                    finalArray(m, n) = sourceDataArray(m - 1, n)
                End If
            Next n
        Next m
        addTitleToDataArray = finalArray
        Erase finalArray
    End If
End If
End Function

