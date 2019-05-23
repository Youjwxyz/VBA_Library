VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TwoDArrayExpander"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dataArray As Variant
Private newColumns As Variant
Private newRows As Variant

Property Let twoDimensionArray(ByRef sourceArray As Variant)
If isTwoDArray(sourceArray) = True Then
    dataArray = sourceArray
Else
    Call resetClass
End If
End Property

Property Let addNewColumns(ByRef columnsToBeAdded As Variant)
If (TypeName(columnsToBeAdded) = "String()" Or TypeName(columnsToBeAdded) = "String") And _
    isTwoDArray(dataArray) = True Then
        newColumns = columnsToBeAdded
        Dim i As Long
        Dim j As Long
        i = UBound(dataArray, 1)
        j = UBound(dataArray, 2)
        If TypeName(newColumns) = "String" Then
            ReDim Preserve dataArray(1 To i, 1 To j + 1) As Variant
            If newColumns <> "" Then
                dataArray(1, j + 1) = newColumns
            End If
        Else
            Dim k As Long
            Dim l As Long
            k = UBound(newColumns)
            ReDim Preserve dataArray(1 To i, 1 To j + k + 1) As Variant
            For l = 0 To k
                If newColumns(l) <> "" Then
                    dataArray(1, j + l + 1) = newColumns(l)
                End If
            Next l
        End If
Else
    Call resetClass
End If
End Property

Property Let addNewRows(ByRef rowsToBeAdded As Variant)
If (TypeName(rowsToBeAdded) = "Integer" Or TypeName(rowsToBeAdded) = "Long") And _
    isTwoDArray(dataArray) = True Then
        newRows = rowsToBeAdded
        Dim i As Long
        Dim j As Long
        i = UBound(dataArray, 1)
        j = UBound(dataArray, 2)
        Dim tempArray() As Variant
        ReDim tempArray(1 To i + newRows, 1 To j) As Variant
        Dim m As Long
        Dim n As Long
        For m = 1 To i
            For n = 1 To j
                tempArray(m, n) = dataArray(m, n)
            Next n
        Next m
        dataArray = tempArray
        Erase tempArray
Else
    Call resetClass
End If
End Property

Property Get twoDimensionArray() As Variant
If (TypeName(newColumns) = "String()" Or TypeName(newColumns) = "String" Or _
    TypeName(newRows) = "Integer" Or TypeName(newRows) = "Long") And _
    isTwoDArray(dataArray) = True Then
        twoDimensionArray = dataArray
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Private Function resetClass()
dataArray = Empty
newColumns = Empty
newRows = Empty
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



