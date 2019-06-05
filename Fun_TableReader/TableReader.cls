VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sourceSheet As Excel.Worksheet
Public valueType As String
Public titleAndDataRefArray As Variant
'TitleA, TitleB, TitleC, TitleD, ...
'D2, D5, $A$6, $N7, ...
'NotEmpty, , , NotEmpty, .,..
Public rowRepeat As Long
Public colRepeat As Long
Private dataArray As Variant

Private Sub resetClass()
Set sourceSheet = Nothing
valueType = ""
titleAndDataRefArray = Empty
rowRepeat = 0
colRepeat = 0
dataArray = Empty
End Sub

Public Sub extractTable()
If TypeName(sourceSheet) = "Worksheet" And TypeName(titleAndDataRefArray) = "Variant()" And valueType <> "" Then
    Dim maxRow As Long
    Dim maxCol As Long
    maxRow = sourceSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
    maxCol = sourceSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Column
    Dim a As Long
    Dim b As Long
    a = UBound(titleAndDataRefArray, 2)
    
    Dim tempArray() As Variant
    Dim rowDelta As Long
    Dim colDelta As Long
    Dim counter As Long
    Dim maxCount As Long
    If rowRepeat = 0 And colRepeat = 0 Then
        ReDim tempArray(1 To 2, 1 To a) As Variant
        Call addTitleToTargetArray(tempArray, titleAndDataRefArray)
        If verifyKeyFields(sourceSheet, titleAndDataRefArray, 0, 0) Then
            For b = 1 To a
                tempArray(2, b) = sourceSheet.Range(titleAndDataRefArray(2, b)).Value
            Next b
        End If
    Else
        maxCount = getMaxCountOfTempArray(rowRepeat, maxRow, colRepeat, maxCol)
        ReDim tempArray(1 To maxCount, 1 To a) As Variant
        Call addTitleToTargetArray(tempArray, titleAndDataRefArray)
        counter = 1
        Do Until verifyCellPosition(sourceSheet, titleAndDataRefArray, rowDelta, maxRow, colDelta, maxCol) = False
            If verifyKeyFields(sourceSheet, titleAndDataRefArray, rowDelta, colDelta) Then
                counter = counter + 1
                For b = 1 To a
                    Select Case valueType
                    Case "Value"
                        tempArray(counter, b) = adjustedOffset(sourceSheet, titleAndDataRefArray(2, b), rowDelta, colDelta)
                    Case "Address"
                        tempArray(counter, b) = adjustedOffsetAddress(sourceSheet, titleAndDataRefArray(2, b), rowDelta, colDelta)
                    End Select
                Next b
            End If
            If colRepeat <> 0 Then
                colDelta = colDelta + colRepeat
                If verifyCellPosition(sourceSheet, titleAndDataRefArray, rowDelta, maxRow, colDelta, maxCol) = False Then
                    rowDelta = rowDelta + rowRepeat
                    colDelta = 0
                End If
            Else
                rowDelta = rowDelta + rowRepeat
            End If
        Loop
    End If
    dataArray = tempArray
    Erase tempArray
End If
End Sub

Property Get tableData() As Variant
If TypeName(dataArray) = "Variant()" Then
    tableData = dataArray
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Private Function getMaxCountOfTempArray(ByVal rowRepeatSub As Long, ByVal maxRow As Long, _
    ByVal colRepeatSub As Long, ByVal maxCol As Long) As Long
Dim rowCount As Long
Dim colCount As Long
rowCount = 1
colCount = 1
If maxRow > 0 And maxCol > 0 Then
    If rowRepeatSub <> 0 Then
        rowCount = Abs(Round(maxRow / rowRepeatSub, 0)) + 1
    End If
    If colRepeatSub <> 0 Then
        colCount = Abs(Round(maxCol / colRepeatSub, 0)) + 1
    End If
End If
getMaxCountOfTempArray = rowCount * colCount
End Function


Private Function adjustedOffsetAddress(ByRef targetSheet As Excel.Worksheet, ByVal cellAddress As String, _
    ByVal rowOffset As Long, ByVal colOffset As Long) As Variant
If TypeName(targetSheet) = "Worksheet" And cellAddress <> "" Then
    Dim tempValue As Variant
    Dim one As Excel.Range
    Set one = targetSheet.Range(cellAddress)
    If Left(cellAddress, 1) = "$" Then
        If InStr(2, cellAddress, "$") > 0 Then
            tempValue = one.Address(False, False)
        Else
            tempValue = targetSheet.Cells(one.Row + rowOffset, one.Column).Address(False, False)
        End If
    Else
        If InStr(2, cellAddress, "$") > 0 Then
            tempValue = targetSheet.Cells(one.Row, one.Column + colOffset).Address(False, False)
        Else
            tempValue = targetSheet.Cells(one.Row + rowOffset, one.Column + colOffset).Address(False, False)
        End If
    End If
    adjustedOffsetAddress = tempValue
End If
End Function

Private Function adjustedOffset(ByRef targetSheet As Excel.Worksheet, ByVal cellAddress As String, _
    ByVal rowOffset As Long, ByVal colOffset As Long) As Variant
If TypeName(targetSheet) = "Worksheet" And cellAddress <> "" Then
    Dim tempValue As Variant
    Dim one As Excel.Range
    Set one = targetSheet.Range(cellAddress)
    If Left(cellAddress, 1) = "$" Then
        If InStr(2, cellAddress, "$") > 0 Then
            tempValue = one.Value
        Else
            tempValue = targetSheet.Cells(one.Row + rowOffset, one.Column).Value
        End If
    Else
        If InStr(2, cellAddress, "$") > 0 Then
            tempValue = targetSheet.Cells(one.Row, one.Column + colOffset).Value
        Else
            tempValue = targetSheet.Cells(one.Row + rowOffset, one.Column + colOffset).Value
        End If
    End If
    adjustedOffset = tempValue
End If
End Function


Private Function verifyKeyFields(ByRef targetSheet As Excel.Worksheet, _
    ByRef targetTitleAndDataRefArray As Variant, ByVal rowOffset As Long, ByVal colOffset As Long) As Boolean
If TypeName(targetSheet) = "Worksheet" And TypeName(targetTitleAndDataRefArray) = "Variant()" Then
    Dim i As Long
    Dim j As Long
    Dim cellAddress As String
    Dim indicator As Boolean
    indicator = True
    i = UBound(targetTitleAndDataRefArray, 2)
    For j = 1 To i
        cellAddress = targetTitleAndDataRefArray(2, j)
        Select Case targetTitleAndDataRefArray(3, j)
        Case "NotEmpty"
            If adjustedOffset(targetSheet, cellAddress, rowOffset, colOffset) = "" Then
                indicator = False
                Exit For
            End If
        Case ""
            'Do Nothing
        Case Else
            If adjustedOffset(targetSheet, cellAddress, rowOffset, colOffset) <> targetTitleAndDataRefArray(3, j) Then
                indicator = False
                Exit For
            End If
        End Select
    Next j
    verifyKeyFields = indicator
End If
End Function

Private Function verifyCellPosition(ByRef targetSheet As Excel.Worksheet, _
    ByRef targetTitleAndDataRefArray As Variant, ByVal rowOffset As Long, ByVal maxRow As Long, _
    ByVal colOffset As Long, ByVal maxCol As Long) As Boolean
If TypeName(targetSheet) = "Worksheet" And TypeName(targetTitleAndDataRefArray) = "Variant()" Then
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim n As Long
    Dim a As Long
    Dim b As Long
    Dim cellAddress As String
    Dim indicator As Boolean
    indicator = True
    i = UBound(targetTitleAndDataRefArray, 2)
    For j = 1 To i
        m = targetSheet.Range(targetTitleAndDataRefArray(2, j)).Row
        n = targetSheet.Range(targetTitleAndDataRefArray(2, j)).Column
        cellAddress = targetTitleAndDataRefArray(2, j)
        If Left(cellAddress, 1) = "$" Then
            If InStr(2, cellAddress, "$") > 0 Then
                a = 0
                b = 0
            Else
                a = rowOffset
                b = 0
            End If
        Else
            If InStr(2, cellAddress, "$") > 0 Then
                a = 0
                b = colOffset
            Else
                a = rowOffset
                b = colOffset
            End If
        End If
        If m + a < 1 Or n + b < 1 Or m + a > maxRow Or n + b > maxCol Then
            indicator = False
            Exit For
        End If
    Next j
    verifyCellPosition = indicator
End If
End Function

Private Sub addTitleToTargetArray(ByRef targetArray As Variant, ByRef sourceArray As Variant)
If TypeName(targetArray) = "Variant()" And TypeName(sourceArray) = "Variant()" Then
    If UBound(targetArray, 2) = UBound(sourceArray, 2) Then
        Dim i As Long
        Dim j As Long
        i = UBound(targetArray, 2)
        For j = 1 To i
            targetArray(1, j) = sourceArray(1, j)
        Next j
    End If
End If
End Sub
