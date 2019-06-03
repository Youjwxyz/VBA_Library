VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ranger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sheetObject As Excel.Worksheet
Private rangeAddress As String
Private rangeLocation As String
Private rangeObject As Excel.Range

Property Set excelRangeObject(ByRef sourceRange As Excel.Range)
If TypeName(sourceRange) = "Range" Then
    Set rangeObject = sourceRange
    Call synchronizeClass
Else
    Call resetClass
End If
End Property

Property Let rangeAddressName(ByVal addressOfRange As String)
addressOfRange = UCase(Replace(Replace(addressOfRange, "$", ""), " ", ""))
If TypeName(sheetObject) = "Worksheet" Then
    If addressOfRange = "USEDRANGE" Then
        addressOfRange = sheetObject.UsedRange.Address(False, False)
    End If
    If verifyAddressName(addressOfRange) = True Then
        rangeAddress = addressOfRange
        Set rangeObject = sheetObject.Range(rangeAddress)
        Call synchronizeClass
    Else
        Call resetClass
    End If
Else
    Call resetClass
End If
End Property

Property Let rangeLocationIndex(ByVal indexOfRange As String)
indexOfRange = Replace(indexOfRange, " ", "")
If TypeName(sheetObject) = "Worksheet" And verifyLocationIndex(indexOfRange) = True Then
    rangeLocation = indexOfRange
    Dim rangeIndex As Variant
    rangeIndex = Split(rangeLocation, ",")
    Dim i As Long
    i = UBound(rangeIndex)
    If i = 1 Then
        Set rangeObject = sheetObject.Range( _
            sheetObject.Cells(CLng(rangeIndex(0)), CLng(rangeIndex(1))))
    End If
    If i = 3 Then
        Set rangeObject = sheetObject.Range( _
            sheetObject.Cells(CLng(rangeIndex(0)), CLng(rangeIndex(1))), _
            sheetObject.Cells(CLng(rangeIndex(2)), CLng(rangeIndex(3))))
    End If
    Call synchronizeClass
Else
    Call resetClass
End If
End Property

Property Get excelRangeObject() As Excel.Range
If TypeName(rangeObject) = "Range" Then
    Set excelRangeObject = rangeObject
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Private Sub resetClass()
Set sheetObject = Nothing
rangeAddress = ""
rangeLocation = ""
Set rangeObject = Nothing
End Sub

Private Sub synchronizeClass()
If TypeName(rangeObject) = "Range" Then
    Set sheetObject = rangeObject.Parent
    rangeAddress = rangeObject.Address(False, False)
    Dim i As Long
    Dim j As Long
    i = rangeObject.Rows.Count
    j = rangeObject.Columns.Count
    If i = 1 And j = 1 Then
        rangeLocation = rangeObject.Cells(1, 1).Row & "," & rangeObject.Cells(1, 1).Column
    Else
        rangeLocation = rangeObject.Cells(1, 1).Row & "," & rangeObject.Cells(1, 1).Column & "," & _
            rangeObject.Cells(i, j).Row & "," & rangeObject.Cells(i, j).Column
    End If
End If
End Sub

Private Function verifyAddressName(ByVal addressOfRange As String) As Boolean
addressOfRange = Replace(Replace(addressOfRange, "$", ""), " ", "")
If addressOfRange <> "" Then
    Dim rangeName As Variant
    rangeName = Split(addressOfRange, ":")
    Dim i As Long
    i = UBound(rangeName)
    If i = 0 Or i = 1 Then
        Dim indicator As Boolean
        indicator = True
        Dim j As Long
        Dim subAddressName As Variant
        For j = 0 To i
            subAddressName = rangeName(j)
            If verifySubAddressName(subAddressName) = False Then
                indicator = False
                Exit For
            End If
        Next j
        verifyAddressName = indicator
    End If
End If
End Function

Private Function verifySubAddressName(ByVal subAddressOfRange As String) As Boolean
subAddressOfRange = Replace(Replace(subAddressOfRange, "$", ""), " ", "")
If subAddressOfRange <> "" Then
    Dim i As Long
    i = Len(subAddressOfRange)
    If i > 0 Then
        Dim indicator As Boolean
        indicator = True
        Dim j As Long
        Dim tempValue As Variant
        Dim numberIndicator As Boolean
        For j = 1 To i
            If numberIndicator = False Then
                tempValue = UCase(Mid(subAddressOfRange, j, 1))
                If IsNumeric(tempValue) = True Then
                    numberIndicator = True
                    If Asc(tempValue) = 48 Then
                        indicator = False
                        Exit For
                    End If
                End If
            End If
            If numberIndicator = False Then
                If Asc(tempValue) < 65 Or Asc(tempValue) > 90 Then
                    indicator = False
                    Exit For
                End If
            Else
                If Asc(tempValue) < 48 Or Asc(tempValue) > 57 Then
                    indicator = False
                    Exit For
                End If
            End If
        Next j
        verifySubAddressName = indicator
    End If
End If
End Function

Private Function verifyLocationIndex(ByVal indexOfRange As String) As Boolean
indexOfRange = Replace(indexOfRange, " ", "")
If indexOfRange <> "" Then
    Dim rangeIndex As Variant
    rangeIndex = Split(indexOfRange, ",")
    Dim i As Long
    i = UBound(rangeIndex)
    If i = 1 Or i = 3 Then
        Dim indicator As Boolean
        indicator = True
        Dim j As Long
        Dim tempValue As Variant
        For j = 0 To i
            tempValue = rangeIndex(j)
            If IsNumeric(tempValue) = False Then
                indicator = False
                Exit For
            Else
                If tempValue <= 0 Then
                    indicator = False
                    Exit For
                End If
            End If
        Next j
        verifyLocationIndex = indicator
    End If
End If
End Function

