VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dictionarer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sourceRange As Excel.Range
Public keyCol As Long
Public valueCol As Long
Public valueType As String
Private dict As Variant
Private twoDArray As Variant
Private oneDArray As Variant

Property Set dictObject(ByRef sourceDict As Variant)
If TypeName(sourceDict) = "Dictionary" Then
    Set dict = sourceDict
Else
    Call resetClass
End If
End Property

Public Sub generateDictAndArray()
If TypeName(sourceRange) = "Range" And keyCol > 0 And valueCol > 0 Then
    If keyCol <= sourceRange.Columns.Count And valueCol <= sourceRange.Columns.Count Then
        If Not TypeName(dict) = "Dictionary" Then
            Set dict = CreateObject("Scripting.Dictionary")
        End If
        Dim i As Long
        Dim j As Long
        i = sourceRange.Rows.Count
        For j = 1 To i
            Select Case valueType
            Case "LastValue"
                dict.Item(sourceRange(j, keyCol).Value) = sourceRange(j, valueCol).Value
            Case "CountKey"
                dict.Item(sourceRange(j, keyCol).Value) = dict.Item(sourceRange(j, keyCol).Value) + 1
            Case "SumValue"
                dict.Item(sourceRange(j, keyCol).Value) = dict.Item(sourceRange(j, keyCol).Value) + _
                    sourceRange(j, valueCol).Value
            End Select
        Next j
        twoDArray = toTwoDArray(dict)
        oneDArray = toOneDArray(dict)
    Else
        Call resetClass
    End If
Else
    Call resetClass
End If
End Sub

Property Get dictObject() As Variant
If TypeName(dict) = "Dictionary" Then
    Set dictObject = dict
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Property Get twoDimensionArray() As Variant
If TypeName(twoDArray) = "Variant()" Then
    twoDimensionArray = twoDArray
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Property Get oneDimensionArray() As Variant
If TypeName(oneDArray) = "Variant()" Then
    oneDimensionArray = oneDArray
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Private Sub resetClass()
Set sourceRange = Nothing
keyCol = 0
valueCol = 0
valueType = ""
dict = Empty
oneDArray = Empty
twoDArray = Empty
End Sub

Private Function toTwoDArray(ByRef dictObject As Variant) As Variant
If TypeName(dictObject) = "Dictionary" Then
    Dim i As Long
    i = dictObject.Count
    If i > 0 Then
        Dim Temp() As Variant
        ReDim Temp(1 To i, 1 To 2) As Variant
        Dim j As Long
        For j = 1 To i
            Temp(j, 1) = dictObject.keys()(j - 1)
            Temp(j, 2) = dictObject(dictObject.keys()(j - 1))
        Next j
        toTwoDArray = Temp
        Erase Temp
    End If
End If
End Function

Private Function toOneDArray(ByRef dictObject As Variant) As Variant
If TypeName(dictObject) = "Dictionary" Then
    Dim i As Long
    i = dictObject.Count
    If i > 0 Then
        Dim Temp() As Variant
        ReDim Temp(0 To i - 1) As Variant
        Dim j As Long
        For j = 0 To i - 1
            Temp(j) = dictObject.keys()(j)
        Next j
        toOneDArray = Temp
        Erase Temp
    End If
End If
End Function
