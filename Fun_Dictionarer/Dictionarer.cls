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
Public dict As Object
Public oneDArray As Variant
Public twoDArray As Variant

Public Sub fromRange()
If Not sourceRange Is Nothing And keyCol > 0 And valueCol > 0 Then
    If keyCol <= sourceRange.Columns.Count And valueCol <= sourceRange.Columns.Count Then
        If Not TypeName(dict) = "Dictionary" Then
            Set dict = CreateObject("Scripting.Dictionary")
        End If
        Dim i As Long
        Dim j As Long
        i = sourceRange.Rows.Count
        For j = 1 To i
            Select Case valueType
            Case "lastValue"
                dict.Item(sourceRange(j, keyCol).Value) = sourceRange(j, valueCol).Value
            Case "countKey"
                dict.Item(sourceRange(j, keyCol).Value) = dict.Item(sourceRange(j, keyCol).Value) + 1
            Case "sumValue"
                dict.Item(sourceRange(j, keyCol).Value) = dict.Item(sourceRange(j, keyCol).Value) + _
                    sourceRange(j, valueCol).Value
            End Select
        Next j
    Else
        Call resetClass
    End If
Else
    Call resetClass
End If
End Sub

Public Function toTwoDArray() As Variant
If TypeName(dict) = "Dictionary" Then
    Dim i As Long
    i = dict.Count
    If i > 0 Then
        Dim Temp() As Variant
        ReDim Temp(1 To i, 1 To 2) As Variant
        Dim j As Long
        For j = 1 To i
            Temp(j, 1) = dict.keys()(j - 1)
            Temp(j, 2) = dict(dict.keys()(j - 1))
        Next j
        toTwoDArray = Temp
        Erase Temp
    Else
        Call resetClass
    End If
Else
    Call resetClass
End If
End Function

Public Function toOneDArray() As Variant
If TypeName(dict) = "Dictionary" Then
    Dim i As Long
    i = dict.Count
    If i > 0 Then
        Dim Temp() As Variant
        ReDim Temp(0 To i - 1) As Variant
        Dim j As Long
        For j = 0 To i - 1
            Temp(j) = dict.keys()(j)
        Next j
        toOneDArray = Temp
        Erase Temp
    Else
        Call resetClass
    End If
Else
    Call resetClass
End If
End Function

Private Sub resetClass()
Set sourceRange = Nothing
keyCol = 0
valueCol = 0
valueType = ""
Set dict = Nothing
Set oneDArray = Nothing
Set twoDArray = Nothing
End Sub
