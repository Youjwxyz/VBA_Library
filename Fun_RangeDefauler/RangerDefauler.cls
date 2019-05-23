VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangerDefauler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public columnsIndex As Long
Public defaultValue As Variant
Public rangeWithTitle As Boolean
Public overWrite As Boolean
Private targetRange As Excel.Range

Property Set excelRange(ByRef sourceRange As Excel.Range)
If TypeName(sourceRange) = "Range" And columnsIndex > 0 And IsEmpty(defaultValue) = False Then
    Set targetRange = sourceRange
    Dim i As Long
    i = targetRange.Columns(1).Cells.Count
    Dim j As Long
    Dim one As Excel.Range
    For j = 1 To i
        Set one = targetRange.Columns(columnsIndex).Cells(j, 1)
        If rangeWithTitle = True Then
            If j > 1 Then
                If one.Value = "" Then
                    one.Value = defaultValue
                Else
                    If overWrite = True Then
                        one.Value = defaultValue
                    End If
                End If
            End If
        Else
            If one.Value = "" Then
                one.Value = defaultValue
            Else
                If overWrite = True Then
                    one.Value = defaultValue
                End If
            End If
        End If
    Next j
Else
    Call resetClass
End If
End Property

Property Get excelRange() As Excel.Range
If TypeName(targetRange) = "Range" Then
    Set excelRange = targetRange
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Private Sub resetClass()
columnsIndex = 0
defaultValue = Empty
rangeWithTitle = True
overWrite = False
Set targetRange = Nothing
End Sub
