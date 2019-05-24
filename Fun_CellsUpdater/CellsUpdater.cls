VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CellsUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public cellID As String
Public cellObject As Excel.Range
Public cellValue As Variant
Public valueType As String
Private cellList As Variant
Private valueList As Variant

Property Set excelCellList(ByRef sourceCellList As Variant)
If TypeName(sourceCellList) = "Dictionary" Then
    Set cellList = sourceCellList
Else
    Call resetClass
End If
End Property

Property Get excelCellList() As Variant
If TypeName(cellList) = "Dictionary" Then
    Set excelCellList = cellList
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Property Set excelValueList(ByRef sourceValueList As Variant)
If TypeName(sourceValueList) = "Dictionary" Then
    Set valueList = sourceValueList
Else
    Call resetClass
End If
End Property

Property Get excelValueList() As Variant
If TypeName(valueList) = "Dictionary" Then
    Set excelValueList = valueList
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Public Sub addCell()
If cellID <> "" And TypeName(cellObject) = "Range" Then
    If TypeName(cellList) <> "Dictionary" Then
        Set cellList = CreateObject("Scripting.Dictionary")
    End If
    If cellObject.Cells.Count = 1 Then
        Set cellList.Item(cellID) = cellObject
    Else
        Call resetClass
    End If
Else
    Call resetClass
End If
End Sub

Public Sub defineCellValue()
If valueType <> "" And cellID <> "" And TypeName(cellValue) <> "Empty" Then
    If TypeName(valueList) <> "Dictionary" Then
        Set valueList = CreateObject("Scripting.Dictionary")
    End If
    valueList.Item(cellID) = cellValue
Else
    Call resetClass
End If
End Sub

Public Sub updateCells()
If valueType <> "" And TypeName(cellList) = "Dictionary" And TypeName(valueList) = "Dictionary" Then
    Application.ScreenUpdating = False
    Dim one As Variant
    For Each one In cellList.keys()
        If valueList.exists(one) Then
            Select Case valueType
            Case "Value"
                cellList(one).value = valueList(one)
            Case "Formula"
                cellList(one).Formula = valueList(one)
            Case Else
                MsgBox "Error in valueType, which should be either Value or Formula."
            End Select
        End If
    Next one
    Application.ScreenUpdating = True
Else
    Call resetClass
End If
End Sub

Private Sub resetClass()
cellID = ""
Set cellObject = Nothing
cellValue = Empty
valueType = ""
cellList = Empty
valueList = Empty
End Sub
