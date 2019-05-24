Attribute VB_Name = "Fun_CellsUpdater"
Option Explicit

'addUpdateTargetCellObject
'addUpdateTargetCellValue
'updateCellValueToExcel

Private Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("CellsUpdater")
Dim classFolder As String
classFolder = "C:\Business\Macros\VBA_Class"
Dim one As Variant
Dim com As Variant
Dim comFound As Boolean
For Each one In classArr
    comFound = False
    For Each com In ThisWorkbook.VBProject.VBComponents
        If com.Name = one Then
            comFound = True
            Exit For
        End If
    Next com
    If Not comFound Then
        ThisWorkbook.VBProject.VBComponents.Import classFolder & "\" & one & ".cls"
    End If
Next one
End Sub

Public Function addUpdateTargetCellObject(ByVal cellUniquer As String, _
    ByRef cellObject As Excel.Range, Optional ByRef sourceCellList As Variant) As Variant
Dim CU As New CellsUpdater
CU.cellID = cellUniquer
Set CU.cellObject = cellObject
If TypeName(sourceCellList) = "Dictionary" Then
    Set CU.excelCellList = sourceCellList
End If
CU.addCell
Set addUpdateTargetCellObject = CU.excelCellList
End Function

Public Function addUpdateTargetCellValue(ByVal cellUniquer As String, _
    ByRef cellValue As Variant, ByVal valueType As String, _
    Optional ByRef sourceValueList As Variant) As Variant
Dim CU As New CellsUpdater
CU.cellID = cellUniquer
CU.cellValue = cellValue
CU.valueType = valueType
If TypeName(sourceValueList) = "Dictionary" Then
    Set CU.excelValueList = sourceValueList
End If
CU.defineCellValue
Set addUpdateTargetCellValue = CU.excelValueList
End Function

Public Sub updateCellValueToExcel(ByRef sourceCellList As Variant, _
    ByRef sourceValueList As Variant, ByVal valueType As String)
Dim CU As New CellsUpdater
CU.valueType = valueType
Set CU.excelCellList = sourceCellList
Set CU.excelValueList = sourceValueList
CU.updateCells
End Sub

