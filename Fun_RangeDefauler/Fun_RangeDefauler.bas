Attribute VB_Name = "Fun_RangeDefauler"
Option Explicit

'assignDefaultValueToRange
'attainRangeAfterValueUpdated

Private Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("RangerDefauler")
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

Public Sub assignDefaultValueToRange(ByRef targetRange As Excel.Range, _
    ByVal columnIndex As Long, ByVal defaultValue As Variant, _
    Optional ByVal withTitleRange As Boolean = True, Optional ByVal overWrite As Boolean = False)
Dim RD As New RangerDefauler
RD.columnsIndex = columnIndex
RD.defaultValue = defaultValue
RD.rangeWithTitle = withTitleRange
RD.overWrite = overWrite
Set RD.excelRange = targetRange
End Sub

Public Function attainRangeAfterValueUpdated(ByRef targetRange As Excel.Range, _
    ByVal columnIndex As Long, ByVal defaultValue As Variant, _
    Optional ByVal withTitleRange As Boolean = True, Optional ByVal overWrite As Boolean = False) As Excel.Range
Dim RD As New RangerDefauler
RD.columnsIndex = columnIndex
RD.defaultValue = defaultValue
RD.rangeWithTitle = withTitleRange
RD.overWrite = overWrite
Set RD.excelRange = targetRange
Set attainRangeAfterValueUpdated = RD.excelRange
End Function

