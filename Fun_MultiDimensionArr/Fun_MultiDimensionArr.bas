Attribute VB_Name = "Fun_MultiDimensionArr"
Option Explicit

'attainArrayFromMultiDimRange

Private Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("MultiDArrayGenerator")
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

Public Function attainArrayFromMultiDimRange(ByRef rowDimension As Excel.Range, _
    ByRef colDimension As Excel.Range, ByRef titleArr As Variant) As Variant
Dim MDAG As New MultiDArrayGenerator
Set MDAG.rowDemension = rowDimension
Set MDAG.colDemension = colDimension
MDAG.titleArr = titleArr
MDAG.generateDataArray
attainArrayFromMultiDimRange = MDAG.outputArray
End Function
