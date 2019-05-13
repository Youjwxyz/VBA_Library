Attribute VB_Name = "Fun_MultiDimensionArr"
Option Explicit

'attainMultiDimensionData

Public Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("MultiDArrayer")
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

Sub attainMultiDimensionData(ByRef rowDimension As Excel.Range, ByRef colDimension As Excel.Range, _
    ByRef titleArr As Variant, ByVal outputSheetName As String)
Dim MDA As New MultiDArrayer
Set MDA.rowDemension = rowDimension
Set MDA.colDemension = colDimension
MDA.generateDataArray
MDA.titleArr = titleArr
MDA.outputSheet = outputSheetName
MDA.outputDataArray
End Sub
