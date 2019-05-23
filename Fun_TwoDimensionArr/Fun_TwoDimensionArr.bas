Attribute VB_Name = "Fun_TwoDimensionArr"
Option Explicit

'addColumnsRowsToDataArray

Private Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("TwoDArrayExpander")
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

Public Function addColumnsRowsToDataArray(ByRef sourceArray As Variant, _
    Optional ByRef newColumns As Variant, Optional ByRef newRows As Variant) As Variant
Dim TDAE As New TwoDArrayExpander
TDAE.twoDimensionArray = sourceArray
If IsMissing(newColumns) = False Then
    TDAE.addNewColumns = newColumns
End If
If IsMissing(newRows) = False Then
    TDAE.addNewRows = newRows
End If
addColumnsRowsToDataArray = TDAE.twoDimensionArray
End Function
