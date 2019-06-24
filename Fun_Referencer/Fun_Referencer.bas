Attribute VB_Name = "Fun_Referencer"
Option Explicit

'addReferenceData

Private Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("ACCESSer", "Referencer")
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

Public Sub addReferenceData(ByRef sourceRange As Excel.Range, ByRef keyCols As Variant)
Dim RFC As New Referencer
Set RFC.sourceRange = sourceRange
RFC.refCols = keyCols
RFC.createReference
Set RFC = Nothing
End Sub
