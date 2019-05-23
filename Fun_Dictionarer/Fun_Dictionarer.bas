Attribute VB_Name = "Fun_Dictionarer"
Option Explicit

'attainDictObjectFromRange
'attainDictTwoDArrayFromRange
'attainDictOneDArrayFromRange

Private Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("Dictionarer")
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

Public Function attainDictObjectFromRange(ByRef sourceRange As Excel.Range, _
    Optional ByVal keyCol As Long = 1, Optional ByVal valueCol As Long = 1, _
    Optional ByVal valueType As String = "LastValue", _
    Optional ByRef sourceDict As Variant) As Variant
If IsMissing(sourceDict) = True Then
    Set attainDictObjectFromRange = _
        attainDictionaryFromRange("Dictionary", sourceRange, keyCol, valueCol, valueType)
Else
    Set attainDictObjectFromRange = _
        attainDictionaryFromRange("Dictionary", sourceRange, keyCol, valueCol, valueType, sourceDict)
End If
End Function

Public Function attainDictTwoDArrayFromRange(ByRef sourceRange As Excel.Range, _
    Optional ByVal keyCol As Long = 1, Optional ByVal valueCol As Long = 1, _
    Optional ByVal valueType As String = "LastValue", _
    Optional ByRef sourceDict As Variant) As Variant
If IsMissing(sourceDict) = True Then
    attainDictTwoDArrayFromRange = _
        attainDictionaryFromRange("TwoDimensionArray", sourceRange, keyCol, valueCol, valueType)
Else
    attainDictTwoDArrayFromRange = _
        attainDictionaryFromRange("TwoDimensionArray", sourceRange, keyCol, valueCol, valueType, sourceDict)
End If
End Function

Public Function attainDictOneDArrayFromRange(ByRef sourceRange As Excel.Range, _
    Optional ByVal keyCol As Long = 1, Optional ByVal valueCol As Long = 1, _
    Optional ByVal valueType As String = "LastValue", _
    Optional ByRef sourceDict As Variant) As Variant
If IsMissing(sourceDict) = True Then
    attainDictOneDArrayFromRange = _
        attainDictionaryFromRange("OneDimensionArray", sourceRange, keyCol, valueCol, valueType)
Else
    attainDictOneDArrayFromRange = _
        attainDictionaryFromRange("OneDimensionArray", sourceRange, keyCol, valueCol, valueType, sourceDict)
End If
End Function

Private Function attainDictionaryFromRange(ByVal targetObjectType As String, _
    ByRef sourceRange As Excel.Range, Optional ByVal keyCol As Long = 1, _
    Optional ByVal valueCol As Long = 1, Optional ByVal valueType As String = "LastValue", _
    Optional ByRef sourceDict As Variant) As Variant
Dim DT As New Dictionarer
Set DT.sourceRange = sourceRange
DT.keyCol = keyCol
DT.valueCol = valueCol
DT.valueType = valueType
If IsMissing(sourceDict) = False Then
    Set DT.dictObject = sourceDict
End If
DT.generateDictAndArray
Select Case targetObjectType
Case "Dictionary"
    Set attainDictionaryFromRange = DT.dictObject
Case "TwoDimensionArray"
    attainDictionaryFromRange = DT.twoDimensionArray
Case "OneDimensionArray"
    attainDictionaryFromRange = DT.oneDimensionArray
End Select
End Function



