Attribute VB_Name = "Fun_Dictionarer"
Option Explicit

'createLastValueDictFromRange
'createSumValueDictFromRange
'createCountKeyDictFromRange
'appendLastValueDictFromRange
'appendSumValueDictFromRange
'appendCountKeyDictFromRange
'convertDictToTwoDimensionArray
'convertDictToOneDimensionArray

Public Sub importFunctionRelatedClass()
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

Public Function createLastValueDictFromRange(ByRef sourceRange As Excel.Range, _
    ByVal keyCol As Long, ByVal valueCol As Long) As Object
    Set createLastValueDictFromRange = _
        createDictionaryFromRange(sourceRange, keyCol, valueCol, "lastValue")
End Function

Public Function createSumValueDictFromRange(ByRef sourceRange As Excel.Range, _
    ByVal keyCol As Long, ByVal valueCol As Long) As Object
    Set createSumValueDictFromRange = _
        createDictionaryFromRange(sourceRange, keyCol, valueCol, "sumValue")
End Function

Public Function createCountKeyDictFromRange(ByRef sourceRange As Excel.Range, _
    ByVal keyCol As Long, ByVal valueCol As Long) As Object
    Set createCountKeyDictFromRange = _
        createDictionaryFromRange(sourceRange, keyCol, valueCol, "countKey")
End Function

Public Function appendLastValueDictFromRange(ByRef sourceDict As Object, _
    ByRef sourceRange As Excel.Range, ByVal keyCol As Long, ByVal valueCol As Long) As Object
    Set appendLastValueDictFromRange = _
        appendDictionaryFromRange(sourceDict, sourceRange, keyCol, valueCol, "lastValue")
End Function

Public Function appendSumValueDictFromRange(ByRef sourceDict As Object, _
    ByRef sourceRange As Excel.Range, ByVal keyCol As Long, ByVal valueCol As Long) As Object
    Set appendSumValueDictFromRange = _
        appendDictionaryFromRange(sourceDict, sourceRange, keyCol, valueCol, "sumValue")
End Function

Public Function appendCountKeyDictFromRange(ByRef sourceDict As Object, _
    ByRef sourceRange As Excel.Range, ByVal keyCol As Long, ByVal valueCol As Long) As Object
    Set appendCountKeyDictFromRange = _
        appendDictionaryFromRange(sourceDict, sourceRange, keyCol, valueCol, "countKey")
End Function

Public Function convertDictToTwoDimensionArray(ByRef sourceDict As Object) As Variant
If TypeName(sourceDict) = "Dictionary" Then
    Dim Temp As New Dictionarer
    Set Temp.dict = sourceDict
    convertDictToTwoDimensionArray = Temp.toTwoDArray
    Set Temp = Nothing
End If
End Function

Public Function convertDictToOneDimensionArray(ByRef sourceDict As Object) As Variant
If TypeName(sourceDict) = "Dictionary" Then
    Dim Temp As New Dictionarer
    Set Temp.dict = sourceDict
    convertDictToOneDimensionArray = Temp.toOneDArray
    Set Temp = Nothing
End If
End Function

Private Function createDictionaryFromRange(ByRef sourceRange As Excel.Range, _
    ByVal keyCol As Long, ByVal valueCol As Long, ByVal valueType As String) As Object
If Not sourceRange Is Nothing And keyCol > 0 And valueCol > 0 Then
    If keyCol <= sourceRange.Columns.Count And valueCol <= sourceRange.Columns.Count Then
        Dim Temp As New Dictionarer
        Set Temp.sourceRange = sourceRange
        Temp.keyCol = keyCol
        Temp.valueCol = valueCol
        Temp.valueType = valueType
        Temp.fromRange
        Set createDictionaryFromRange = Temp.dict
        Set Temp = Nothing
    End If
End If
End Function

Private Function appendDictionaryFromRange(ByRef sourceDict As Object, _
    ByRef sourceRange As Excel.Range, ByVal keyCol As Long, ByVal valueCol As Long, _
    ByVal valueType As String) As Object
If TypeName(sourceDict) = "Dictionary" And Not sourceRange Is Nothing And keyCol > 0 And valueCol > 0 Then
    If keyCol <= sourceRange.Columns.Count And valueCol <= sourceRange.Columns.Count Then
        Dim Temp As New Dictionarer
        Set Temp.sourceRange = sourceRange
        Temp.keyCol = keyCol
        Temp.valueCol = valueCol
        Temp.valueType = valueType
        Set Temp.dict = sourceDict
        Temp.fromRange
        Set appendDictionaryFromRange = Temp.dict
        Set Temp = Nothing
    End If
End If
End Function

