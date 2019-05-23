Attribute VB_Name = "Fun_ExcelRanger"
Option Explicit

'attainExcelRange
'attainArrayFromRange

Private Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("Workbooker", "Worksheeter", "Ranger")
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

Function attainExcelRange(ByVal workbookFullPath As String, ByVal worksheetName As String, _
    Optional ByVal rangeAddress As String = "UsedRange", Optional ByVal rangeLocation As String = "", _
    Optional ByVal bookUpdateLink As Boolean = False, Optional ByVal bookReadOnly As Boolean = True, Optional ByVal bookPassword As String = "", Optional ByVal bookIngoreReadOnly As Boolean = True, _
    Optional ByVal sheetRemoveExistingSheet As Boolean = False, Optional ByVal sheetCreateIfNotFound As Boolean = True) As Excel.Range
If workbookFullPath <> "" And worksheetName <> "" Then
    Dim WB As New Workbooker
    WB.updateLink = bookUpdateLink
    WB.readOnly = bookReadOnly
    WB.password = bookPassword
    WB.ignoreReadOnly = bookIngoreReadOnly
    WB.workbookFullPath = workbookFullPath
    
    Dim WS As New Worksheeter
    Set WS.bookObject = WB.workbookObject
    WS.createIfNotFound = sheetCreateIfNotFound
    WS.removeExistingSheet = sheetRemoveExistingSheet
    WS.worksheetName = worksheetName
    
    Dim R As New Ranger
    Set R.sheetObject = WS.worksheetObject
    If rangeLocation = "" Then
        R.rangeAddressName = rangeAddress
        Set attainExcelRange = R.excelRangeObject
    Else
        R.rangeLocationIndex = rangeLocation
        Set attainExcelRange = R.excelRangeObject
    End If
End If
End Function

Function attainArrayFromRange(ByVal workbookFullPath As String, ByVal worksheetName As String, _
    Optional ByVal rangeAddress As String = "UsedRange", Optional ByVal rangeLocation As String = "", _
    Optional ByVal bookUpdateLink As Boolean = False, Optional ByVal bookReadOnly As Boolean = True, Optional ByVal bookPassword As String = "", Optional ByVal bookIngoreReadOnly As Boolean = True, _
    Optional ByVal sheetRemoveExistingSheet As Boolean = False, Optional ByVal sheetCreateIfNotFound As Boolean = True) As Variant
Dim tempRange As Excel.Range
Set tempRange = attainExcelRange(workbookFullPath, worksheetName, rangeAddress, rangeLocation, _
    bookUpdateLink, bookReadOnly, bookPassword, bookIngoreReadOnly, _
    sheetRemoveExistingSheet, sheetCreateIfNotFound)
Dim i As Long
i = tempRange.Cells.Count
If i = 1 Then
    Dim tempArray() As Variant
    ReDim tempArray(1 To 1, 1 To 1) As Variant
    tempArray(1, 1) = tempRange.Cells(1, 1).Value
    attainArrayFromRange = tempArray
End If
If i > 1 Then
    attainArrayFromRange = tempRange.Value
End If
End Function

