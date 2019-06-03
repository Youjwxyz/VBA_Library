Attribute VB_Name = "Fun_TableReader"
Option Explicit


'attainDataFromTable

Private Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("Workbooker", "Worksheeter", "TableReader")
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

Function attainDataFromTable(ByVal workbookFullPath As String, ByVal worksheetName As String, _
    ByRef titleAndDataRefArray As Variant, Optional ByVal rowRepeat As Long = 0, Optional ByVal colRepeat As Long = 0, _
    Optional ByVal bookUpdateLink As Boolean = False, Optional ByVal bookReadOnly As Boolean = True, _
    Optional ByVal bookPassword As String = "", Optional ByVal bookIngoreReadOnly As Boolean = True) As Variant
If workbookFullPath <> "" And worksheetName <> "" And TypeName(titleAndDataRefArray) = "Variant()" Then
    Dim WB As New Workbooker
    WB.updateLink = bookUpdateLink
    WB.readOnly = bookReadOnly
    WB.password = bookPassword
    WB.ignoreReadOnly = bookIngoreReadOnly
    WB.workbookFullPath = workbookFullPath
    
    Dim WS As New Worksheeter
    Set WS.bookObject = WB.workbookObject
    WS.createIfNotFound = False
    WS.removeExistingSheet = False
    WS.worksheetName = worksheetName
    
    Dim TR As New TableReader
    Set TR.sourceSheet = WS.worksheetObject
    TR.titleAndDataRefArray = titleAndDataRefArray
    TR.rowRepeat = rowRepeat
    TR.colRepeat = colRepeat
    TR.extractTable
    attainDataFromTable = TR.tableData
End If
End Function


