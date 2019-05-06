Attribute VB_Name = "Fun_Worksheeter"
Option Explicit

'importInternalWorksheetDataToThisWorkbook
'importInternalWorksheetDataToAnotherWorkbook
'importExternalWorksheetDataToThisWorkbook
'importExternalWorksheetDataToAnotherWorkbook

Public Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("TwoDArrayer", "Workbooker")
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

Sub importInternalWorksheetDataToThisWorkbook(ByVal sourceSheetName As String, _
    ByVal targetSheetName As String)
Call importWorksheetData(ThisWorkbook.FullName, sourceSheetName, False, False, "", False, False, _
    ThisWorkbook.FullName, targetSheetName, False, False, "", False, False)
End Sub

Sub importInternalWorksheetDataToAnotherWorkbook(ByVal sourceSheetName As String, _
    ByVal targetBookFullPath As String, _
    ByVal targetSheetName As String, ByVal targetUpdateLink As Boolean, _
    ByVal targetReadOnly As Boolean, ByVal targetPassword As String, _
    ByVal targetIgnoreReadOnly As Boolean, ByVal closeTargetBook As Boolean)
Call importWorksheetData(ThisWorkbook.FullName, sourceSheetName, False, False, "", False, False, _
    targetBookFullPath, targetSheetName, targetUpdateLink, targetReadOnly, targetPassword, targetIgnoreReadOnly, closeTargetBook)
End Sub

Sub importExternalWorksheetDataToThisWorkbook(ByVal sourceBookFullPath As String, ByVal sourceSheetName As String, _
    ByVal sourceUpdateLink As Boolean, ByVal sourceReadOnly As Boolean, _
    ByVal sourcePassword As String, ByVal sourceIgnoreReadOnly As Boolean, _
    ByVal closeSourceBook As Boolean, ByVal targetSheetName As String)
Call importWorksheetData(sourceBookFullPath, sourceSheetName, sourceUpdateLink, sourceReadOnly, sourcePassword, sourceIgnoreReadOnly, closeSourceBook, _
    ThisWorkbook.FullName, targetSheetName, False, False, "", False, False)
End Sub

Sub importExternalWorksheetDataToAnotherWorkbook(ByVal sourceBookFullPath As String, ByVal sourceSheetName As String, _
    ByVal sourceUpdateLink As Boolean, ByVal sourceReadOnly As Boolean, _
    ByVal sourcePassword As String, ByVal sourceIgnoreReadOnly As Boolean, _
    ByVal closeSourceBook As Boolean, ByVal targetBookFullPath As String, _
    ByVal targetSheetName As String, ByVal targetUpdateLink As Boolean, _
    ByVal targetReadOnly As Boolean, ByVal targetPassword As String, _
    ByVal targetIgnoreReadOnly As Boolean, ByVal closeTargetBook As Boolean)
Call importWorksheetData(sourceBookFullPath, sourceSheetName, sourceUpdateLink, sourceReadOnly, sourcePassword, sourceIgnoreReadOnly, closeSourceBook, _
    targetBookFullPath, targetSheetName, targetUpdateLink, targetReadOnly, targetPassword, targetIgnoreReadOnly, closeTargetBook)
End Sub

Private Sub importWorksheetData(ByVal sourceBookFullPath As String, ByVal sourceSheetName As String, _
    ByVal sourceUpdateLink As Boolean, ByVal sourceReadOnly As Boolean, _
    ByVal sourcePassword As String, ByVal sourceIgnoreReadOnly As Boolean, _
    ByVal closeSourceBook As Boolean, ByVal targetBookFullPath As String, _
    ByVal targetSheetName As String, ByVal targetUpdateLink As Boolean, _
    ByVal targetReadOnly As Boolean, ByVal targetPassword As String, _
    ByVal targetIgnoreReadOnly As Boolean, ByVal closeTargetBook As Boolean)
Dim tempData As Variant
tempData = attainData(sourceBookFullPath, sourceSheetName, sourceUpdateLink, sourceReadOnly, _
    sourcePassword, sourceIgnoreReadOnly, closeSourceBook)
If TypeName(tempData) <> "Empty" Then
    If targetBookFullPath <> "" Then
        Dim WB As New Workbooker
        WB.updateLink = targetUpdateLink
        WB.readOnly = targetReadOnly
        WB.password = targetPassword
        WB.ignoreReadOnly = targetIgnoreReadOnly
        WB.workbookFullPath = targetBookFullPath
        If TypeName(WB.workbookObject) = "Workbook" Then
            Dim tempBook As Excel.Workbook
            Set tempBook = WB.workbookObject
            If targetSheetName <> "" Then
                Dim TDA As New TwoDArrayer
                Set TDA.parentWorkbook = tempBook
                TDA.worksheetName = targetSheetName
                TDA.sourceData = tempData
                TDA.toWorksheet
                If closeTargetBook = True Then
                    Application.DisplayAlerts = False
                        tempBook.Close True
                    Application.DisplayAlerts = True
                End If
                Set TDA = Nothing
            End If
            Set tempBook = Nothing
        End If
        Set WB = Nothing
    End If
    tempData = Empty
End If
End Sub

Private Function attainData(ByVal sourceBookFullPath As String, ByVal sourceSheetName As String, _
    ByVal updateLink As Boolean, ByVal readOnly As Boolean, ByVal password As String, _
    ByVal ignoreReadOnly As Boolean, ByVal closeSourceBook As Boolean) As Variant
If sourceBookFullPath <> "" Then
    Dim WB As New Workbooker
    WB.updateLink = updateLink
    WB.readOnly = readOnly
    WB.password = password
    WB.ignoreReadOnly = ignoreReadOnly
    WB.workbookFullPath = sourceBookFullPath
    If TypeName(WB.workbookObject) = "Workbook" Then
        Dim tempBook As Excel.Workbook
        Set tempBook = WB.workbookObject
        If sourceSheetName <> "" Then
            Dim TDA As New TwoDArrayer
            Set TDA.parentWorkbook = tempBook
            TDA.worksheetName = sourceSheetName
            TDA.fromWorksheet
            attainData = TDA.sourceData
            If closeSourceBook = True Then
                Application.DisplayAlerts = False
                    tempBook.Close False
                Application.DisplayAlerts = True
            End If
            Set TDA = Nothing
        End If
        Set tempBook = Nothing
    End If
    Set WB = Nothing
End If
End Function


