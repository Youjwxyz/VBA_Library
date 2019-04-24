Attribute VB_Name = "Fun_Dataseter"
Option Explicit

'runSQLAndSaveInternalData
'runSQLAndSaveExtenalData
'runSQLAndSaveDataInWorksheet
'runSQLsAndSaveInternalData
'runSQLsAndSaveExtenalData
'runSQLsAndSaveDataInWorksheet

Public Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("Dataseter")
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


Public Sub runSQLAndSaveInternalData(ByVal sqlStr As String)
Call runSQLAndSaveData(sqlStr)
End Sub

Public Sub runSQLAndSaveExtenalData(ByVal sourceFile As String, _
    ByVal sqlStr As String)
Call runSQLAndSaveData(sqlStr, sourceFile)
End Sub

Public Sub runSQLAndSaveDataInWorksheet(ByVal sourceFile As String, _
    ByVal sqlStr As String, ByVal targetSheetName As String)
Call runSQLAndSaveData(sqlStr, sourceFile, targetSheetName)
End Sub

Public Sub runSQLsAndSaveInternalData(ByRef sqlStr() As String)
Call runSQLAndSaveData(sqlStr)
End Sub

Public Sub runSQLsAndSaveExtenalData(ByVal sourceFile As String, _
    ByRef sqlStr() As String)
Call runSQLAndSaveData(sqlStr, sourceFile)
End Sub

Public Sub runSQLsAndSaveDataInWorksheet(ByVal sourceFile As String, _
    ByRef sqlStr() As String, ByVal targetSheetName As String)
Call runSQLAndSaveData(sqlStr, sourceFile, targetSheetName)
End Sub

Private Sub runSQLAndSaveData(ByVal sqlStr As Variant, _
    Optional ByVal sourceFile As String, Optional ByVal targetSheetName As String)
Dim DS As New Dataseter
If sourceFile = "" Then
    DS.sourceFileFullName = ThisWorkbook.FullName
Else
    DS.sourceFileFullName = sourceFile
End If
DS.openADODBConnection
If TypeName(sqlStr) = "String" Then
    DS.sqlStr = sqlStr
    DS.runSQLToAttainDataset
Else
    Dim one As Variant
    For Each one In sqlStr
        DS.sqlStr = one
        DS.runSQLToAttainDataset
    Next one
End If
If targetSheetName = "" Then
    DS.outputWorksheetName = "TempRS"
Else
    DS.outputWorksheetName = targetSheetName
End If
DS.outputRecordSet
DS.closeADODBConnection
Set DS = Nothing
End Sub
