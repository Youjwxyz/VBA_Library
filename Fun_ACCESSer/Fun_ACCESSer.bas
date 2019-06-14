Attribute VB_Name = "Fun_ACCESSer"
Option Explicit

'importDataToNewTable
'importDataToExistingTable
'retainTableFieldsInfo
'runSQLScript
'runSQLScriptToAttainData

Private Sub importFunctionRelatedClass()
Dim classArr As Variant
classArr = Array("ACCESSer")
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

Public Function retainTableFieldsInfo(Optional ByVal tableName As String = "TEMP", _
    Optional ByVal fileFullPath As String = "TEMP.accdb") As Variant
Dim AO As New ACCESSer
AO.tableName = tableName
If fileFullPath = "TEMP.accdb" Then
    AO.filePath = ThisWorkbook.Path & "\" & fileFullPath
End If
retainTableFieldsInfo = AO.fieldInfo
Set AO = Nothing
End Function

Public Function importDataToNewTable(ByRef sourceDataArray As Variant, ByRef fieldsInfo As Variant, _
    Optional ByVal tableName As String = "TEMP", Optional ByVal fileFullPath As String = "TEMP.accdb")
Dim AO As New ACCESSer
AO.tableName = tableName
If fileFullPath = "TEMP.accdb" Then
    AO.filePath = ThisWorkbook.Path & "\" & fileFullPath
Else
    AO.filePath = fileFullPath
End If
AO.fieldInfo = fieldsInfo
AO.transactionDataArray = sourceDataArray
AO.importDataToNewTable
Set AO = Nothing
End Function

Public Function importDataToExistingTable(ByRef sourceDataArray As Variant, _
    Optional ByVal tableName As String = "TEMP", Optional ByVal fileFullPath As String = "TEMP.accdb")
Dim AO As New ACCESSer
AO.tableName = tableName
If fileFullPath = "TEMP.accdb" Then
    AO.filePath = ThisWorkbook.Path & "\" & fileFullPath
Else
    AO.filePath = fileFullPath
End If
AO.transactionDataArray = sourceDataArray
AO.appendDataToExistingTable
Set AO = Nothing
End Function

Public Function runSQLScript(ByVal sqlString As String, Optional ByVal tableName As String = "TEMP", _
    Optional ByVal fileFullPath As String = "TEMP.accdb")
Dim AO As New ACCESSer
AO.tableName = tableName
If fileFullPath = "TEMP.accdb" Then
    AO.filePath = ThisWorkbook.Path & "\" & fileFullPath
Else
    AO.filePath = fileFullPath
End If
AO.sqlScripts = sqlString
AO.runSQLScripts
Set AO = Nothing
End Function

Public Function runSQLScriptToAttainData(ByVal sqlString As String, Optional ByVal tableName As String = "TEMP", _
    Optional ByVal fileFullPath As String = "TEMP.accdb") As Variant
Dim AO As New ACCESSer
AO.tableName = tableName
If fileFullPath = "TEMP.accdb" Then
    AO.filePath = ThisWorkbook.Path & "\" & fileFullPath
Else
    AO.filePath = fileFullPath
End If
AO.sqlScripts = sqlString
AO.runSQLScriptsToAttainData
runSQLScriptToAttainData = AO.transactionDataArray
Set AO = Nothing
End Function
