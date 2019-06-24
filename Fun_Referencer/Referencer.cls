VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Referencer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sourceRange As Excel.Range
Public refCols As Variant

Public Sub createReference()
If TypeName(sourceRange) = "Range" And notEmptyArray(refCols) = True Then
    Dim sourceData As Variant
    sourceData = sourceRange.Value
    Dim i As Long
    Dim j As Long
    i = UBound(sourceData, 2)
    Dim tempArray() As Variant
    ReDim tempArray(1 To i, 1 To 4) As Variant
    For j = 1 To i
        tempArray(j, 1) = cleanFieldName(sourceData(1, j))
        tempArray(j, 2) = "VARCHAR"
        tempArray(j, 3) = False
        tempArray(j, 4) = True
    Next j
    Dim tableName As String
    Dim accessFilePath As String
    tableName = "REF"
    accessFilePath = sourceRange.Parent.Parent.path & "\Reference" & Format(Now, "yyyyMMdd") & ".accdb"
    Dim ACC As New ACCESSer
    ACC.tableName = tableName
    ACC.filePath = accessFilePath
    ACC.fieldInfo = tempArray
    ACC.transactionDataArray = sourceData
    ACC.importDataToNewTable
    Erase tempArray
    Erase sourceData
    Set ACC = Nothing

    With sourceRange.Parent
        j = .UsedRange.SpecialCells(xlCellTypeLastCell).Column
        If j < sourceRange.Cells(1, 1).End(xlToRight).Offset(0, 1).Column Then
            j = sourceRange.Cells(1, 1).End(xlToRight).Offset(0, 1).Column
        End If
        .Range(sourceRange.Cells(1, 1).End(xlToRight).Offset(0, 1), _
            sourceRange.Cells(1, j)).EntireColumn.Delete
    End With
    Call retrieveRefData(accessFilePath, tableName)
    Call createDefaultRef
Else
    Call resetClass("Source range or key columns array is not provided.")
End If
End Sub

Private Sub createDefaultRef()
If TypeName(sourceRange) = "Range" And notEmptyArray(refCols) = True Then
    Dim i As Long
    Dim j As Long
    Dim refName As String
    i = UBound(refCols)
    For j = 1 To i
        refName = String(j, "_") & cleanFieldName(sourceRange.Cells(1, refCols(j)).Value)
        sourceRange.Parent.Parent.Names.Add Name:=refName, _
            RefersTo:="=" & sourceRange.Parent.Name & "!" & _
            sourceRange.Cells(1, sourceRange.Columns.Count + 1).Address
    Next j
Else
    Call resetClass("Source range or key columns array is not provided.")
End If
End Sub

Private Sub retrieveRefData(ByVal dbFilePath As String, ByVal tableName As String, _
    Optional ByRef criterias As Variant)
If TypeName(sourceRange) = "Range" And notEmptyArray(refCols) = True And _
    dbFilePath <> "" And tableName <> "" Then
    Dim cirteriaCount As Long
    If notEmptyArray(criterias) = True Then
        cirteriaCount = UBound(criterias) + 1
    End If
    
    If cirteriaCount >= UBound(refCols) + 1 Then
        'This is the condition to exit this self-called sub.
    Else
        Dim newCriterias() As Variant
        ReDim newCriterias(cirteriaCount) As Variant
        Dim targetCell As Excel.Range
        Set targetCell = sourceRange.Cells(1, 1).End(xlToRight).Offset(0, 2)
        Dim i As Long
        For i = 0 To cirteriaCount - 1
            newCriterias(i) = criterias(i)
            If i > 0 Then
                Set targetCell = targetCell.End(xlToRight).Offset(0, 2)
            Else
                Set targetCell = targetCell.Offset(0, 2)
            End If
        Next i
        If targetCell.End(xlDown).Row < sourceRange.Parent.Rows.Count Then
            Set targetCell = targetCell.End(xlDown).Offset(1, 0)
        End If
        
        Dim tempData As Variant
        Dim m As Long
        Dim n As Long
        Dim referenceName As String
        Dim sqlScriptStr As String
        sqlScriptStr = generateSQLScripts(tableName, _
            generateFieldArray(sourceRange.Value, refCols, cirteriaCount + 1), criterias)
            
            Dim ACC As New ACCESSer
            ACC.filePath = dbFilePath
            ACC.tableName = tableName
            ACC.sqlScripts = sqlScriptStr
            ACC.runSQLScriptsToAttainData
            tempData = ACC.transactionDataArray
            Set ACC = Nothing
            
        m = UBound(tempData, 1)
        n = UBound(tempData, 2)
        targetCell.Resize(m, n).Value = tempData
        referenceName = createRefName(tempData)
        targetCell.Parent.Parent.Names.Add Name:=referenceName, _
            RefersTo:="=" & targetCell.Parent.Name & "!" & _
            targetCell.Offset(1, n - 1).Resize(m - 1, 1).Address
        If targetCell.Row > sourceRange.Cells(1, 1).Row Then
            targetCell.Resize(1, n).Delete Shift:=xlUp
        End If
        Set targetCell = Nothing
        
            For i = 2 To m
                newCriterias(cirteriaCount) = tempData(i, n)
                Call retrieveRefData(dbFilePath, tableName, newCriterias)
            Next i
            
        Erase tempData
        Erase newCriterias
    End If
Else
    Call resetClass("Source range, key columns array, or temp Access file is not provided.")
End If
End Sub

Private Sub resetClass(Optional ByVal errorMessage As String = "Unknown Error")
Set sourceRange = Nothing
refCols = Empty
MsgBox errorMessage
End Sub

Private Function createRefName(ByRef sourceData As Variant) As String
Dim tempString As String
If notEmptyArray(sourceData) = True Then
    Dim i As Long
    Dim j As Long
    i = UBound(sourceData, 2)
    For j = 1 To i
        If j < i Then
            tempString = tempString & "_" & sourceData(2, j)
        Else
            tempString = tempString & "_" & sourceData(1, j)
        End If
    Next j
    tempString = Mid(tempString, 2)
End If
createRefName = tempString
End Function

Private Function generateFieldArray(ByRef sourceData As Variant, ByRef refColumns As Variant, _
    ByVal feildsCount As Long) As Variant
Dim tempString As String
If notEmptyArray(sourceData) = True And notEmptyArray(refColumns) = True And feildsCount > 0 Then
    If UBound(refColumns) + 1 >= feildsCount Then
        Dim i As Long
        For i = 1 To feildsCount
            tempString = tempString & "," & cleanFieldName(sourceData(1, refColumns(i - 1)))
        Next i
        tempString = Mid(tempString, 2)
    End If
End If
generateFieldArray = Split(tempString, ",")
End Function

Private Function generateSQLScripts(ByVal tableName As String, ByRef fieldArray As Variant, _
    Optional ByVal criterias As Variant) As String
Dim tempString As String
If tableName <> "" And notEmptyArray(fieldArray) = True Then
    Dim criteriaCount As Long
    If notEmptyArray(criterias) = True Then
        criteriaCount = UBound(criterias) + 1
    Else
        criteriaCount = 0
    End If
    If UBound(fieldArray) = criteriaCount Then
        Dim j As Long
        Dim fieldString As String
        Dim criteriaString As String
        For j = 0 To criteriaCount
            fieldString = fieldString & "," & fieldArray(j)
            If j > 0 Then
                criteriaString = criteriaString & "And " & _
                    fieldArray(j - 1) & " = '" & Replace(criterias(j - 1), "'", "''") & "' "
            End If
        Next j
        fieldString = Mid(fieldString, 2)
        If criteriaCount > 0 Then
            criteriaString = Mid(criteriaString, 5)
            tempString = "Select " & fieldString & " From " & tableName & _
                " Where " & criteriaString & " Group By " & fieldString & " Order By " & fieldString
        Else
            tempString = "Select " & fieldString & " From " & tableName & _
                " Group By " & fieldString & " Order By " & fieldString
        End If
    End If
End If
generateSQLScripts = tempString
End Function

Private Function notEmptyArray(ByRef sourceArray As Variant) As Boolean
On Error Resume Next
Dim Indicator As Boolean
If UBound(sourceArray) < LBound(sourceArray) Then
    Indicator = False
Else
    Indicator = True
End If
notEmptyArray = Indicator
End Function

Private Function cleanFieldName(ByVal sourceFieldName) As String
Dim tempStr As String
If sourceFieldName <> "" Then
    tempStr = sourceFieldName
    tempStr = Replace(tempStr, " ", "_")
    tempStr = Replace(tempStr, "-", "_")
    tempStr = Replace(tempStr, "'", "_")
    tempStr = Replace(tempStr, "/", "_")
    tempStr = Replace(tempStr, "\", "_")
    tempStr = Replace(tempStr, Chr(10), "_")
Else
    tempStr = ""
End If
cleanFieldName = tempStr
End Function
