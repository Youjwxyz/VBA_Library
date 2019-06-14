VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACCESSer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public filePath As String
Public tableName As String
Public sqlScripts As String
Private fieldsInfoArray As Variant
Private dataArray As Variant

Public Sub importDataToNewTable()
If filePath <> "" And tableName <> "" And TypeName(fieldsInfoArray) = "Variant()" _
    And TypeName(dataArray) = "Variant()" Then
    If verifyTableName(filePath, tableName) = True Then
        If dropTable(filePath, tableName) = False Then
            Call resetClass("Error In Droping Existing Table, Before Data Importing.")
        End If
    End If
    If createTable(filePath, tableName, fieldsInfoArray) = False Then
        Call resetClass("Error In Creating New Table, Before Data Importing.")
    End If
    Dim sourceFieldInfo As Variant
    sourceFieldInfo = fieldInfo
    If importData(filePath, tableName, sourceFieldInfo, dataArray) = False Then
        Call resetClass("Error In Importing Data Into New Table")
    End If
    Erase sourceFieldInfo
Else
    Call resetClass("File Path, Table Name etc Source Data Error.")
End If
End Sub

Public Sub appendDataToExistingTable()
If filePath <> "" And tableName <> "" And TypeName(dataArray) = "Variant()" Then
    If verifyTableName(filePath, tableName) = False Then
        If createTable(filePath, tableName, fieldsInfoArray) = False Then
            Call resetClass("Error In Creating New Table, Before Data Importing.")
        End If
    End If
    Dim sourceFieldInfo As Variant
    sourceFieldInfo = fieldInfo
    If importData(filePath, tableName, sourceFieldInfo, dataArray) = False Then
        Call resetClass("Error In Importing Data Into Existing Table")
    End If
    Erase sourceFieldInfo
Else
    Call resetClass("File Path, Table Name etc Source Data Error.")
End If
End Sub

Public Sub runSQLScripts()
If filePath <> "" And tableName <> "" And sqlScripts <> "" Then
    If verifyTableName(filePath, tableName) = False Then
        Call resetClass("Target File/Table Is Not Found.")
    End If
    dataArray = runScripts(filePath, tableName, sqlScripts, False)
Else
    Call resetClass("File Path, Table Name etc Source Data Error.")
End If
End Sub

Public Sub runSQLScriptsToAttainData()
If filePath <> "" And tableName <> "" And sqlScripts <> "" Then
    If verifyTableName(filePath, tableName) = False Then
        Call resetClass("Target File/Table Is Not Found.")
    End If
    Dim tempArray As Variant
    tempArray = runScripts(filePath, tableName, sqlScripts, True)
    If TypeName(tempArray) = "Variant()" Then
        dataArray = tempArray
        Erase tempArray
    Else
        Call resetClass("Error In Running SQL Scripts To Attain Data.")
    End If
Else
    Call resetClass("File Path, Table Name etc Source Data Error.")
End If
End Sub


Property Let transactionDataArray(ByRef sourceDataArray As Variant)
'sourceDataArray should be with title row
If TypeName(sourceDataArray) = "Variant()" Then
    If UBound(sourceDataArray, 2) >= 1 And UBound(sourceDataArray, 1) >= 1 Then
        dataArray = sourceDataArray
    Else
        Call resetClass("Source Data Array Is Empty.")
    End If
Else
    Call resetClass("Source Data Is Not Of 'Variant()' Type.")
End If
End Property

Property Get transactionDataArray() As Variant
'dataArray will be with title row
If TypeName(dataArray) = "Variant()" Then
    If UBound(dataArray, 2) >= 1 And UBound(dataArray, 1) >= 1 Then
        transactionDataArray = dataArray
    Else
        Call resetClass("Data Array Is Empty.")
    End If
Else
    Call resetClass("Data Is Not Of 'Variant()' Type.")
End If
End Property

Property Let fieldInfo(ByRef sourceFieldInfo As Variant)
'sourceFieldInfo is without title row
'FieldName, DataType(VARCHAR,MEMO,DECIMAL,LONG,CURRENCY,DATETIME,YESNO,COUNTER), PrimaryKey, NotNullField
If TypeName(sourceFieldInfo) = "Variant()" Then
    If UBound(sourceFieldInfo, 2) = 4 And UBound(sourceFieldInfo, 1) >= 1 Then
        Dim i As Long
        Dim j As Long
        i = UBound(sourceFieldInfo, 1)
        For j = 1 To i
            sourceFieldInfo(j, 1) = cleanFieldName(sourceFieldInfo(j, 1))
        Next j
        fieldsInfoArray = sourceFieldInfo
    Else
        Call resetClass("Source Field Array Is Empty Or Non-4-Columns Array.")
    End If
Else
    Call resetClass("Source Data Is Not Of 'Variant()' Type.")
End If
End Property

Property Get fieldInfo() As Variant
'fieldsInfoArray is without title row
If verifyTableName(filePath, tableName) = True Then
    Dim conStr As String
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Persist Security Info=False;Data Source=" & filePath & ";"
    Dim sourceCnn As Variant
    Set sourceCnn = CreateObject("ADODB.Connection")
    sourceCnn.Open conStr
    
    Dim tempArray() As Variant
    Dim counter As Long
    Dim rs As Variant
    ReDim tempArray(1 To 1000, 1 To 4)
    Set rs = sourceCnn.OpenSchema(4) 'adSchemaColumns
    Do Until rs.EOF
        If rs.Fields("TABLE_NAME") = tableName Then
            counter = counter + 1
            tempArray(counter, 1) = rs.Fields("COLUMN_NAME")
            If rs.Fields("DATA_TYPE") = 130 And rs.Fields("CHARACTER_MAXIMUM_LENGTH") <> 255 Then
                tempArray(counter, 2) = convertDataType(9999)
            Else
                tempArray(counter, 2) = convertDataType(rs.Fields("DATA_TYPE"))
            End If
            tempArray(counter, 3) = False
            tempArray(counter, 4) = Not rs.Fields("IS_NULLABLE")
        End If
        rs.movenext
    Loop
    rs.Close
    Set rs = sourceCnn.OpenSchema(12) 'adSchemaIndexes
    If counter > 0 Then
        Dim Final_Array() As Variant
        ReDim Final_Array(1 To counter, 1 To 4) As Variant
        Dim i As Long
        For i = 1 To counter
            Final_Array(i, 1) = tempArray(i, 1)
            Final_Array(i, 2) = tempArray(i, 2)
            Final_Array(i, 3) = tempArray(i, 3)
            rs.movefirst
            Do Until rs.EOF
                If rs.Fields("TABLE_NAME") = tableName And rs.Fields("COLUMN_NAME") = tempArray(i, 1) Then
                    Final_Array(i, 3) = rs.Fields("PRIMARY_KEY")
                    Exit Do
                End If
                rs.movenext
            Loop
            Final_Array(i, 4) = tempArray(i, 4)
        Next i
        fieldsInfoArray = Final_Array
        Erase Final_Array
        Erase tempArray
    Else
        Call resetClass("Target Table Is Not Found.")
    End If
    rs.Close
    Set rs = Nothing

    sourceCnn.Close
    Set sourceCnn = Nothing
Else
    Call resetClass("Target File/Table Is Not Found.")
End If
fieldInfo = fieldsInfoArray
End Property

Private Sub resetClass(Optional ByVal errorMessage As String = "Unknown Error")
filePath = ""
tableName = ""
sqlScripts = ""
fieldsInfoArray = Empty
dataArray = Empty
MsgBox errorMessage
End Sub


Private Function runScripts(ByVal fileFullPaht As String, ByVal accessTableName As String, _
    ByRef sourceSQLScripts As String, ByVal returnDataset As Boolean) As Variant
Dim tempArr As Variant
If fileFullPaht <> "" And accessTableName <> "" And sourceSQLScripts <> "" Then
    If verifyTableName(fileFullPaht, accessTableName) = True Then

        Dim conStr As String
            conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Persist Security Info=False;Data Source=" & fileFullPaht & ";"
        Dim sourceCnn As Variant
        Set sourceCnn = CreateObject("ADODB.Connection")
        sourceCnn.Open conStr
    
        Dim rs As Variant
        Set rs = CreateObject("ADODB.RecordSet")
        rs.Open sourceSQLScripts, sourceCnn, 1  'AdOpenKeyset
        If returnDataset = True Then
            tempArr = convertRecordSet(rs)
            rs.Close
        End If
        Set rs = Nothing
        
        sourceCnn.Close
        Set sourceCnn = Nothing
    Else
        tempArr = Empty
    End If
Else
    tempArr = Empty
End If
runScripts = tempArr
If TypeName(tempArr) = "Variants()" Then
    Erase tempArr
End If
End Function

Private Function convertRecordSet(ByRef sourceRecordSet As Variant) As Variant
If TypeName(sourceRecordSet) = "Recordset" Then
    Dim recordCount As Long
    Dim fieldCount As Long
    Dim counter As Long
    Dim j As Long
    recordCount = sourceRecordSet.recordCount
    fieldCount = sourceRecordSet.Fields.Count
    Dim tempArr() As Variant
    ReDim tempArr(1 To recordCount + 1, 1 To fieldCount) As Variant
    counter = 1
    For j = 1 To fieldCount
        tempArr(counter, j) = sourceRecordSet.Fields(j - 1).Name
    Next j
    sourceRecordSet.movefirst
    Do Until sourceRecordSet.EOF
        counter = counter + 1
        For j = 1 To fieldCount
            tempArr(counter, j) = sourceRecordSet.Fields(j - 1).Value
        Next j
        sourceRecordSet.movenext
    Loop
    convertRecordSet = tempArr
    Erase tempArr
End If
End Function

Private Function importData(ByVal fileFullPaht As String, ByVal accessTableName As String, _
    ByRef sourceFieldInfo As Variant, ByRef sourceDataArray As Variant) As Boolean
Dim indicator As Boolean
If fileFullPaht <> "" And accessTableName <> "" And TypeName(sourceFieldInfo) = "Variant()" _
    And TypeName(sourceDataArray) = "Variant()" Then
    If verifyTableName(fileFullPaht, accessTableName) = True Then

        Dim conStr As String
            conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Persist Security Info=False;Data Source=" & fileFullPaht & ";"
        Dim sourceCnn As Variant
        Set sourceCnn = CreateObject("ADODB.Connection")
        sourceCnn.Open conStr
    
        Dim SQL_String As String
        Dim Field_String As String
        Dim M_Value_String As String
        Dim maxRow As Long
        Dim maxCol As Long
        Dim i As Long
        Dim j As Long
        maxRow = UBound(sourceDataArray, 1)
        maxCol = UBound(sourceDataArray, 2)
        For j = 1 To maxCol
            sourceDataArray(1, j) = cleanFieldName(sourceDataArray(1, j))
        Next j
        For i = 2 To maxRow
            Field_String = "( "
            M_Value_String = "values ( "
            For j = 1 To maxCol
                Field_String = Field_String & sourceDataArray(1, j) & ", "
                Select Case sourceFieldInfo(locateFieldIndex(sourceDataArray(1, j), sourceFieldInfo), 2)
                Case "LONG/COUNTER", "DECIMAL", "CURRENCY", "YESNO"
                    If sourceDataArray(i, j) = "" Then
                        M_Value_String = M_Value_String & "0, "
                    Else
                        M_Value_String = M_Value_String & sourceDataArray(i, j) & ", "
                    End If
                Case "VARCHAR", "MEMO"
                    M_Value_String = M_Value_String & "'" & Replace(sourceDataArray(i, j), "'", "''") & "', "
                Case "DATETIME"
                    If sourceDataArray(i, j) = "" Then
                        M_Value_String = M_Value_String & "Null, "
                    Else
                        M_Value_String = M_Value_String & "#" & sourceDataArray(i, j) & "#, "
                    End If
                End Select
            Next j
            Field_String = Left(Field_String, Len(Field_String) - 2) & ") "
            M_Value_String = Left(M_Value_String, Len(M_Value_String) - 2) & ") "
            SQL_String = "Insert into " & accessTableName & " " & Field_String & M_Value_String
            sourceCnn.Execute SQL_String
        Next i
        
        sourceCnn.Close
        Set sourceCnn = Nothing
        indicator = True
    Else
        indicator = False
    End If
Else
    indicator = False
End If
importData = indicator
End Function


Private Function locateFieldIndex(ByVal fieldName As String, ByRef sourceFieldInfo As Variant) As Long
Dim fieldIndex As Long
If fieldName <> "" And TypeName(sourceFieldInfo) = "Variant()" Then
    Dim i As Long
    Dim j As Long
    i = UBound(sourceFieldInfo, 1)
    For j = 1 To i
        If sourceFieldInfo(j, 1) = fieldName Then
            fieldIndex = j
            Exit For
        End If
    Next j
Else
    fieldIndex = 0
End If
locateFieldIndex = fieldIndex
End Function

Private Function verifyTableName(ByVal fileFullPaht As String, ByVal accessTableName As String) As Boolean
Dim indicator As Boolean
If fileFullPaht <> "" And createAccessFile(fileFullPaht) = True _
    And accessTableName <> "" Then
    Dim conStr As String
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Persist Security Info=False;Data Source=" & fileFullPaht & ";"
    Dim sourceCnn As Variant
    Set sourceCnn = CreateObject("ADODB.Connection")
    sourceCnn.Open conStr
    
    Dim rs As Variant
    Set rs = sourceCnn.OpenSchema(20) 'adSchemaTables
    Do Until rs.EOF
        If rs.Fields("TABLE_TYPE") = "TABLE" Then
            If UCase(rs.Fields("TABLE_NAME")) = UCase(accessTableName) Then
                indicator = True
                Exit Do
            End If
        End If
        rs.movenext
    Loop
    rs.Close
    Set rs = Nothing
    
    sourceCnn.Close
    Set sourceCnn = Nothing
Else
    indicator = False
End If
verifyTableName = indicator
End Function

Private Function dropTable(ByVal fileFullPaht As String, ByVal accessTableName As String) As Boolean
Dim indicator As Boolean
If fileFullPaht <> "" And createAccessFile(fileFullPaht) = True _
    And accessTableName <> "" Then
    Dim conStr As String
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Persist Security Info=False;Data Source=" & fileFullPaht & ";"
    Dim sourceCnn As Variant
    Set sourceCnn = CreateObject("ADODB.Connection")
    sourceCnn.Open conStr
    
    Dim SQL_String As String
    SQL_String = "DROP TABLE " & accessTableName
    sourceCnn.Execute SQL_String
                            
    sourceCnn.Close
    Set sourceCnn = Nothing
    indicator = True
Else
    indicator = False
End If
dropTable = indicator
End Function

Private Function createTable(ByVal fileFullPaht As String, ByVal accessTableName As String, _
    ByRef sourceFieldInfo As Variant) As Boolean
Dim indicator As Boolean
If fileFullPaht <> "" And createAccessFile(fileFullPaht) = True _
    And accessTableName <> "" And TypeName(sourceFieldInfo) = "Variant()" Then
    Dim conStr As String
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Persist Security Info=False;Data Source=" & fileFullPaht & ";"
    Dim sourceCnn As Variant
    Set sourceCnn = CreateObject("ADODB.Connection")
    sourceCnn.Open conStr
    
    Dim SQL_String As String
    Dim PK_String As String
    Dim NN_String As String
    SQL_String = "CREATE TABLE " & accessTableName & " ("
    PK_String = "PRIMARY KEY ("
    Dim i As Long
    Dim j As Long
    i = UBound(sourceFieldInfo, 1)
    For j = 1 To i
        If sourceFieldInfo(j, 4) = True Then
            NN_String = " NOT NULL"
        Else
            NN_String = ""
        End If
        Select Case sourceFieldInfo(j, 2)
        Case "VARCHAR"
            SQL_String = SQL_String & sourceFieldInfo(j, 1) & " " & sourceFieldInfo(j, 2) & "(255)" & NN_String & ", "
        Case "DECIMAL"
            SQL_String = SQL_String & sourceFieldInfo(j, 1) & " " & sourceFieldInfo(j, 2) & "(28,8)" & NN_String & ", "
        Case Else
            SQL_String = SQL_String & sourceFieldInfo(j, 1) & " " & sourceFieldInfo(j, 2) & NN_String & ", "
        End Select
        
        If fieldsInfoArray(j, 3) = True Then
            PK_String = PK_String & sourceFieldInfo(j, 1) & ","
        End If
    Next j
    SQL_String = Left(SQL_String, Len(SQL_String) - 2)
    If PK_String <> "PRIMARY KEY (" Then
        SQL_String = SQL_String & ", " & Left(PK_String, Len(PK_String) - 1) & ")"
    End If
    SQL_String = SQL_String & ")"
    sourceCnn.Execute SQL_String
    
    sourceCnn.Close
    Set sourceCnn = Nothing
    indicator = True
Else
    indicator = False
End If
createTable = indicator
End Function

Private Function createAccessFile(ByVal fileFullPaht As String) As Boolean
Dim indicator As Boolean
If fileFullPaht <> "" Then
    Dim fileObject As Variant
    Set fileObject = CreateObject("Scripting.FileSystemObject")
    If fileObject.FileExists(fileFullPaht) = False Then
        Dim accessApp As Variant
        Set accessApp = CreateObject("Access.Application")
        accessApp.NewCurrentDatabase fileFullPaht, 12  'acNewDatabaseFormatAccess12
        Set accessApp = Nothing
        Call waitFewSecond(2)
    End If
    Set fileObject = Nothing
    indicator = True
Else
    indicator = False
End If
createAccessFile = indicator
End Function

Private Function waitFewSecond(ByVal secoundCount As Long)
Dim sTime As Date
sTime = Now
Do Until DateDiff("s", sTime, Now) > secoundCount
    DoEvents
Loop
End Function

Private Function convertDataType(ByVal dataTypeIndex As Long) As String
Dim typeString As String
'VARCHAR,MEMO,DECIMAL,LONG,CURRENCY,DATETIME,YESNO,COUNTER
Select Case dataTypeIndex
Case 130
    typeString = "VARCHAR"
Case 9999 'Dummy number which is created to identify Memo from Varchar
    typeString = "MEMO"
Case 131
    typeString = "DECIMAL"
Case 3
    typeString = "LONG/COUNTER"
Case 6
    typeString = "CURRENCY"
Case 7
    typeString = "DATETIME"
Case 11
    typeString = "YESNO"
Case Else
    MsgBox "Field Data Type Not Defined."
    convertDataType = "UNKNOWN_" & dataTypeIndex
End Select
convertDataType = typeString
End Function

Private Function cleanFieldName(ByVal sourceFieldName) As String
Dim tempStr As String
If sourceFieldName <> "" Then
    tempStr = sourceFieldName
    tempStr = Replace(tempStr, " ", "_")
    tempStr = Replace(tempStr, "-", "_")
    tempStr = Replace(tempStr, "'", "_")
    tempStr = Replace(tempStr, Chr(10), "_")
Else
    tempStr = ""
End If
cleanFieldName = tempStr
End Function
