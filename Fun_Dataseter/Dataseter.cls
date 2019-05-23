VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dataseter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sqlStrings As Variant
Private sourceFile As String
Private dataArray As Variant

Public Property Let sourceFileFullName(ByVal fileFullName As String)
fileFullName = Trim(fileFullName)
Dim conStr As String
conStr = attainConnectionString(fileFullName)
If conStr <> "" And (TypeName(sqlStrings) = "String()" Or TypeName(sqlStrings) = "String") Then
    sourceFile = fileFullName
    Dim sourceCnn As Variant
    Set sourceCnn = CreateObject("ADODB.Connection")
    sourceCnn.Open conStr
    
    Dim dataSet As Variant
    If TypeName(sqlStrings) = "String()" Then
        Dim one As Variant
        For Each one In sqlStrings
            Set dataSet = sourceCnn.Execute(one)
        Next one
    Else
        Set dataSet = sourceCnn.Execute(sqlStrings)
    End If
    If dataSet.Fields.Count > 0 Then
        dataArray = storeRecordSetInArray(dataSet)
    End If
    Set dataSet = Nothing
    dataSet = Empty
    
    sourceCnn.Close
    Set sourceCnn = Nothing
    sourceCnn = Empty
Else
    Call resetClass
End If
End Property

Property Get resultArray() As Variant
If TypeName(dataArray) = "Variant()" Then
    resultArray = dataArray
Else
    MsgBox "Error in VBA codes, please contact alex.you@walmart.com for help."
    Stop
    Call resetClass
End If
End Property

Private Sub resetClass()
sqlStrings = Empty
sourceFile = ""
dataArray = ""
End Sub

Private Function attainConnectionString(ByVal fileFullName As String) As String
fileFullName = Trim(fileFullName)
Dim i As Long
i = InStrRev(fileFullName, ".")
If i > 0 Then
    Dim conStr As String
    Dim fileExtention As String
    fileExtention = UCase(Mid(fileFullName, i + 1))
    Select Case fileExtention
    Case "XLSX", "XLSM", "XLS", "XLM"
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';" & _
            "Data Source=" & fileFullName & ";"
'        HDR=Yes means header in source range, which could be "Select * from [Sheet4$E10:G15]"
'        IMEX=1 tells the driver to always read "intermixed" data columns as text
    Case "CSV"
        i = InStrRev(sourceFile, "\", i - 1)
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Extended Properties='Text;HDR=Yes;FMT=Delimited';" & _
            "Data Source=" & Left(fileFullName, i - 1) & ";"
'        TSV, Tab delimited files require a schema.ini file to be accepted.
    Case "ACCDB", "MDB"
        conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
            "Persist Security Info=False;" & _
            "Data Source=" & fileFullName & ";"
'        If password is required, use "Jet OLEDB:Database Password=MyDbPassword;"
    End Select
    attainConnectionString = conStr
End If
End Function

Private Function storeRecordSetInArray(ByRef dataSet As Variant) As Variant
If TypeName(dataSet) = "Recordset" Then
    Dim tempArr As Variant
    tempArr = dataSet.getrows()
    Dim i As Long
    Dim j As Long
    i = UBound(tempArr, 2) + 2
    j = UBound(tempArr, 1) + 1
    Dim FinalArr() As Variant
    ReDim FinalArr(1 To i, 1 To j) As Variant
    Dim m As Long
    Dim n As Long
    For m = 1 To i
        For n = 1 To j
            If m = 1 Then
                FinalArr(m, n) = dataSet.Fields(n - 1).Name
            Else
                FinalArr(m, n) = tempArr(n - 1, m - 2)
            End If
        Next n
    Next m
    Erase tempArr
    storeRecordSetInArray = FinalArr
    Erase FinalArr
End If
End Function



