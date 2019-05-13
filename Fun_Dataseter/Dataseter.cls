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

Private sourceFile As String
Private conStr As String
Public sqlStr As String
Private cnn As Variant
Private dataSet As Variant
Private outputSheetName As String

Public Property Let outputWorksheetName(ByVal worksheetName As String)
If worksheetName <> "" Then
    outputSheetName = worksheetName
    Application.DisplayAlerts = False
    Dim one As Excel.Worksheet
    For Each one In ThisWorkbook.Worksheets
        If one.Name = outputSheetName Then
            one.Delete
            Exit For
        End If
    Next one
    Set one = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    one.Name = outputSheetName
    Set one = Nothing
    Application.DisplayAlerts = True
Else
    Call resetClass
End If
End Property

Public Property Let sourceFileFullName(ByVal fileFullName As String)
If fileFullName <> "" Then
    sourceFile = fileFullName
    Dim i As Long
    i = InStrRev(sourceFile, ".")
    If i > 0 Then
        Dim fileExtention As String
        fileExtention = UCase(Mid(sourceFile, i + 1))
        Select Case fileExtention
        Case "XLSX", "XLSM", "XLS", "XLM"
            conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';" & _
                "Data Source=" & sourceFile & ";"
'        HDR=Yes means header in source range, which could be "Select * from [Sheet4$E10:G15]"
'        IMEX=1 tells the driver to always read "intermixed" data columns as text
        Case "CSV"
            i = InStrRev(sourceFile, "\", i - 1)
            conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Extended Properties='Text;HDR=Yes;FMT=Delimited';" & _
                "Data Source=" & Left(sourceFile, i - 1) & ";"
'        TSV, Tab delimited files require a schema.ini file to be accepted.
        Case "ACCDB", "MDB"
            conStr = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                "Persist Security Info=False;" & _
                "Data Source=" & sourceFile & ";"
'        If password is required, use "Jet OLEDB:Database Password=MyDbPassword;"
        Case Else
            conStr = ""
        End Select
    Else
        Call resetClass
    End If
Else
    Call resetClass
End If
End Property

Public Sub openADODBConnection()
If conStr <> "" Then
    Set cnn = CreateObject("ADODB.Connection")
    cnn.Open conStr
Else
    Call resetClass
End If
End Sub

Public Sub closeADODBConnection()
If TypeName(cnn) = "Connection" Then
    cnn.Close
    Set cnn = Nothing
    cnn = Empty
Else
    Call resetClass
End If
End Sub

Public Sub runSQLToAttainDataset()
If TypeName(cnn) = "Connection" And sqlStr <> "" Then
    Set dataSet = cnn.Execute(sqlStr)
    If dataSet.Fields.Count = 0 Then
        Set dataSet = Nothing
        dataSet = Empty
    End If
Else
    Call resetClass
End If
End Sub

Public Sub outputRecordSetToWorksheet()
If TypeName(dataSet) = "Recordset" And outputSheetName <> "" Then
    With ThisWorkbook.Worksheets(outputSheetName)
        Dim i As Long
        Dim j As Long
        i = dataSet.Fields.Count
        For j = 1 To i
            .Cells(1, j).Value = dataSet.Fields(j - 1).Name
        Next j
        .Range("A2").CopyFromRecordset dataSet
    End With
Else
    Call resetClass
End If
End Sub

Public Function outputRecordSetToArray() As Variant
If TypeName(dataSet) = "Recordset" Then
    Dim tempArr As Variant
    tempArr = dataSet.getrows()
    Dim i As Long
    Dim j As Long
    i = UBound(tempArr, 2) + 1
    j = UBound(tempArr, 1) + 1
    Dim FinalArr() As Variant
    ReDim FinalArr(1 To i, 1 To j) As Variant
    Dim m As Long
    Dim n As Long
    For m = 1 To i
        For n = 1 To j
            FinalArr(m, n) = tempArr(n - 1, m - 1)
        Next n
    Next m
    Erase tempArr
    outputRecordSetToArray = FinalArr
    Erase FinalArr
Else
    Call resetClass
End If
End Function

Private Sub resetClass()
sourceFile = ""
conStr = ""
sqlStr = ""
If TypeName(cnn) = "Connection" Then
    cnn.Close
    Set cnn = Nothing
End If
cnn = Empty
If TypeName(dataSet) = "Recordset" Then
    Set dataSet = Nothing
End If
dataSet = Empty
outputSheetName = ""
End Sub
