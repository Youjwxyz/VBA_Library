Attribute VB_Name = "Fun_Dataseter"
Option Explicit

'attainArrayFromSQLScripts

Private Sub importFunctionRelatedClass()
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


Public Function attainArrayFromSQLScripts(ByRef SQLScripts As Variant, _
    Optional ByVal sourceFile As String) As Variant
If sourceFile = "" Then
    sourceFile = ThisWorkbook.FullName
End If
If verifySQLStripts(SQLScripts) = True Then
    Dim DS As New Dataseter
    DS.sqlStrings = SQLScripts
    DS.sourceFileFullName = sourceFile
    attainArrayFromSQLScripts = DS.resultArray
End If
End Function

Private Function verifySQLStripts(ByRef SQLScripts As Variant) As Boolean
Dim indicator As Boolean
Select Case TypeName(SQLScripts)
Case "String"
    If SQLScripts <> "" Then
        indicator = True
    End If
Case "String()"
    If isEmptyArray(SQLScripts) = False Then
        indicator = True
        Dim one As Variant
        For Each one In SQLScripts
            If one = "" Then
                indicator = False
                Exit For
            End If
        Next one
    End If
End Select
verifySQLStripts = indicator
End Function

Private Function isEmptyArray(ByVal sourceArray As Variant) As Boolean
On Error Resume Next
If UBound(sourceArray, 1) < LBound(sourceArray, 1) Then
    isEmptyArray = True
End If
End Function



