VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrPather"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private file_Full_Path As String
Private path_Only As String
Private file_Name_With_Extension As String
Private file_Name_Without_Extension As String
Private extension_Name As String

Property Let fileFullPath(ByVal stringText As String)
If stringText <> "" Then
    Dim i As Long
    i = InStrRev(stringText, "\")
    Dim j As Long
    j = InStrRev(stringText, ".")
    Dim k As Long
    k = Len(stringText)
    If i > 0 And j > 0 And i + 1 < j And j < k Then
        file_Full_Path = stringText
        path_Only = Left(stringText, i - 1)
        file_Name_With_Extension = Mid(stringText, i + 1)
        file_Name_Without_Extension = Mid(stringText, i + 1, j - i - 1)
        extension_Name = Mid(stringText, j + 1)
    Else
        Call resetClass("Source file full path looks incorrect.")
    End If
Else
    Call resetClass("Source file full path is not provided.")
End If
End Property

Property Get fileFullPath() As String
If file_Full_Path <> "" Then
    fileFullPath = file_Full_Path
Else
    Call resetClass
    fileFullPath = ""
End If
End Property

Property Get path() As String
If path_Only <> "" Then
    path = path_Only
Else
    Call resetClass
    path = ""
End If
End Property

Property Get fileNameWithExt() As String
If file_Name_With_Extension <> "" Then
    fileNameWithExt = file_Name_With_Extension
Else
    Call resetClass
    fileNameWithExt = ""
End If
End Property

Property Get fileNameWithoutExt() As String
If file_Name_Without_Extension <> "" Then
    fileNameWithoutExt = file_Name_Without_Extension
Else
    Call resetClass
    fileNameWithoutExt = ""
End If
End Property

Property Get extensionName() As String
If extension_Name <> "" Then
    extensionName = extension_Name
Else
    Call resetClass
    extensionName = ""
End If
End Property

Private Function resetClass(Optional ByVal errorMessage As String = "Unknown Error")
file_Full_Path = ""
path_Only = ""
file_Name_With_Extension = ""
file_Name_Without_Extension = ""
extension_Name = ""
MsgBox errorMessage
End Function
