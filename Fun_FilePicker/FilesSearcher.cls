VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilesSearcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public folder As String
Public extension As String
Public files As Variant

Public Sub searchFile()
Dim fileName As String
fileName = Dir(folder & "\*." & extension)
Dim counter As Long
counter = -1
Dim Temp() As String
Do Until fileName = ""
    counter = counter + 1
    ReDim Preserve Temp(counter) As String
    Temp(counter) = folder & "\" & fileName
    fileName = Dir
Loop
files = Temp
Erase Temp
End Sub

Private Sub resetClass()
folder = ""
extension = ""
files = Empty
End Sub

