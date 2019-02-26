VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TaskScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private TSShell As Object
Private TSName As String
Private TSApp As String
Private TSFile As String
Private TSResult As Boolean

Private Sub Class_Initialize()
Set TSShell = CreateObject("WSCript.shell")
End Sub

Public Property Let TaskName(ByVal ScheduleTaskName As String)
TSName = ScheduleTaskName
End Property

Public Property Let AppName(ByVal ApplicationName As String)
TSApp = ApplicationName
End Property

Public Property Let ArgFileName(ByVal ArgumentFileName As String)
TSFile = ArgumentFileName
End Property

Public Property Get LastRunResult() As Boolean
LastRunResult = TSResult
End Property

Public Sub CreateScheduleTask(ByVal TaskDate As Date, ByVal TaskTime As Date)
Dim CmdStr As String
Dim Result As Long
If TSName <> "" And TSApp <> "" Then
    CmdStr = "schtasks /create "
    CmdStr = CmdStr & "/tn """ & TSName & """ "
    If TSFile = "" Then
        CmdStr = CmdStr & "/tr ""\""" & TSApp & "\"""" "
    Else
        CmdStr = CmdStr & "/tr ""\""" & TSApp & "\"" -file \""" & TSFile & "\"""" "
    End If
    CmdStr = CmdStr & "/sc once "
    CmdStr = CmdStr & "/sd " & Format(TaskDate, "MM/dd/yyyy") & " "
    CmdStr = CmdStr & "/st " & Format(TaskTime, "HH:mm") & " "
    CmdStr = CmdStr & "/f"
'    Debug.Print CmdStr
    Result = TSShell.Run(CmdStr, 0, True)
    TSResult = Not CBool(Result)
End If
End Sub

Public Sub DeleteScheduleTask()
Dim CmdStr As String
Dim Result As Long
If TSName <> "" Then
    CmdStr = "schtasks /delete "
    CmdStr = CmdStr & "/tn """ & TSName & """ "
    CmdStr = CmdStr & "/f"
'    Debug.Print CmdStr
    Result = TSShell.Run(CmdStr, 0, True)
    TSResult = Not CBool(Result)
End If
End Sub
