
'Main Process
Dim MasterFolder
MasterFolder = "C:\Users\A0Y0074\Documents\UiPath"
Call PrepareFile(MasterFolder & "\AdHoc.csv")
Call PrepareFile(MasterFolder & "\Pending_List.csv")
Call PrepareFile(MasterFolder & "\Schedule.csv")
Call PrepareFile(MasterFolder & "\Log.csv")

Dim AdHocArr
AdHocArr = ReadFile(MasterFolder & "\AdHoc.csv")
Dim PendingArr
PendingArr = ReadFile(MasterFolder & "\Pending_List.csv")
Dim ScheduleArr
ScheduleArr = ReadFile(MasterFolder & "\Schedule.csv")

Dim Fso
Set Fso = CreateObject("Scripting.FileSystemObject")
Dim RPAIndicator
RPAIndicator = Fso.FileExists(MasterFolder & "\Working.File")
Set Fso = nothing
RPAIndicator = ProcessSourceData(AdHocArr, RPAIndicator, MasterFolder)
RPAIndicator = ProcessSourceData(PendingArr, RPAIndicator, MasterFolder)
RPAIndicator = ProcessSourceData(ScheduleArr, RPAIndicator, MasterFolder)

'Function To Create Task In TaskScheduler And Log File, Or Add Record To Pending List File.
Function ProcessSourceData(SourceData, RPAIndicator, FilePath)

Dim Fso
Set Fso = CreateObject("Scripting.FileSystemObject")
Dim PendingFile
Set PendingFile = Fso.OpenTextFile(FilePath & "\Pending_List.csv", 8)
Dim LogFile
Set LogFile = Fso.OpenTextFile(FilePath & "\Log.csv", 8)

Dim m
Dim DataCount
DataCount = Ubound(SourceData,1)
For m = 1 to DataCount
	If Now < SourceData(m,1) And SourceData(m,1) - Now <= CDate("00:10:00") Then
		If RPAIndicator = False Then
			Call CreateTask(SourceData(m,0), SourceData(m,1), SourceData(m,2))
			LogFile.WriteLine SourceData(m,0) & "," & cstr(SourceData(m,1)) & "," & SourceData(m,2)
			RPAIndicator = True
		Else
			PendingFile.WriteLine SourceData(m,0) & "," & cstr(SourceData(m,1) + CDate("00:10:00")) & "," & SourceData(m,2)
		End If
	End If
Next
PendingFile.Close
LogFile.Close
Set Fso = Nothing
ProcessSourceData = RPAIndicator

End Function

'Function To Create Task In Task Scheduler.
Function CreateTask(TaskName, TaskTime, TaskFile)

Dim CmdStr
Dim TimeStr
CmdStr = "schtasks /create /sc HOURLY "

TimeStr = AttainTwoDigitNumber(Year(Now))
TimeStr = TimeStr & "_" & AttainTwoDigitNumber(Month(Now))
TimeStr = TimeStr & "_" & Day(Now)
TimeStr = TimeStr & "_" & AttainTwoDigitNumber(Hour(Now))
TimeStr = TimeStr & "_" & AttainTwoDigitNumber(Minute(Now))
CmdStr = CmdStr & "/tn " & """" & TaskName & "_" & TimeStr & """" & " "

TimeStr = AttainTwoDigitNumber(Month(TaskTime))
TimeStr = TimeStr & "/" & AttainTwoDigitNumber(Day(TaskTime))
TimeStr = TimeStr & "/" & Year(TaskTime)
CmdStr = CmdStr & "/sd " & TimeStr & " "

TimeStr = AttainTwoDigitNumber(Hour(TaskTime))
TimeStr = TimeStr & ":" & AttainTwoDigitNumber(Minute(TaskTime))
CmdStr = CmdStr & "/st " & TimeStr & " "

TimeStr = AttainTwoDigitNumber(Month(TaskTime + CDate("00:30:00")))
TimeStr = TimeStr & "/" & AttainTwoDigitNumber(Day(TaskTime + CDate("00:30:00")))
TimeStr = TimeStr & "/" & Year(TaskTime + CDate("00:30:00"))
CmdStr = CmdStr & "/ed " & TimeStr & " "

TimeStr = AttainTwoDigitNumber(Hour(TaskTime + CDate("00:30:00")))
TimeStr = TimeStr & ":" & AttainTwoDigitNumber(Minute(TaskTime + CDate("00:30:00")))
CmdStr = CmdStr & "/et " & TimeStr & " "

CmdStr = CmdStr & "/tr " & """" & "C:\Program Files (x86)\UiPath Platform\UiRobot.exe -file " & TaskFile & " --monitor" & """" & " "
CmdStr = CmdStr & "/Z"
Dim MyShell
Set MyShell = CreateObject("WSCript.shell")
MyShell.Run CmdStr, 0 ,True
Set MyShell = Nothing

End Function


'Function To Attain Two Digit Format Of Source Number
Function AttainTwoDigitNumber(SourceNumber)

Dim Temp
If Len(SourceNumber) = 1 Then
	Temp = "0" & SourceNumber
Else
	Temp = SourceNumber
End If
AttainTwoDigitNumber = Temp

End Function


'Function To Read File Data Into A 2-Dimension Array
Function ReadFile(FilePath)

Dim Fso
Set Fso = CreateObject("Scripting.FileSystemObject")
Set FileData = Fso.OpenTextFile(FilePath, 1)
Dim FinalArr()
Dim TempArr()
Dim Temp
Dim LineStr
Dim Counter
Counter = -1
Dim i
Do Until FileData.AtEndOfStream
	LineStr = FileData.readline
	If Len(LineStr) > 0 then
		Counter = Counter + 1
		Redim Preserve TempArr(Counter)
		TempArr(Counter) = LineStr
	End If
Loop
FileData.Close
ReDim FinalArr(Counter, 2)
For i = 0 to Counter
	Temp = split(TempArr(i),",")
	FinalArr(i,0) = Temp(0)
	If i > 0 Then
		If Instr(1, FilePath, "Schedule") > 0 Then
			If CDate(Temp(1)) <= CDate("00:10:00") Then
				FinalArr(i,1) = Date + CDate(Temp(1)) + CDate("24:00:00")
			Else
				FinalArr(i,1) = Date + CDate(Temp(1))
			End If
		Else
			FinalArr(i,1) = CDate(Temp(1))
		End If
	Else
		FinalArr(i,1) = Temp(1)
	End If
	FinalArr(i,2) = Temp(2)
next
If Instr(1, FilePath, "Schedule") = 0 Then
	Set FileData = Fso.OpenTextFile(FilePath, 2)
	FileData.WriteLine "Task_Name,Task_Date,Task_File"
	FileData.Close
End If
Set Fso = Nothing
ReadFile = FinalArr

End Function


'Reset Pending List File And Log File
Function PrepareFile(FilePath)

Dim Fso
Set Fso = CreateObject("Scripting.FileSystemObject")
If Fso.FileExists(FilePath) = False Then
	Fso.CreateTextFile(FilePath)
	Set FileData = Fso.OpenTextFile(FilePath, 2)
	FileData.WriteLine "Task_Name,Task_Date,Task_File"
	FileData.Close
end if
Set Fso = Nothing

End Function


'OpenTextFile
'ForReading = 1
'ForWriting = 2
'ForAppending = 8
