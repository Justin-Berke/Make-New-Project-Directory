'Grab project name from user
dim input
input = InputBox ("Enter the name of the new project", "Make Project Directory")

'Generate & format date objects
dim txtYear
dim txtMonth
dim txtDay
txtYear = Year(Date())
txtMonth = Month(Date())
txtDay = Day(Date())

If txtMonth < 10 Then txtMonth = "0" & txtMonth
If txtDay < 10 Then txtDay = "0" & txtDay

'Create file system object
Dim oFSO
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Create folder path
Dim path
path = "C:\- Projects\" & _
	txtYear & "-" & txtMonth & "-" & txtDay & " - " & input
	
'Create folder
oFSO.CreateFolder(path)

'Create subdirectories
oFSO.CreateFolder(path & "\Subdirectory Name")