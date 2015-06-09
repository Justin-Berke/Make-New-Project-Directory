'1) Create a directory with the format: YYYY-MM-DD - Description
'2) Prompt user with option to create common subdirectories

'Grab project name from user
dim txtProjectName
txtProjectName = InputBox ("Enter the name of the new project", "Make Project Directory")

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
	txtYear & "-" & txtMonth & "-" & txtDay & " - " & txtProjectName
	
'Create main folder
oFSO.CreateFolder(path)

'Prompt to create common subdirectories
Dim boolAddSubdirectories
boolAddSubdirectories = MsgBox("Do you want to add common subdirectories?", vbYesNo, "Add Subdirectories")

If boolAddSubdirectories = vbYes Then
	oFSO.CreateFolder(path & "\Exports")
	' Add additional directories here...
End If