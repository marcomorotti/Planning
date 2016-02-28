Option Compare Database
' global variable for path to original database location
Public g_strFilePath As String
' global variable for path to database to copy from
Public g_strCopyLocation As String


Public Sub UpdateFrontEnd()
Dim strCmdBatch As String
Dim notNotebook As Object
Dim FSys As Object
Dim TestFile As String
Dim strKillFile As String
Dim strReplFile As String
Dim strRestart As String

' sets the file name and location for the file to delete
strKillFile = g_strFilePath
' sets the file name and location for the file to copy
strReplFile = g_strCopyLocation & "\" & CurrentProject.name
' sets the file name of the batch file to create
TestFile = CurrentProject.Path & "\UpdateDbFE.cmd"
' sets the restart file name
strRestart = """" & strKillFile & """"
' creates the batch file
Open TestFile For Output As #1
Print #1, "Echo Off"
Print #1, "ECHO Deleting old file"
Print #1, ""
Print #1, "ping 1.1.1.1 -n 1 -w 2000"
Print #1, ""
Print #1, "Del """ & strKillFile & """"
Print #1, ""
Print #1, "ECHO Copying new file"
Print #1, "Copy /Y """ & strReplFile & """ """ & strKillFile & """"
Print #1, ""
Print #1, "CLICK ANY KEY TO RESTART THE ACCESS PROGRAM"
Print #1, "START /I " & """MSAccess.exe"" " & strRestart
Close #1
'Exit Sub
' runs the batch file
Shell TestFile

'closes the current version and runs the batch file
DoCmd.Quit



End Sub