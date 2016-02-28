Option Compare Database
Option Explicit
Function GetUser() As String
'** Procedure to Get the User's Name from the Windows Login
Dim si As SystemInfo
Set si = New SystemInfo
Dim strOut As String
'strGetUser = si.UserName
'strOut = si.UserName & " is logged into " & si.ComputerName
'GetUser = si.UserName  'Modificato MM 11/02/11
GetUser = si.ComputerName
If GetUser = "" Then
    MsgBox ("There is a problem with your Network Login Name!! Please contact your Network Administrator.")
    DoCmd.Quit
End If
End Function