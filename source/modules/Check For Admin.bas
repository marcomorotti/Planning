Option Compare Database
Function IsAdmin(LID As Variant) As Boolean
'This function simply checks to see if the current loginID from getuser() exists in the dbo_AdminTable.
'If it does exist, the function returns True to the calling line.
'If it doesn't exist, the function returns False to the calling line.
Dim strSQL As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
strSQL = "Select * from dbo_AdminTable where LoginID = '" & LID & "'"
rs.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
If rs.EOF And rs.BOF Then
    IsAdmin = False
Else
    IsAdmin = True
End If
rs.Close
Set rs = Nothing
End Function
Function RetSecLevel(LID As Variant) As Integer
'This function returns the SecurityLevel found in the dbo_AdminTable.
'If the loginID is not found in the dbo_AdminTable, 0 is returned (i.e. no security.)
Dim strSQL As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
strSQL = "Select * from dbo_AdminTable where LoginID = '" & LID & "'"
rs.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
If rs.EOF And rs.BOF Then
    RetSecLevel = 0
Else
    RetSecLevel = rs!SecurityLevel
End If
rs.Close
Set rs = Nothing
End Function
Function IsNewUser()
'This function is called when the GetUser form is initially opened.
'It checks to see if the current loginID from getuser() is in the dbo_AdminTable.
'If it is not found, then it asks if you'd like to add the current loginID to the dbo_AdminTable.
Dim LID As Variant
LID = GetUser()
Dim strSQL As String
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
strSQL = "Select * from dbo_AdminTable where LoginID = '" & LID & "'"
rs.Open strSQL, CurrentProject.Connection, adOpenDynamic, adLockOptimistic
If rs.EOF And rs.BOF Then
    Dim QI As Integer
    QI = MsgBox("Il tuo LoginID: " & LID & " non è presente nel Data Base.  Vuoi aggiungere la loginID al Data Base?", vbYesNo)
    If QI = vbYes Then
        rs.AddNew
        rs!LoginID = LID
        rs!SecurityLevel = 1
        rs.Update
        MsgBox LID & " è stato aggiunto con il livello 1 (Chiudi la mappa Principale con il bottone in alto a dx e rilancia il programma.)"
    End If
End If
rs.Close
Set rs = Nothing
End Function