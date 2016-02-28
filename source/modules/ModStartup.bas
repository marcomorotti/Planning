Option Compare Database
Option Explicit

Public Function AttachAgain(strPath As String) As Integer
' This is a generic function that accepts a new path name
'  and attempts to refresh the links of all attached tables
' Input: Path name as C:\SomeFolder\SomeSubFolder
' Output: True if successful
Dim Db As DAO.Database, tdf As DAO.TableDef, rst As DAO.Recordset
Dim strFilePath As String, varRet As Variant, intFirst As Integer
Dim intI As Integer, intK As Integer, intL As Integer

    ' Get a pointer to the database
    Set Db = CurrentDb
    ' Initialize the full file path
    strFilePath = strPath & "\PortafoglioOrdini-Data.accdb"
    ' Set the "first table" indicator
    intFirst = True
    ' Turn on the progress meter
    varRet = SysCmd(acSysCmdInitMeter, "Reconnecting Data...", Db.TableDefs.Count)
    ' Set an error trap
    On Error GoTo Err_Attach
    ' Attempt to reattach the tables
    intI = 0 ' Reset the status meter counter
    For Each tdf In Db.TableDefs
        ' Looking for attached tables
        If (tdf.Attributes And dbAttachedTable) Then
            ' Figure out if this is mdb or xls file attached
            If InStr(tdf.connect, ".accdb") <> 0 Then
                ' Change the Connect property to point to the new file
                tdf.connect = ";DATABASE=" & strFilePath
                ' Attempt to refresh the link definition
                tdf.RefreshLink
                ' If the first table, then open a recordset
                '  to make this go faster
                If (intFirst = True) Then
                    Set rst = Db.OpenRecordset(tdf.name)
                    intFirst = False
                End If
                ElseIf InStr(tdf.connect, ".xls") <> 0 Then
                ' One of the Excel attached files - find the DATABASE part
                intK = InStr(tdf.connect, ";DATABASE=")
                ' Make sure we found it
                If intK <> 0 Then
                    ' Now find the file name
                    intL = InStrRev(tdf.connect, "\")
                    ' Make sure we found it
                    If intL <> 0 Then
                        ' Fix the Connect property
                        tdf.connect = left(tdf.connect, intK + 9) & strPath & _
                            Mid(tdf.connect, intL)
                        ' Attempt to refresh
                        tdf.RefreshLink
                    End If
                End If
            End If
        End If
        ' Update the status counter
        intI = intI + 1
        ' .. and update the progress meter
        varRet = SysCmd(acSysCmdUpdateMeter, intI)
        ' And pause for a sec so the status bar updates
        DoEvents
    Next tdf
    ' Done - clear the progress meter
    varRet = SysCmd(acSysCmdClearStatus)
    ' Clear the object variables
    Set tdf = Nothing
    rst.Close
    Set rst = Nothing
    Set Db = Nothing
    ' Return attach successful
    AttachAgain = True

Attach_Exit:
    Exit Function
    
Err_Attach:
    ' Uh, oh - failed.  Write a log record
    ErrorLog "AttachAgain " & strPath, Err, Error
    ' Clear the progress meter
    varRet = SysCmd(acSysCmdClearStatus)
    ' Clear the object variables
    Set tdf = Nothing
    Set Db = Nothing
    ' Return attach failed
    AttachAgain = False
    ' Exit
    Resume Attach_Exit

End Function


Public Function CheckVersion(curVNo As Currency) As Integer
    ' Software vs data file version checker
    ' Input: version number from the attached data file
    ' Return:  True if this software version is compatible

    ' Check the integer portion of both versions
    ' This allows minor update revisions to the code (v1.1, v1.2) that
    '  will still work with "base" version of the data and vice-versa.
    
    If Int(curVNo) <> Int(gTHISVERSION) Then
        ' Base versions not equal - display appropriate error and bail
        If curVNo < gTHISVERSION Then
            MsgBox "The version of this application code is later than your data tables. " & _
                "PortafoglioOrdini for the special procedure to upgrade your data tables to work with this code.", _
                vbCritical, "PortafoglioOrdini Ced"
        Else
            MsgBox "The version of this application code is earlier than your data tables. " & _
                "PortafoglioOrdini  for a more up-to-date version of the code.", vbCritical, _
                "PortafoglioOrdini Manager"
        End If
        CheckVersion = False
    Else
        CheckVersion = True
    End If

End Function

Public Function ReConnect()
Dim Db As DAO.Database, tdf As DAO.TableDef, rst As DAO.Recordset, rstV As DAO.Recordset
Dim strFile As String, varRet As Variant, frm As Form, strPath As String, intI As Integer

' This is a slightly different version of reconnect code
' Called by frmSplash - the normal startup form for this application

    On Error Resume Next
    Set Db = CurrentDb

    ' Turn on the hourglass - this may take a few secs.
    DoCmd.Hourglass True
    ' First, check linked table version
    Set rstV = Db.OpenRecordset("ztblVersion")
    ' Got a failure - so try to reattach the tables
    If Err <> 0 Then GoTo Reattach
    ' Make sure we're on the first row
    rstV.MoveFirst
    ' Call the version checker
    If Not CheckVersion(rstV!Version) Then
        ' Tell caller that "reconnect" failed
        ReConnect = False
        ' Close the version recordset
        rstV.Close
        ' Clear the objects
        Set rstV = Nothing
        Set Db = Nothing
        ' Done
        DoCmd.Hourglass False
        Exit Function
    End If
    ' Versions match - now verify all the other tables
    ' NOTE: We're leaving rstV open at this point for better efficiency
    '   in a shared database environment.  JET will share the already established thread.
    ' Turn on the progress meter on the status bar
    varRet = SysCmd(acSysCmdInitMeter, "Verifying data tables...", Db.TableDefs.Count)
    ' Loop through all TableDefs
    For Each tdf In Db.TableDefs
        ' Looking for attached tables
        If (tdf.Attributes And dbAttachedTable) Then
            ' Try to open the table
            Set rst = tdf.OpenRecordset()
            ' If got an error - then try to relink
            If Err <> 0 Then GoTo Reattach
            ' This one OK - close it
            rst.Close
            ' And clear the object
            Set rst = Nothing
        End If
        ' Update the progress counter
        intI = intI + 1
        varRet = SysCmd(acSysCmdUpdateMeter, intI)
    Next tdf
    ' Got through them all - clear the progress meter
    varRet = SysCmd(acSysCmdClearStatus)
    ' Turn off the hourglass
    DoCmd.Hourglass False
    ' Set a good return
    ReConnect = True
    ' Edit the Version table
    rstV.Edit
    ' Update the open count - we check this on exit to recommend a backup
    rstV!OpenCount = rstV!OpenCount + 1
    ' Update the row
    rstV.Update
    ' Close and clear the objects
    rstV.Close
    Set rstV = Nothing
    Set Db = Nothing
    ' DONE!
    Exit Function

Reattach:
    ' Clear the current error
    Err.Clear
    ' Set a new error trap
    On Error GoTo BadReconnect
    ' Turn off the hourglass for now
    DoCmd.Hourglass False
    ' .. and clear the status bar
    varRet = SysCmd(acSysCmdClearStatus)
    ' Tell the user about the problem - about to show an open file dialog
    MsgBox "There's a temporary problem connecting to the Portafoglio Ordini Data.  Please locate the PortafoglioOrdini-Data file in the following dialog.", vbInformation, "PortafoflioOrdini-DATA"
    ' Establish a new ComDlg object
    With New ComDlg
        ' Set the title of the dialog
        .DialogTitle = "Locate PortafoglioOrdini-DATA"
        ' Set the default file name
        .fileName = "PortafoglioOrdini-Data.accdb"
        ' .. and start directory
        .directory = CurrentProject.Path
        ' .. and file extension
        .Extension = "accdb"
        ' .. but show all mdb files just in case
        .Filter = "PortafoglioOrdini-DATA (*.accdb)|*.accdb"
        ' Default directory is where this file is located
        .directory = CurrentProject.Path
        ' Tell the common dialog that the file and path must exist
        .ExistFlags = FileMustExist + PathMustExist
        If .ShowOpen Then
            strFile = .fileName
        Else
            Err.Raise 3999
        End If
    End With
    ' Open the "info" form telling what we're doing
    DoCmd.OpenForm "frmReconnect"
    ' .. and be sure it has the focus
    Forms!frmReconnect.SetFocus
    ' Attempt to re-attach the Version table first and check it
    Set tdf = Db.TableDefs("ztblVersion")
    tdf.connect = ";DATABASE=" & strFile
    tdf.RefreshLink
    ' OK, now check linked table version
    Set rst = Db.OpenRecordset("ztblVersion")
    rst.MoveFirst
    ' Call the version checker
    If Not CheckVersion(rst!Version) Then
        ' Tell the caller that we failed
        ReConnect = False
        ' Close the version recordset
        rst.Close
        ' .. and clear the object
        Set rst = Nothing
        ' Bail
        Exit Function
    End If
    ' Passed version check - edit the version record
    rst.Edit
    ' Update the open count - we check this on exit to recommend a backup
    rst!OpenCount = rst!OpenCount + 1
    ' Write it back
    rst.Update
    ' Close the recordset
    rst.Close
    ' .. and clear the object
    Set rst = Nothing
    ' Now, reattach the other tables
    ' Strip out just the path name
    strPath = left(strFile, InStrRev(strFile, "\") - 1)
    ' Call the generic re-attach function
    If AttachAgain(strPath) = 0 Then
        ' Oops - failed.  Raise an error
        Err.Raise 3999
    End If
    ' Close the information form
    DoCmd.Close acForm, "frmReconnect"
    ' Clear the db object
    Set Db = Nothing
    ' Return a positive result
    ReConnect = True
    ' .. and exit
    
Connect_Exit:
    Exit Function

BadReconnect:
    ' Ooops
    MsgBox "Reconnect to data failed.", vbCritical, "PortafoglioOrdini"
    ' Indicate failure
    ReConnect = False
    ' Close the info form if it is open
    If IsFormLoaded("frmReconnect") Then DoCmd.Close acForm, "frmReconnect"
    ' Clear the progress meter
    varRet = SysCmd(acSysCmdClearStatus)
    ' .. and bail
    Resume Connect_Exit

End Function