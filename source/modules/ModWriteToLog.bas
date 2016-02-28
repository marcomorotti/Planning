' ***************** Funzione Scrive Log
' per usarla:
' WriteToLog ("Generic Log.vbs - Write This")

 Function WriteToLog(strLogMessage As String)
 
 strLogFileName = CurrentProject.Path & "\MyLog.log"
 strLogEntryTime = Now

 Open strLogFileName For Append As #1
 Print #1, strLogEntryTime & vtab & Chr(58) & Chr(9) & amp & strLogMessage
 Close #1
End Function

 