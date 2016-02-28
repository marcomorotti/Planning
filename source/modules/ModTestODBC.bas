Function OracleConnect() As Boolean
Dim ws As Workspace
Dim Db As Database
Dim Lconnect As String
Dim MyQuery As String
Dim MyRs As Recordset
On Error GoTo Err_execute

'Use {Microsoft ODBC for Oracle} ODBC connection
    ' Lconnect = "ODBC;DSN=sun3000.scmgroup.com;UID=VPERAZZINI;PWD=BIC;SERVER=sun3000.scmgroup.com"
    ' 29/06/2012 Modificato la lettura della stringa ODBC che viene letta dalla tabella dbo_AdminTable
    ' e se non esiste dalla tabella Parametri
    Lconnect = LeggiOdbcConnect
    'Point to the current workspace
    Set ws = DBEngine.Workspaces(0)

    'Connect to Oracle
    Set Db = ws.OpenDatabase("", False, True, Lconnect)
    ' Setto il tempo di QueryTimeOut a 120 min
    Db.QueryTimeout = 240

    
    MyQuery = "Select articolo from SIM.STORICO_CONSUMI_RIC where articolo = '0000107048A'"
    Set rst = Db.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
    Do While Not rst.EOF
        Debug.Print rst![Articolo]
        rst.MoveNext
    Loop
    rst.Close
    Db.Close

    OracleConnect = True
    MsgBox "Connessione a Oracle OK."
    Exit Function

Err_execute:
    MsgBox "Connessione a Oracle FALLITA."
    OracleConnect = False



End Function