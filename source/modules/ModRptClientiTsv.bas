
Public Sub GenerateCsvClienti()
' // Per usare questa funzione aggiungere librerie Excel e Word

' Dim conCurrent As ADODB.Connection
Dim brs As DAO.Recordset
Dim rstOutput As New ADODB.Recordset
Dim objField As ADODB.Field
Dim intFile As Integer
Dim strSQL As String, strDataLine As String
Dim filenm As String
Dim i As Integer
Dim MyQuery As String
Dim Db As Database
Dim Lconnect As String



'Use {Microsoft ODBC for Oracle} ODBC connection
    'Lconnect = "ODBC;DSN=sun3000.scmgroup.com;UID=VPERAZZINI;PWD=BIC;SERVER=sun3000.scmgroup.com"
    Lconnect = LeggiOdbcConnect
    'Point to the current workspace
    Set ws = DBEngine.Workspaces(0)
    
    'Connect to Oracle
    Set Db = ws.OpenDatabase("", False, True, Lconnect)
    ' Setto il tempo di QueryTimeOut a 240 min
    Db.QueryTimeout = 240



' Costruzione File da aprire
intFile = FreeFile
i = 1

filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") & _
    "AnagraficaClienti_" & i & ".tsv"

' Cerca File da aprire
    If FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, _
        "yyyymmdd") & "AnagraficaClienti_" & i & ".tsv") = False Then
    Else
      Do While FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & _
          Format(Date, "yyyymmdd") & "AnagraficaClienti_" & i & ".tsv") = True
       i = i + 1
      ' MsgBox "Il ciclo è stato eseguito " & i & " volte."
       Loop
    filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") & _
        "AnagraficaClienti_" & i & ".tsv"
    End If
' Chiede se si vuole esportare il file
If vbYes = MsgBox("Vuoi esportare i dati in  " & filenm, vbQuestion + vbYesNo + _
    vbDefaultButton2, gstrAppTitle) Then
    Open filenm For Output As #intFile
    ' Inserisce testata
    strDataLine = "Flag" & Chr(9) & "Codice_Soggetto" & Chr(9) & "Ragione_Sociale" & Chr(9) & _
        "Nazione" & Chr(9) & "Data_Ins" & Chr(9) & "Data_Mod" & Chr(9) & _
        "Via" & Chr(9) & "Cap" & Chr(9) & "Città" & Chr(9) & _
        "Provincia" & Chr(9) & "Telefono" & Chr(9) & "Fax" & Chr(9) & _
        "Email"

    Print #intFile, strDataLine
   MyQuery = "SELECT a.flg_clifor, " & _
        "a.cd_clfo, " & _
        "a.DS_RAG_soc, " & _
        "b.ds_naz, " & _
        "a.dt_isrt, " & _
        "a.dt_repl, " & _
        "a.ds_via, " & _
        "a.cd_cap, " & _
        "a.ds_citta, "
    MyQuery = MyQuery & _
        "a.cd_provincia," & _
        "a.nu_tel, " & _
        "a.nu_fax, " & _
        "a.Email " & _
        " FROM GRP.AN_CLIFOR_GRP a, grp.tb_nazio_acst b " & _
        "WHERE a.tb_naz = b.tb_naz(+) AND a.FLG_CLIFOR != 'F'"
    'Ora loop per il recordset e scrive un TSV file per ogni record
        Set brs = Db.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
        ' Set brs = DBEngine(0)(0).OpenRecordset(MyQuery)
    If Not brs.EOF And Not brs.BOF Then
    brs.MoveFirst
    End If
    While Not brs.EOF
    strDataLine = brs.Fields("flg_clifor").Value & Chr(9) & _
        brs.Fields("cd_clfo").Value & Chr(9) & brs.Fields("DS_RAG_soc").Value & Chr(9) & _
        brs.Fields("ds_naz").Value & Chr(9) & _
        brs.Fields("dt_isrt").Value & Chr(9) & brs.Fields("dt_repl").Value & _
        Chr(9) & brs.Fields("ds_via").Value & Chr(9) & _
        brs.Fields("cd_cap").Value & Chr(9) & brs.Fields("ds_citta").Value & Chr(9) & _
        brs.Fields("cd_provincia").Value & Chr(9) & brs.Fields("nu_tel").Value & _
        Chr(9) & brs.Fields("nu_fax").Value & _
        Chr(9) & brs.Fields("Email").Value
   
    Print #intFile, strDataLine
    brs.MoveNext
    strDataLine = ""
    Wend
    MsgBox ("I Dati sono stati salvati come TSV file" & CurrentProject.Path & _
        "\Export\" & Format(Date, "yyyymmdd") & "AnagraficaClienti" & i & ".xls")
    brs.Close
    Set brs = Nothing
    Set Db = Nothing
    Close #intFile
Else
Exit Sub
End If
End Sub