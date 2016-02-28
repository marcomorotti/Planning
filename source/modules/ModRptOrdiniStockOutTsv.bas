Option Compare Database


Public Sub GenerateCsvOrdersStockOut(Varwhere As Variant, PrintOut As Boolean)
' // Per usare questa funzione aggiungere librerie Excel e Word

' Dim conCurrent As ADODB.Connection
Dim Db As DAO.Database
Dim brs As DAO.Recordset
Dim rstOutput As New ADODB.Recordset
Dim objField As ADODB.Field
Dim intFile As Integer
Dim strSQL As String, strDataLine As String
Dim filenm As String
Dim i As Integer
' Costruzione File da aprire
intFile = FreeFile
i = 1

filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") _
           & "PortafoglioOrdini_" & i & ".tsv"

' Cerca File da aprire
    If FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, "yyyymmdd") _
    & "PortafoglioOrdini_" & i & ".tsv") = False Then
    Else
      Do While FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, "yyyymmdd") _
            & "PortafoglioOrdini_" & i & ".tsv") = True
       i = i + 1
      ' MsgBox "Il ciclo è stato eseguito " & i & " volte."
       Loop
    filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") _
           & "PortafoglioOrdini_" & i & ".tsv"
    End If
' Chiede se si vuole esportare il file
If vbYes = MsgBox("Vuoi esportare i dati in  " & filenm, _
        vbQuestion + vbYesNo + vbDefaultButton2, gstrAppTitle) Then
    Open filenm For Output As #intFile
    ' Inserisce testata
    strDataLine = "Pr" & Chr(9) & "Utente" & Chr(9) & "N_Doc" & Chr(9) & "Data_Ordine" & Chr(9) & "N_Cliente" & Chr(9) & "Rag_Soc" & Chr(9) & "Codice" & Chr(9) & _
                "Descrizione" & Chr(9) & "Qta_Ordine" & Chr(9) & "Data_Consegna" & Chr(9) & "Qtà_Spedita" & Chr(9) & _
                "Giacenza" & Chr(9) & "Impegnato" & Chr(9) & "Giacenza_Stefani" & Chr(9) & "N_Doc_Acq" & Chr(9) & _
                "Cod_F" & Chr(9) & "Rag_Soc_Fornitore" & Chr(9) & "Data_Ord_F" & Chr(9) & _
                "Qta_Residuo" & Chr(9) & "Data_Consegna"

    Print #intFile, strDataLine
    'Ora loop per il recordset e scrive un TSV file per ogni record
    Set Db = CurrentDb
    If IsNothing(Varwhere) Then
        Set brs = DBEngine(0)(0).OpenRecordset("SELECT * FROM qryCOrdersStockOutExe")
    Else
    Set brs = DBEngine(0)(0).OpenRecordset("SELECT * FROM qryCOrdersStockOutExe WHERE " & _
        Varwhere)
    End If


    If Not brs.EOF And Not brs.BOF Then
    brs.MoveFirst
    End If
    While Not brs.EOF
    strDataLine = brs.Fields("Liv_Urgenza").Value & Chr(9) & _
                    brs.Fields("Utente").Value & Chr(9) & _
                    brs.Fields("Numero_doc").Value & Chr(9) & _
                    brs.Fields("DataOrdine").Value & Chr(9) & _
                    brs.Fields("Cod_Cli").Value & Chr(9) & _
                    brs.Fields("Ds_Rag_soc").Value & Chr(9) & _
                    brs.Fields("Cod_Art").Value & Chr(9) & _
                    brs.Fields("Descrizione").Value & Chr(9) & _
                    brs.Fields("Qta_Ord_umv").Value & Chr(9) & _
                    brs.Fields("Data_Prev_Cons").Value & Chr(9) & _
                    brs.Fields("Qta_Cons_umv").Value & Chr(9) & _
                    brs.Fields("DispSp").Value & Chr(9) & _
                    brs.Fields("Impegnato").Value & Chr(9) & _
                    brs.Fields("DispAh").Value & Chr(9) & _
                    brs.Fields("Ord_Acq").Value & Chr(9) & _
                    brs.Fields("Cod_Forn").Value & Chr(9) & _
                    brs.Fields("Rag_Soc_Forn").Value & Chr(9) & _
                    brs.Fields("Data_Ordine").Value & Chr(9) & _
                    brs.Fields("Qta_Residua").Value & Chr(9) & _
                    brs.Fields("Data_Ric").Value
    Print #intFile, strDataLine
    brs.MoveNext
    strDataLine = ""
    Wend
    MsgBox ("I Dati sono stati salvati come TSV file" & CurrentProject.Path & _
        "\Export\" & Format(Date, "yyyymmdd") & "PortafoglioOrdini" & i & ".xls")
    brs.Close
    Set brs = Nothing
    Set Db = Nothing
    Close #intFile
Else
Exit Sub
End If
End Sub