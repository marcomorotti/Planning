Option Compare Database


Public Sub GenerateCsvArticoli(Varwhere As Variant, PrintOut As Boolean)
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

filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") & _
    "AnagraficaArt_" & i & ".tsv"

' Cerca File da aprire
    If FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, _
        "yyyymmdd") & "AnagraficaArt_" & i & ".tsv") = False Then
    Else
      Do While FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & _
          Format(Date, "yyyymmdd") & "AnagraficaArt_" & i & ".tsv") = True
       i = i + 1
      ' MsgBox "Il ciclo è stato eseguito " & i & " volte."
       Loop
    filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") & _
        "AnagraficaArt_" & i & ".tsv"
    End If
' Chiede se si vuole esportare il file
If vbYes = MsgBox("Vuoi esportare i dati in  " & filenm, vbQuestion + vbYesNo + _
    vbDefaultButton2, gstrAppTitle) Then
    Open filenm For Output As #intFile
    ' Inserisce testata
    strDataLine = "Cod_art" & Chr(9) & "Des_art" & Chr(9) & "Stato" & Chr(9) & _
        "Classe_Evento" & Chr(9) & "LivelloServizio" & Chr(9) & "Cs_Csc" & Chr(9) & _
        "ClasseCosto" & Chr(9) & "Lead_time" & Chr(9) & "ROP" & Chr(9) & _
        "Punto_riordino" & Chr(9) & "ROQ" & Chr(9) & "Lotto_ec_acq" & Chr(9) & _
        "Scorta_Sic_za_Forzata" & Chr(9) & "Giac_Media" & Chr(9) & "Copertura" & Chr(9) & "SConsumo" & Chr(9) & _
        "SSpedito" & Chr(9) & "SConsumo_12" & Chr(9) & "SSpedito_12" & Chr(9) & _
        "AbcGiacenza" & Chr(9) & "AbcConsumo"

    Print #intFile, strDataLine
    'Ora loop per il recordset e scrive un TSV file per ogni record
    Set Db = CurrentDb
    If IsNothing(Varwhere) Then
        Set brs = DBEngine(0)(0).OpenRecordset("SELECT * FROM qryArticoliSearch")
    Else
    Set brs = _
        DBEngine(0)(0).OpenRecordset("SELECT * FROM qryArticoliSearch WHERE " & _
        Varwhere)
    End If


    If Not brs.EOF And Not brs.BOF Then
    brs.MoveFirst
    End If
    While Not brs.EOF
    strDataLine = brs.Fields("Cod_art").Value & Chr(9) & _
        brs.Fields("Des_art").Value & Chr(9) & brs.Fields("Stato").Value & Chr(9) & _
        brs.Fields("Classe_Evento").Value & Chr(9) & _
        brs.Fields("LivelloServizio").Value & Chr(9) & brs.Fields("Cs_Csc").Value & _
        Chr(9) & brs.Fields("ClasseCosto").Value & Chr(9) & _
        brs.Fields("Lead_time").Value & Chr(9) & brs.Fields("ROP").Value & Chr(9) & _
        brs.Fields("Punto_riordino").Value & Chr(9) & brs.Fields("ROQ").Value & _
        Chr(9) & brs.Fields("Lotto_ec_acq").Value & _
        Chr(9) & brs.Fields("ScortaSicurezzaForzata").Value & Chr(9) & _
        brs.Fields("Giac_Media").Value & Chr(9) & brs.Fields("Copertura").Value & _
        Chr(9) & brs.Fields("SConsumo").Value & Chr(9) & _
        brs.Fields("SSpedito").Value & Chr(9) & brs.Fields("SConsumo_12").Value & _
        Chr(9) & brs.Fields("SSpedito_12").Value & Chr(9) & _
        brs.Fields("AbcGiacenza").Value & Chr(9) & brs.Fields("AbcConsumo").Value
   
    Print #intFile, strDataLine
    brs.MoveNext
    strDataLine = ""
    Wend
    MsgBox ("I Dati sono stati salvati come TSV file" & CurrentProject.Path & _
        "\Export\" & Format(Date, "yyyymmdd") & "AnagraficaArt" & i & ".xls")
    brs.Close
    Set brs = Nothing
    Set Db = Nothing
    Close #intFile
Else
Exit Sub
End If
End Sub