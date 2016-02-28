Option Compare Database

Public Sub GenerateCsvRopRoq(Varwhere As Variant, PrintOut As Boolean)
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
    "RopRoq_" & i & ".tsv"

' Cerca File da aprire
    If FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, _
        "yyyymmdd") & "RopRoq_" & i & ".tsv") = False Then
    Else
      Do While FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & _
          Format(Date, "yyyymmdd") & "RopRoq_" & i & ".tsv") = True
       i = i + 1
      ' MsgBox "Il ciclo è stato eseguito " & i & " volte."
       Loop
    filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") & _
        "RopRoq_" & i & ".tsv"
    End If
' Chiede se si vuole esportare il file
    Open filenm For Output As #intFile
    ' Inserisce testata
    strDataLine = "Cod_art" & Chr(9) & "Des_art" & Chr(9) & "Stato" & Chr(9) & "LivelloServizio" & Chr(9) & _
        "Rop_Attuale" & Chr(9) & "Rop_Proposto" & Chr(9) & "ScortaSicurezzaForzata" & Chr(9) & "Roq_Attuale" & Chr(9) & _
        "Roq_proposto" & Chr(9) & "AbcConsValo" & Chr(9) & "Classe_Evento" & Chr(9) & "Giacenza" & _
        Chr(9) & "Consumo_annuo" & Chr(9) & "Csc" & Chr(9) & "Classe_Merc"

    Print #intFile, strDataLine
    'Ora loop per il recordset e scrive un TSV file per ogni record
    Set Db = CurrentDb
    ' Chiede se si vogliono togliere i Very-Slow
    If IsNothing(Varwhere) Then
        Set brs = DBEngine(0)(0).OpenRecordset("SELECT * FROM qryArticoliRopRoq")
    Else
        Set brs = _
            DBEngine(0)(0).OpenRecordset("SELECT * FROM qryArticoliRopRoq WHERE " & _
            Varwhere)
    End If


    If Not brs.EOF And Not brs.BOF Then
    brs.MoveFirst
    End If
    While Not brs.EOF
    strDataLine = brs.Fields("Cod_art").Value & Chr(9) & _
        brs.Fields("Des_art").Value & Chr(9) & brs.Fields("Stato").Value & Chr(9) & brs.Fields("LivelloServizio").Value & Chr(9) & _
        brs.Fields("RopAct").Value & Chr(9) & brs.Fields("RopProp").Value & Chr(9) & _
        brs.Fields("ScortaSicurezzaForzata").Value & Chr(9) & brs.Fields("RoqAct").Value & _
        Chr(9) & brs.Fields("RoqProp").Value & Chr(9) & _
        brs.Fields("AbcConsumoValoreLs").Value & Chr(9) & brs.Fields("Classe_Evento").Value & _
        Chr(9) & brs.Fields("Giac_Media").Value & _
        Chr(9) & brs.Fields("SConsumo_12").Value & _
        Chr(9) & brs.Fields("Cs_Csc").Value & _
        Chr(9) & brs.Fields("Categ_merc")
    Print #intFile, strDataLine
    brs.MoveNext
    strDataLine = ""
    Wend
    MsgBox ("I Dati sono stati salvati come TSV file" & CurrentProject.Path & _
        "\Export\" & Format(Date, "yyyymmdd") & "RopRoq_" & i & ".xls")
    brs.Close
    Set brs = Nothing
    Set Db = Nothing
    Close #intFile

End Sub