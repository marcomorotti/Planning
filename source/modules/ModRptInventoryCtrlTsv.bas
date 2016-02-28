Option Compare Database

Public Sub GenerateInventoryControl(Varwhere As Variant, PrintOut As Boolean)


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
    "Inventory_Ctrl_" & i & ".tsv"

' Cerca File da aprire
    If FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, _
        "yyyymmdd") & "Inventory_Ctrl_" & i & ".tsv") = False Then
    Else
      Do While FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & _
          Format(Date, "yyyymmdd") & "Inventory_Ctrl_" & i & ".tsv") = True
       i = i + 1
      ' MsgBox "Il ciclo è stato eseguito " & i & " volte."
       Loop
    filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") & _
        "Inventory_Ctrl_" & i & ".tsv"
    End If
' Chiede se si vuole esportare il file
    Open filenm For Output As #intFile
    ' Inserisce testata
    strDataLine = "Cod_art" & Chr(9) & "Des_art" & Chr(9) & "Qty_Venduta" & Chr(9) & "Qty_Giac" & Chr(9) & _
        "Qty_Acquistata" & Chr(9) & _
        "Punto_Riordino" & Chr(9) & "Scorta_Sicurezza" & Chr(9) & "Lotto_Riordino" & Chr(9) & "Disponibile" & Chr(9) & _
        "QTY_DA_RIORDINARE_PER_STOCK_OUT" & Chr(9) & "Cs_Csc"

    Print #intFile, strDataLine
    'Ora loop per il recordset e scrive un TSV file per ogni record
    Set Db = CurrentDb
    ' Chiede se si vogliono togliere i Very-Slow
    If IsNothing(Varwhere) Then
        Set brs = DBEngine(0)(0).OpenRecordset("SELECT * FROM qryInventoryControl")
    Else
        Set brs = _
            DBEngine(0)(0).OpenRecordset("SELECT * FROM qryInventoryControl WHERE " & _
            Varwhere)
    End If


    If Not brs.EOF And Not brs.BOF Then
    brs.MoveFirst
    End If
    While Not brs.EOF
    strDataLine = brs.Fields("Cd_art").Value & Chr(9) & _
        brs.Fields("Des_art").Value & Chr(9) & brs.Fields("Qty_Sale").Value & Chr(9) & brs.Fields("Qty_Stock").Value & Chr(9) & _
        brs.Fields("Qty_Acq").Value & Chr(9) & brs.Fields("Punto_Riordino").Value & Chr(9) & brs.Fields("ScortaSicurezza").Value & Chr(9) & _
        brs.Fields("Lotto_ec_acq").Value & Chr(9) & brs.Fields("Disponibile").Value & _
        Chr(9) & brs.Fields("Lotto_Acquisto").Value & Chr(9) & _
        brs.Fields("Cs_Csc").Value
    Print #intFile, strDataLine
    brs.MoveNext
    strDataLine = ""
    Wend
    MsgBox ("I Dati sono stati salvati come TSV file" & CurrentProject.Path & _
        "\Export\" & Format(Date, "yyyymmdd") & "Inventory_Ctrl_" & i & ".xls")
    brs.Close
    Set brs = Nothing
    Set Db = Nothing
    Close #intFile

End Sub