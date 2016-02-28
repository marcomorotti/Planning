Option Compare Database
Public Sub ImportFromOracle()
    Dim AnnoC As String
    Dim MeseC As String
    Dim AnnoI As String
    Dim MeseI As String
    Dim AnnoF As String
    Dim MeseF As String
    Dim Cmd As New ADODB.Command
    Dim bqry As DAO.QueryDef
    Dim brs As DAO.Recordset
    Dim qr1 As QueryDef
    Dim Classe_Evento As String
    Dim MyQuery As String
    Dim dbo As Database
    Dim Db As Database
    Dim Lconnect As String
    Dim qm As QueryDef
    Dim rst As Object
    Dim rs As Recordset
    Dim conn As ADODB.Connection
    Dim DatiGen As New ADODB.Recordset

    ' 201602 Chiude Fase 01
    Call FaseEseguita("01")
    
    ' 201602 Inizia Fase 02
    Call Tempofase("02", True)
    
    ' Dim nRecords As Integer 'Long

    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    NumMesiC = DatiGen.Fields("Mesi_consumo")
    intYear = DatiGen.Fields("Anno_Calcolo")
    intMonth = DatiGen.Fields("Mese_Calcolo")
    intLastDay = Day(DateAdd("m", 1, DateSerial(intYear, intMonth, 1)) - 1)
    intEndDate = DateSerial(intYear, intMonth, intLastDay)
    intEndDate = Format(intEndDate, "dd/mmm/yyyy")
    intStartDate = DateAdd("m", -NumMesiC + 1, DateSerial(intYear, intMonth, 1))
    intStartDate = Format(intStartDate, "dd/mmm/yyyy")
    AnnoC = Year(intStartDate)
    MeseC = Month(intStartDate)
    AnnoF = Year(intEndDate)
    MeseF = Format(Month(intEndDate), "00")

    AnnoI = Year(DateAdd("m", -11, DateSerial(intYear, intMonth, 1)))
    MeseI = Format(Month(DateAdd("m", -11, DateSerial(intYear, intMonth, 1))), "00")

    ' *** CONNESSIONE a Oracle

    ' On Error GoTo Err_Execute

    'Use {Microsoft ODBC for Oracle} ODBC connection
    'Lconnect = "ODBC;DSN=sun3000.scmgroup.com;UID=VPERAZZINI;PWD=BIC;SERVER=sun3000.scmgroup.com"
    Lconnect = LeggiOdbcConnect
    'Point to the current workspace
    Set ws = DBEngine.Workspaces(0)

    'Connect to Oracle
    Set Db = ws.OpenDatabase("", False, True, Lconnect)
    ' Setto il tempo di QueryTimeOut a 240 min
    Db.QueryTimeout = 240

    '    GoTo start

    ' Scrive LOG FILE
    WriteToLog ("'Inizio: 1 - IMPORTA CONSUMI'")

    ' *** Inizio caricamento SpeditoMeseEventi.sql
    ' 14-09-12 Ho aggiunto spedizioni ' DA' Celaschi che dal 1/8/12 è entrata in SP e tolto le spedizioni
    '          al SP cod_cli = 15225
    ' 01-12-11 Ho aggiunto spedizioni ' AH' Stefani



    WriteToLog ("'  1.1 Import Codici Spediti'")
    ' ***************************************************************************************************
    '                             Import Codici Spediti
    ' ***************************************************************************************************

    'Visualizza lo Status Meter
    Call acbInitMeter("1-SPEDITO MESE da ORACLE", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    ' Cancella dati
    conn.Execute "Delete * From tblImportSpedito"
    MyQuery = "SELECT Spedito.cod_art, " & "Anagrafica.DESCR_ART, " & _
              "Spedito.Data_Bolla DataBolla, " & "ROUND(SUM (Spedito.qta_umv),0) qta_out, " & _
              "Anagrafica.LEAD_TIME, " & "Anagrafica.MIN_MINMAX_QUANTITY ROP,  " & _
              "Anagrafica.MAX_MINMAX_QUANTITY,  " & _
              "Anagrafica.MINIMUM_ORDER_QUANTITY ROQ,  " & _
              "Anagrafica.MAXIMUM_ORDER_QUANTITY, " & _
              "TRIM(TO_CHAR(CS_CSC,'999999D9999','nls_numeric_characters = '',.''')) as CS_CSC, grp.AN_ARTICOLO.tb_cl_merc, Anagrafica.Stato, "
    MyQuery = MyQuery & _
              "TRIM(TO_CHAR(Anagrafica.unit_weight,'999999D9999','nls_numeric_characters = '',.'''))  AS PESO_LORDO, " & _
              "TRIM(TO_CHAR(Anagrafica.unit_weight,'999999D9999','nls_numeric_characters = '',.'''))  AS PESO_NETTO"
    MyQuery = MyQuery & _
              " FROM grp.ricsped Spedito, grp.ricanart Anagrafica, grp.AN_ARTICOLO " & _
              "WHERE anagrafica.societa= ' SP' " & _
              "AND Spedito.cod_cli NOT IN " & _
              "('32913', '00999', '73894', '11738', '17180', '07210', '03884', '66709', " _
              & "'65927', '10815', '10360', '99845', '90348', '78901', '15225', '07100', '89161') " & _
              "AND Spedito.Societa IN (' SP', ' DA', ' AH', 'FF', 'FD', 'FE', 'FGB', 'TE') " & _
              "AND Spedito.COD_ART = Anagrafica.COD_ART " & _
              "AND Spedito.Data_Bolla BETWEEN' " & intStartDate & "' AND '" & intEndDate & _
              "' " & "AND CD_SOC(+) = ' SP' " & "AND Spedito.COD_ART = CD_ART(+) " & _
              "AND Spedito.COD_ART NOT LIKE '00014299%' " & _
              "AND Spedito.COD_ART NOT LIKE '00F%' " & _
              "AND Spedito.COD_ART NOT LIKE '40%' " & _
              "AND Spedito.COD_ART NOT LIKE '00005%' " & _
              "AND Spedito.COD_ART NOT LIKE '0200%' " & _
              "AND Spedito.TIPO_DOC_ORD NOT LIKE '%VF%' " & _
              "AND descr_caus_trasp NOT IN ('VEND.PER RIACQUISTO', 'TRASFERIMENTO' , 'OMAGGIO', " & _
              "'OMAGGIO DEPLIANT', 'VENDITA MAT.PUBBLICITARIO', 'VENDITA RIC.CONSIGLIATI', 'VENDITA RETROFIT') " & _
              "AND Spedito.qta_umv > 0 " & _
              "AND numero_doc_ord NOT IN ('1111511428', '1111511444', '1111511583', '1111525465', '1111525474', " & _
              " '1111525477', '1111525539', '1111525549', '1111526015') "
    ' 20151105 aggiunta esclusione ordini sopra per Zonghy
    MyQuery = MyQuery & _
              "GROUP BY Spedito.cod_art, " & _
              "Anagrafica.DESCR_ART, " & "Spedito.Data_Bolla, " & _
              "Anagrafica.TIPO_RIORD, " & _
              "Anagrafica.LEAD_TIME, " & "Anagrafica.MIN_MINMAX_QUANTITY, " & _
              "Anagrafica.MAX_MINMAX_QUANTITY, " & "Anagrafica.MINIMUM_ORDER_QUANTITY, " & _
              "Anagrafica.MAXIMUM_ORDER_QUANTITY, " & "CS_CSC, " & "grp.AN_ARTICOLO.tb_cl_merc, " & "Anagrafica.Stato, " & _
              "Anagrafica.unit_weight"

    Set rst = Db.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
    ' Numero Transazioni
    rst.MoveLast
    intMassimo = rst.RecordCount
    rst.MoveFirst
    Do While Not rst.EOF
        Set qm = db0.QueryDefs("qrySpeditoIsrt")
        qm!iCOD_ART = rst![Cod_art]
        qm!iDESCRIZIONE = rst![DESCR_ART]
        qm!iDATABOLLA = rst![DATABOLLA]
        qm!iQta_Out = rst![qta_out]
        qm!iLEAD_TIME = rst![Lead_time]
        qm!iROP = rst![ROP]
        qm!iMAX_MINMAX_QUANTITY = rst![MAX_MINMAX_QUANTITY]
        qm!iROQ = rst![ROQ]
        qm!iMAXIMUM_ORDER_QUANTITY = rst![MAXIMUM_ORDER_QUANTITY]
        'Replace(str(decNum), ",", ".") & ")"
        qm!iCS_CSC = rst![Cs_Csc]
        qm!iCATEG_MERC = rst![tb_cl_merc]
        qm!iSTATO = rst![Stato]
        qm!iPESO_LORDO = rst![Peso_Lordo]
        qm!iPESO_NETTO = rst![Peso_Netto]
        'Debug.Print ("Costo_dwh: " & rst![Cs_Csc])
        'Debug.Print ("Costo: " & Replace(rst![Cs_Csc], ".", ","))
        qm!iUpdateDate = Format(Date, "mm/dd/yyyy")
        qm.Execute
        rst.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        '                Call acbUpdateMeter(Int(intI))
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    rst.Close
start:
    '20160216 Puntini avanzamento
    '    Me.box02.Visible = True
    ' 201602 Chiude Fase 02
    Call FaseEseguita("02")
    
    ' 201602 Inizio Fase 03
    Call Tempofase("03", True)
    
    Forms("frmCalcoli").Form.box02.Visible = True
    WriteToLog ("'  1.2 Import Giacenza'")
    ' ***************************************************************************************************
    '                             Import Codici GIACENZA
    ' ***************************************************************************************************

    'Visualizza lo Status Meter
    Call acbInitMeter("2-GIACENZA da ORACLE", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    ' Cancella dati
    conn.Execute "Delete * From tblImportGiacenza"
    ' 20151028  was: "AND tb_ubic IN ('SATT', 'SAUT', 'SDIR', 'XFGB', 'XFPL') "
    MyQuery = "SELECT giacenza.CD_ART, " & "Anagrafica.DESCR_ART, " & _
              "ROUND(Sum(giacenza.qt_giac),0) As qt_giac, " & "Anagrafica.LEAD_TIME, " & _
              "Anagrafica.MIN_MINMAX_QUANTITY ROP,  " & _
              "Anagrafica.MAX_MINMAX_QUANTITY,  " & _
              "Anagrafica.MINIMUM_ORDER_QUANTITY ROQ,  " & _
              "Anagrafica.MAXIMUM_ORDER_QUANTITY, " & _
              "TRIM(TO_CHAR(AN_ARTICOLO.Cs_Csc,'999999D9999','nls_numeric_characters = '',.''')) as CS_CSC, " & _
              "grp.AN_ARTICOLO.tb_cl_merc, " & _
              "Anagrafica.Stato, " & _
              "TRIM(TO_CHAR(Anagrafica.unit_weight,'999999D9999','nls_numeric_characters = '',.''')) AS PESO_NETTO, " & _
              "TRIM(TO_CHAR(Anagrafica.unit_weight,'999999D9999','nls_numeric_characters = '',.'''))  AS PESO_LORDO "
    MyQuery = MyQuery & _
              "FROM grp.mag_articolo giacenza, GRP.AN_ARTICOLO AN_ARTICOLO, GRP.RICANART     Anagrafica " & _
              "WHERE Anagrafica.societa = ' SP' " & _
              "AND giacenza.cd_soc = ' SP' " & _
              "AND AN_ARTICOLO.CD_SOC = ' SP' " & _
              "AND giacenza.qt_giac > 0 " & _
              "AND AN_ARTICOLO.CD_ART = giacenza.CD_ART " & _
              "AND giacenza.CD_ART = Anagrafica.cod_art " & _
              "AND giacenza.CD_ART NOT LIKE '00014%' " & _
              "AND giacenza.CD_ART NOT LIKE '00F%' " & _
              "AND giacenza.CD_ART NOT LIKE '40%' " & _
              "AND giacenza.CD_ART NOT LIKE '00005%' " & _
              "AND giacenza.CD_ART NOT LIKE '0200%' " & _
              "AND giacenza.tb_ubic IN ('SATT', 'SAUT', ' COL') " & _
              "GROUP BY GIACENZA.CD_ART, Anagrafica.DESCR_ART, Anagrafica.LEAD_TIME, " & _
              "Anagrafica.MIN_MINMAX_QUANTITY, " & _
              "Anagrafica.MAX_MINMAX_QUANTITY, " & _
              "Anagrafica.MINIMUM_ORDER_QUANTITY, " & _
              "Anagrafica.MAXIMUM_ORDER_QUANTITY , AN_ARTICOLO.Cs_Csc, " & "grp.AN_ARTICOLO.tb_cl_merc, " & "Anagrafica.Stato, " & _
              "Anagrafica.UNIT_WEIGHT "
    Set rst = Db.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
    ' Numero Transazioni
    rst.MoveLast
    intMassimo = rst.RecordCount
    rst.MoveFirst
    Do While Not rst.EOF
        Set qm = db0.QueryDefs("qryGiacenzaIsrt")
        qm!iCD_ART = rst![Cd_Art]
        qm!iDescr_Art = rst![DESCR_ART]
        qm!iQT_GIAC = rst![qt_giac]
        qm!iLEAD_TIME = rst![Lead_time]
        qm!iROP = rst![ROP]
        qm!iMAX_MINMAX_QUANTITY = rst![MAX_MINMAX_QUANTITY]
        qm!iROQ = rst![ROQ]
        qm!iMAXIMUM_ORDER_QUANTITY = rst![MAXIMUM_ORDER_QUANTITY]
        qm!iCS_CSC = rst![Cs_Csc]
        qm!iCATEG_MERC = rst![tb_cl_merc]
        qm!iSTATO = rst![Stato]
        qm!iPESO_LORDO = rst![Peso_Lordo]
        qm!iPESO_NETTO = rst![Peso_Netto]
        qm!iUpdateDate = Format(Date, "mm/dd/yyyy")
        ' Debug.Print ("Costo: " & rst![Cs_Csc])
        qm.Execute
        rst.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    rst.Close
End Sub