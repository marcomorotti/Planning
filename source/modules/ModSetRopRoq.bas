Option Compare Database
Public Sub SetRopRoq()
    Dim intI, IntR, IntC As Integer
    Dim LS(8, 9) As Variant
    Dim rs As ADODB.Recordset
    Dim db0 As Database
    Dim conn As ADODB.Connection
    Dim bqry As DAO.QueryDef
    Dim brs As DAO.Recordset
    Dim qr1 As QueryDef
'   GoTo start
    ' **************************************************************************************************
    ' leggo i Livelli di Servizio
    ' **************************************************************************************************
    Set rs = New ADODB.Recordset
    ' Multiple array
    rs.Open "SELECT * FROM [tblCategoriaEventiCva];", CurrentProject.Connection, _
            adOpenKeyset, adLockOptimistic

    Do While Not rs.EOF
        For IntR = 1 To 8
            For IntC = 1 To 9
                ' LS(intI, intJ) = Replace(rs.Fields(intJ + 1) * 100, ",", ".") 'Livello servizio
                LS(IntR, IntC) = rs.Fields(IntC + 1) * 100    'Livello servizio da 1 a 10
                ' Debug.Print intI & intJ & LS(intI, intJ)
            Next IntC
            rs.MoveNext
        Next IntR
    Loop
    rs.Close
    Set rs = Nothing

    ' **************************************************************************************************
    ' aggiorno i Livelli di Servizio
    ' **************************************************************************************************
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("Aggiorna Ls", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryArticoliAbcConsumoEvento")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i Ls in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art"), brs.Fields("Classe_Evento"), brs.Fields("AbcConsumoValoreLs")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryLsUpdate")
        qr1!iCOD_ART = brs![Cod_art]
        Select Case brs![Classe_Evento]
        Case "Very-Slow"
            Select Case brs![AbcConsumoValoreLs]
            Case "A1"
                qr1!iLivelloServizio = LS(2, 1)
            Case "A2"
                qr1!iLivelloServizio = LS(2, 2)
            Case "A3"
                qr1!iLivelloServizio = LS(2, 3)
            Case "A4"
                qr1!iLivelloServizio = LS(2, 4)
            Case "B1"
                qr1!iLivelloServizio = LS(2, 5)
            Case "B2"
                qr1!iLivelloServizio = LS(2, 6)
            Case "B3"
                qr1!iLivelloServizio = LS(2, 7)
            Case "C1"
                qr1!iLivelloServizio = LS(2, 8)
            Case "C2"
                qr1!iLivelloServizio = LS(2, 9)
            End Select
        Case "Slow"
            Select Case brs![AbcConsumoValoreLs]
            Case "A1"
                qr1!iLivelloServizio = LS(3, 1)
            Case "A2"
                qr1!iLivelloServizio = LS(3, 2)
            Case "A3"
                qr1!iLivelloServizio = LS(3, 3)
            Case "A4"
                qr1!iLivelloServizio = LS(3, 4)
            Case "B1"
                qr1!iLivelloServizio = LS(3, 5)
            Case "B2"
                qr1!iLivelloServizio = LS(3, 6)
            Case "B3"
                qr1!iLivelloServizio = LS(3, 7)
            Case "C1"
                qr1!iLivelloServizio = LS(3, 8)
            Case "C2"
                qr1!iLivelloServizio = LS(3, 9)
            End Select
        Case "Medium-Slow"
            Select Case brs![AbcConsumoValoreLs]
            Case "A1"
                qr1!iLivelloServizio = LS(4, 1)
            Case "A2"
                qr1!iLivelloServizio = LS(4, 2)
            Case "A3"
                qr1!iLivelloServizio = LS(4, 3)
            Case "A4"
                qr1!iLivelloServizio = LS(4, 4)
            Case "B1"
                qr1!iLivelloServizio = LS(4, 5)
            Case "B2"
                qr1!iLivelloServizio = LS(4, 6)
            Case "B3"
                qr1!iLivelloServizio = LS(4, 7)
            Case "C1"
                qr1!iLivelloServizio = LS(4, 8)
            Case "C2"
                qr1!iLivelloServizio = LS(4, 9)
            End Select
        Case "Medium"
            Select Case brs![AbcConsumoValoreLs]
            Case "A1"
                qr1!iLivelloServizio = LS(5, 1)
            Case "A2"
                qr1!iLivelloServizio = LS(5, 2)
            Case "A3"
                qr1!iLivelloServizio = LS(5, 3)
            Case "A4"
                qr1!iLivelloServizio = LS(5, 4)
            Case "B1"
                qr1!iLivelloServizio = LS(5, 5)
            Case "B2"
                qr1!iLivelloServizio = LS(5, 6)
            Case "B3"
                qr1!iLivelloServizio = LS(5, 7)
            Case "C1"
                qr1!iLivelloServizio = LS(5, 8)
            Case "C2"
                qr1!iLivelloServizio = LS(5, 9)
            End Select
        Case "Medium-Fast"
            Select Case brs![AbcConsumoValoreLs]
            Case "A1"
                qr1!iLivelloServizio = LS(6, 1)
            Case "A2"
                qr1!iLivelloServizio = LS(6, 2)
            Case "A3"
                qr1!iLivelloServizio = LS(6, 3)
            Case "A4"
                qr1!iLivelloServizio = LS(6, 4)
            Case "B1"
                qr1!iLivelloServizio = LS(6, 5)
            Case "B2"
                qr1!iLivelloServizio = LS(6, 6)
            Case "B3"
                qr1!iLivelloServizio = LS(6, 7)
            Case "C1"
                qr1!iLivelloServizio = LS(6, 8)
            Case "C2"
                qr1!iLivelloServizio = LS(6, 9)
            End Select
        Case "Fast"
            Select Case brs![AbcConsumoValoreLs]
            Case "A1"
                qr1!iLivelloServizio = LS(7, 1)
            Case "A2"
                qr1!iLivelloServizio = LS(7, 2)
            Case "A3"
                qr1!iLivelloServizio = LS(7, 3)
            Case "A4"
                qr1!iLivelloServizio = LS(7, 4)
            Case "B1"
                qr1!iLivelloServizio = LS(7, 5)
            Case "B2"
                qr1!iLivelloServizio = LS(7, 6)
            Case "B3"
                qr1!iLivelloServizio = LS(7, 7)
            Case "C1"
                qr1!iLivelloServizio = LS(7, 8)
            Case "C2"
                qr1!iLivelloServizio = LS(7, 9)
            End Select
        Case "Very-Fast"
            Select Case brs![AbcConsumoValoreLs]
            Case "A1"
                qr1!iLivelloServizio = LS(8, 1)
            Case "A2"
                qr1!iLivelloServizio = LS(8, 2)
            Case "A3"
                qr1!iLivelloServizio = LS(8, 3)
            Case "A4"
                qr1!iLivelloServizio = LS(8, 4)
            Case "B1"
                qr1!iLivelloServizio = LS(8, 5)
            Case "B2"
                qr1!iLivelloServizio = LS(8, 6)
            Case "B3"
                qr1!iLivelloServizio = LS(8, 7)
            Case "C1"
                qr1!iLivelloServizio = LS(8, 8)
            Case "C2"
                qr1!iLivelloServizio = LS(8, 9)
            End Select
        Case Else
            qr1!iLivelloServizio = 0
            qr1.Execute
        End Select
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    'start:

    ' ****************************************************************************************
    ' **** Maggiorazione Livello Servizio per articoli con stato = STRATEGICO (9)
    ' ****************************************************************************************
    ' DoCmd.RunSQL strSQL
    DoCmd.SetWarnings False
    strSQL = "UPDATE tblArticoli INNER JOIN tblArticoliStato ON tblArticoli.Cod_art = tblArticoliStato.Cod_Art"
    strSQL = strSQL & "   SET tblArticoli.LivelloServizio = 85 "
    strSQL = strSQL & " WHERE tblArticoliStato.ID_StatoArticolo = 9 "
    strSQL = strSQL & "   AND tblArticoli.Classe_Evento <> 'Very-Slow' "
    strSQL = strSQL & "   AND tblArticoli.LivelloServizio < 85;"
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
' TEST DA QUI
start:
    ' ****************************************************************************************
    ' **** AGGIORNO ROP e ROQ in tblArticoli
    ' ****************************************************************************************

    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("AGGIORNO ROP e ROQ ", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloRopRoq")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art"), brs.Fields("SConsumo"), brs.Fields("SSpedito"), brs.Fields("Num_Eventi")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryArticoliRopRoqUpdate")
        If brs![StatoArticolo] = 1 Or brs![StatoArticolo] = 4 Then    ' Se OBSOLETO o SOSTITUTIVO Rop e ROQ = 0
            qr1!iCOD_ART = brs![Cod_art]
            qr1!iPunto_riordino = 0
            qr1!iScortaSicurezza = brs![ScortaSicurezzaForzata]
            ' 20151207 Modificato ROQ was qr1!iLotto_ec_acq = 0
            ' Calcolo Lotto Acq --
            If brs![Lotto_min] > 0 Then
                qr1!iLotto_ec_acq = MaxOfList(brs![Lotto_min], brs![ScortaSicurezzaForzata]) ' 20160118
            Else
                If IsNull(brs![Lotto_multiplo]) Then
                    qr1!iLotto_ec_acq = brs![ScortaSicurezzaForzata]
                Else
                    qr1!iLotto_ec_acq = RoundUp(brs![ScortaSicurezzaForzata], brs![Lotto_multiplo])
                End If
            End If
            qr1!iInd_Rotaz = brs![Ind_Rotaz]
            qr1!iCopertura = brs![Copertura]
        Else
            Select Case brs![ScortaSicurezzaForzata]    ' Se scorta sicurezza Forzata >= 0
                ' Modificato come da mail Landriscina 10/12/12
            Case Is = 0    'Se SSF = 0 allora tutto = 0
                qr1!iCOD_ART = brs![Cod_art]
                qr1!iPunto_riordino = 0
                qr1!iScortaSicurezza = 0
                qr1!iLotto_ec_acq = 0
                qr1!iInd_Rotaz = brs![Ind_Rotaz]
                qr1!iCopertura = brs![Copertura]
            Case Is >= brs![ROP]    ' Modificato 13/02/2014 se SSF > ROP --> ROP = SSF
                qr1!iCOD_ART = brs![Cod_art]
                qr1!iPunto_riordino = 0
                qr1!iScortaSicurezza = brs![ScortaSicurezzaForzata]
                ' Calcolo Lotto Acq -- Marco Morotti 13-01-2015
                If brs![Lotto_min] > 0 And brs![Lotto_min] > brs![ScortaSicurezzaForzata] Then ' 14/12/2015
                    qr1!iLotto_ec_acq = MaxOfList(brs![Lotto_min], brs![ScortaSicurezzaForzata])
                Else
                    qr1!iLotto_ec_acq = brs![ScortaSicurezzaForzata]
                End If
                qr1!iInd_Rotaz = brs![Ind_Rotaz]
                qr1!iCopertura = brs![Copertura]
            Case Is < brs![ROP]    ' Modificato 13/02/2014 se
                qr1!iCOD_ART = brs![Cod_art]
                qr1!iPunto_riordino = brs![ROP]
                qr1!iScortaSicurezza = brs![ScortaSicurezza]
                ' Calcolo Lotto Acq -- Marco Morotti 13-01-2015
                If brs![Lotto_min] > brs![ROQ] Then
                    qr1!iLotto_ec_acq = MaxOfList(brs![Lotto_min], brs![ROP] + brs![ScortaSicurezza]) ' 20160118
                Else
                    If IsNull(brs![Lotto_multiplo]) Then
                        qr1!iLotto_ec_acq = MaxOfList(brs![ROQ], brs![ROP] + brs![ScortaSicurezza]) ' 20160118
                    Else
                        qr1!iLotto_ec_acq = MaxOfList(RoundUp(brs![ROQ], brs![Lotto_multiplo]), brs![ROP] + brs![ScortaSicurezza]) ' 20160118
                    End If
                End If
                qr1!iInd_Rotaz = brs![Ind_Rotaz]
                qr1!iCopertura = brs![Copertura]
            Case Else
                Select Case brs![LivelloServizio]
                Case Is > 0
                    qr1!iCOD_ART = brs![Cod_art]
                    qr1!iPunto_riordino = brs![ROP]
                    qr1!iScortaSicurezza = brs![ScortaSicurezza]
                    ' Calcolo Lotto Acq -- Marco Morotti 13-01-2015
                    If brs![Lotto_min] > brs![ROQ] Then
                        qr1!iLotto_ec_acq = MaxOfList(brs![Lotto_min], brs![ROP] + brs![ScortaSicurezza]) ' 20160118
                    Else
                        If IsNull(brs![Lotto_multiplo]) Then
                            qr1!iLotto_ec_acq = MaxOfList(brs![ROQ], brs![ROP] + brs![ScortaSicurezza]) ' 20160118
                        Else
                            qr1!iLotto_ec_acq = MaxOfList(RoundUp(brs![ROQ], brs![Lotto_multiplo]), brs![ROP] + brs![ScortaSicurezza]) ' 20160118
                        End If
                    End If
                    qr1!iInd_Rotaz = brs![Ind_Rotaz]
                    qr1!iCopertura = brs![Copertura]
                Case Else
                    qr1!iCOD_ART = brs![Cod_art]
                    qr1!iPunto_riordino = 0
                    qr1!iScortaSicurezza = brs![ScortaSicurezza]    ' Modificato 16/07/13 Morotti verificare articolo 0000107049C
                    ' 20151207 Modificato ROQ was qr1!iLotto_ec_acq = 0
                    ' Calcolo Lotto Acq --
                    If brs![Lotto_min] > 0 Then
                        qr1!iLotto_ec_acq = MaxOfList(brs![Lotto_min], brs![ScortaSicurezza]) ' 20160118
                    Else
                        If IsNull(brs![Lotto_multiplo]) Then
                            qr1!iLotto_ec_acq = brs![ScortaSicurezza]
                        Else
                            qr1!iLotto_ec_acq = RoundUp(brs![ScortaSicurezza], brs![Lotto_multiplo])
                        End If
                    End If

                    qr1!iInd_Rotaz = brs![Ind_Rotaz]
                    qr1!iCopertura = brs![Copertura]
                End Select
            End Select
        End If
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    'Start:
    ' ******************************************************************
    ' Secondo Giro per Tutti gli articoli che hanno SSF >= 0 e Costo = 0
    ' ******************************************************************
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("AGGIORNO SSF ", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloRopRoqSSF")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art"), brs.Fields("SConsumo"), brs.Fields("SSpedito"), brs.Fields("Num_Eventi")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryArticoliRopRoqUpdate")
        Select Case brs![ScortaSicurezzaForzata]    ' Se scorta sicurezza Forzata >= 0
            ' Modificato come da mail Landriscina 10/12/12
        Case Is = 0
            qr1!iCOD_ART = brs![Cod_art]
            qr1!iPunto_riordino = 0
            qr1!iScortaSicurezza = 0
            qr1!iLotto_ec_acq = 0
            qr1!iInd_Rotaz = brs![Ind_Rotaz]
            qr1!iCopertura = brs![Copertura]
        Case Is > 0
            qr1!iCOD_ART = brs![Cod_art]
            qr1!iPunto_riordino = 0
            qr1!iScortaSicurezza = brs![ScortaSicurezzaForzata]
            qr1!iLotto_ec_acq = MaxOfList(brs![ROQ], brs![ScortaSicurezzaForzata])
            qr1!iInd_Rotaz = brs![Ind_Rotaz]
            qr1!iCopertura = brs![Copertura]
        Case Else
            Select Case brs![LivelloServizio]
            Case Is > 0
                qr1!iCOD_ART = brs![Cod_art]
                qr1!iPunto_riordino = brs![ROP]
                qr1!iScortaSicurezza = brs![ScortaSicurezza]
                qr1!iLotto_ec_acq = MaxOfList(brs![ROQ], brs![ROP] + brs![ScortaSicurezza]) ' 20160118
                qr1!iInd_Rotaz = brs![Ind_Rotaz]
                qr1!iCopertura = brs![Copertura]
            Case Else
                qr1!iCOD_ART = brs![Cod_art]
                qr1!iPunto_riordino = 0
                qr1!iScortaSicurezza = brs![ScortaSicurezza]    ' Modificato 16/07/13 Morotti verificare articolo 0000107049C
                qr1!iLotto_ec_acq = brs![ScortaSicurezza]       ' Modificato 14/12/2015 era = 0
                qr1!iInd_Rotaz = brs![Ind_Rotaz]
                qr1!iCopertura = brs![Copertura]
            End Select
        End Select

        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
'start:
    ' *******************************************************************************
    ' Terzo Giro per Tutti gli articoli che hanno SSF >= 0 e Costo = 0 e Consumo = 0
    ' *******************************************************************************

    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("AGG. ROP e ROQ SSF-Consumo 0", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloRopRoqSSF_Cons0")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art"), brs.Fields("SConsumo"), brs.Fields("SSpedito"), brs.Fields("Num_Eventi")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryArticoliRopRoqUpdate")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iPunto_riordino = 0
        qr1!iScortaSicurezza = brs![ScortaSicurezzaForzata]
        qr1!iLotto_ec_acq = brs![ScortaSicurezzaForzata]
        qr1!iInd_Rotaz = 0
        qr1!iCopertura = 0
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    'DoCmd.SetWarnings True
    'DoCmd.Hourglass False
End Sub