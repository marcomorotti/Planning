Option Compare Database
Function CalcGiacenzaMedia()
    Dim CalGiac As New ADODB.Recordset
    Dim Art As New ADODB.Recordset
    Set Db = CurrentProject.Connection
    'lancio la query di calcolo media giacenze
    CalGiac.Open "qryGiacMedia", Db, adOpenKeyset, adLockOptimistic
    If Not (CalGiac.EOF And CalGiac.BOF) Then
        Do While Not CalGiac.EOF
            'memorizzo i risultati nella tabella anagrafica articoli
            Art.Open "SELECT * FROM tblArticoli WHERE Cod_Art='" + CalGiac.Fields("Cod_Art") + "'", Db, adOpenKeyset, adLockOptimistic
            If Art.EOF And Art.BOF Then
            Else
                Art.Fields("GiacenzaMediaMese") = CalGiac.Fields("StockMedio")
                '   Art.Fields("Mesi_giacenze") = DatiGen.Fields("Mesi_giacenze")
                Art.Update
            End If
            Art.Close
            CalGiac.MoveNext
        Loop
    End If
End Function