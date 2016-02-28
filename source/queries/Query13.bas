dbMemo "SQL" ="SELECT tblArticoliStato.ID_ArticoliStato, tblArticoliStato.Cod_Art, tblArticoliS"
    "tato.ID_StatoArticolo, tblArticoliStato.DES_ART, tblArticoliStato.Note, tblArtic"
    "oliStato.Cod_Art_Correlato, tblArticoliStato.Data_Modifica, tblArticoliStato.Sco"
    "rtaSicurezzaForzata, tblArticoliStato.Lotto_min, tblArticoliStato.Lotto_multiplo"
    "\015\012FROM tblArticoliStato INNER JOIN tblDoppiStatoArticolo ON tblArticoliSta"
    "to.Cod_Art = tblDoppiStatoArticolo.Cod_art\015\012ORDER BY tblArticoliStato.Cod_"
    "Art, tblArticoliStato.ID_ArticoliStato;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblArticoliStato.ID_StatoArticolo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Cod_Art_Correlato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Data_Modifica"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Lotto_min"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Lotto_multiplo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ID_ArticoliStato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.DES_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Note"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
End
