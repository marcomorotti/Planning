dbMemo "SQL" ="SELECT tblArticoli.Classe_Evento, AbcConsumoValoreLs, sum(tblArticoliStato.Scort"
    "aSicurezzaForzata*tblArticoli.Cs_Csc) AS SSFValo\015\012FROM tblArticoli INNER J"
    "OIN tblArticoliStato ON tblArticoli.Cod_art = tblArticoliStato.Cod_Art\015\012WH"
    "ERE tblArticoliStato.ScortaSicurezzaForzata > 0\015\012GROUP BY tblArticoli.Clas"
    "se_Evento, AbcConsumoValoreLs;\015\012"
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
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AbcConsumoValoreLs"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SSFValo"
        dbLong "AggregateType" ="-1"
    End
End
