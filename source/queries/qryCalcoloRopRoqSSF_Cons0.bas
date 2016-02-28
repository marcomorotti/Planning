dbMemo "SQL" ="SELECT tblArticoli.cod_art, tblArticoliStato.ScortaSicurezzaForzata\015\012FROM "
    "tblArticoli LEFT JOIN tblArticoliStato ON tblArticoli.Cod_art = tblArticoliStato"
    ".Cod_Art\015\012WHERE tblArticoliStato.ScortaSicurezzaForzata >= 0\015\012   and"
    "  tblArticoli.AvgConsumoMese is  null;\015\012"
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
        dbText "Name" ="tblArticoli.cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbInteger "ColumnWidth" ="3165"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
