dbMemo "SQL" ="SELECT qryStockOut.Evadibile, qryStockOut.liv_urgenza\015\012FROM qryStockOut LE"
    "FT JOIN (tblArticoli LEFT JOIN tblArticoliStato ON tblArticoli.Cod_art=tblArtico"
    "liStato.Cod_Art) ON qryStockOut.COD_ART=tblArticoli.Cod_art\015\012WHERE (((qryS"
    "tockOut.Evadibile)=\"StockOut\") AND ((qryStockOut.liv_urgenza)=2));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryStockOut.Evadibile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStockOut.liv_urgenza"
        dbLong "AggregateType" ="-1"
    End
End
