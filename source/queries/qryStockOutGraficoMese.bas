dbMemo "SQL" ="SELECT qryCOrdersTotaleMese.Mese, qryCOrdersTotaleMese.Totale, [TotaleTopP]/[Tot"
    "ale] AS TopPriority, [TotStockOutTopP]/[Totale] AS StockOutTopPriority, [TotStoc"
    "kOutAll]/[Totale] AS StockOutAll\015\012FROM ((qryCOrdersTotaleMese INNER JOIN q"
    "ryCOrdersTotaleTopPMese ON qryCOrdersTotaleMese.Mese=qryCOrdersTotaleTopPMese.Me"
    "se) INNER JOIN qryStockOutTopPMese ON qryCOrdersTotaleTopPMese.Mese=qryStockOutT"
    "opPMese.Mese) INNER JOIN qryStockOutAllMese ON qryStockOutTopPMese.Mese=qryStock"
    "OutAllMese.Mese;\015\012"
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
        dbText "Name" ="TopPriority"
        dbInteger "ColumnWidth" ="1860"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Percent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StockOutTopPriority"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Percent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StockOutAll"
        dbInteger "ColumnWidth" ="2115"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Percent"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCOrdersTotaleMese.Mese"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCOrdersTotaleMese.Totale"
        dbLong "AggregateType" ="-1"
    End
End
