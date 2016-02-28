dbMemo "SQL" ="SELECT qryCOrdersTotaleTopP.DATA_ORDINE, qryCOrdersTotale.Totale, [TotaleTopP]/["
    "Totale] AS TopPriority, [TotStockOutTopP]/[Totale] AS StockOutTopPriority, [TotS"
    "tockOutAll]/[Totale] AS StockOutAll\015\012FROM qryCOrdersTotale INNER JOIN ((qr"
    "yCOrdersTotaleTopP INNER JOIN qryStockOutTopP ON qryCOrdersTotaleTopP.DATA_ORDIN"
    "E=qryStockOutTopP.DATA_ORDINE) INNER JOIN qryStockOutAll ON qryStockOutTopP.DATA"
    "_ORDINE=qryStockOutAll.DATA_ORDINE) ON qryCOrdersTotale.DATA_ORDINE=qryCOrdersTo"
    "taleTopP.DATA_ORDINE;\015\012"
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
        dbText "Name" ="qryCOrdersTotaleTopP.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryCOrdersTotale.Totale"
        dbLong "AggregateType" ="-1"
    End
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
End
