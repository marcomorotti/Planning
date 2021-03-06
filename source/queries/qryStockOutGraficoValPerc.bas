﻿dbMemo "SQL" ="SELECT qryCOrdersTotaleTopP.DATA_ORDINE, qryCOrdersTotale.Totale, qryCOrdersTota"
    "leTopP.[TotaleTopP] AS TopPriority, [TotStockOutTopP]/[Totale] AS StockOutTopPri"
    "ority, [TotStockOutAll]/[Totale] AS StockOutAll\015\012FROM qryCOrdersTotale INN"
    "ER JOIN ((qryCOrdersTotaleTopP INNER JOIN qryStockOutTopP ON qryCOrdersTotaleTop"
    "P.DATA_ORDINE=qryStockOutTopP.DATA_ORDINE) INNER JOIN qryStockOutAll ON qryStock"
    "OutTopP.DATA_ORDINE=qryStockOutAll.DATA_ORDINE) ON qryCOrdersTotale.DATA_ORDINE="
    "qryCOrdersTotaleTopP.DATA_ORDINE;\015\012"
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
