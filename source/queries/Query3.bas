dbMemo "SQL" ="SELECT qryStockOutGraficoValPercMng.[DATA_ORDINE], qryStockOutGraficoValPercMng."
    "[Totale], (qryStockOutGraficoValPercMng.[TopPriority]) AS TopPriority, (qryStock"
    "OutGraficoValPercMng.[StockOutTopPriority]*100) AS StockOutTopPriority, (qryStoc"
    "kOutGraficoValPercMng.[StockOutAll]*100) AS StockOutAll\015\012FROM qryStockOutG"
    "raficoValPercMng\015\012WHERE qryStockOutGraficoValPercMng.DATA_ORDINE Between #"
    "7/1/2011# And #7/28/2011#;\015\012"
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
        dbText "Name" ="qryStockOutGraficoValPercMng.[DATA_ORDINE]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStockOutGraficoValPercMng.[Totale]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TopPriority"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StockOutTopPriority"
        dbInteger "ColumnWidth" ="1995"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StockOutAll"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
