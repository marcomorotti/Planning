dbMemo "SQL" ="SELECT qryStockOutMng.DATA_ORDINE, Count(*) AS TotStockOutAll\015\012FROM qrySto"
    "ckOutMng\015\012WHERE (((qryStockOutMng.Evadibile)='StockOut'))\015\012GROUP BY "
    "qryStockOutMng.DATA_ORDINE;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="240"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TotStockOutAll"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStockOutMng.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
End
