dbMemo "SQL" ="SELECT qryStockOutMng.DATA_ORDINE, Count(*) AS TotStockOutTopP\015\012FROM qrySt"
    "ockOutMng\015\012WHERE (((qryStockOutMng.Evadibile)='StockOut'))\015\012GROUP BY"
    " qryStockOutMng.DATA_ORDINE, qryStockOutMng.liv_urgenza\015\012HAVING (((qryStoc"
    "kOutMng.liv_urgenza)=2));\015\012"
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
        dbText "Name" ="TotStockOutTopP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStockOutMng.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
End
