dbMemo "SQL" ="SELECT format([qryStockOut.DATA_ORDINE],\"yyyy mm\") AS Mese, Count(*) AS TotSto"
    "ckOutAll\015\012FROM qryStockOut\015\012WHERE (((qryStockOut.Evadibile)='StockOu"
    "t'))\015\012GROUP BY format([qryStockOut.DATA_ORDINE],\"yyyy mm\");\015\012"
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
        dbText "Name" ="Mese"
        dbLong "AggregateType" ="-1"
    End
End
