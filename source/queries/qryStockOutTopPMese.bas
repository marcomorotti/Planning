dbMemo "SQL" ="SELECT format([qryStockOut.DATA_ORDINE],\"yyyy mm\") AS Mese, Count(*) AS TotSto"
    "ckOutTopP\015\012FROM qryStockOut\015\012WHERE (((qryStockOut.Evadibile)='StockO"
    "ut'))\015\012GROUP BY format([qryStockOut.DATA_ORDINE],\"yyyy mm\"), qryStockOut"
    ".liv_urgenza\015\012HAVING (((qryStockOut.liv_urgenza)=2));\015\012"
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
        dbText "Name" ="Mese"
        dbLong "AggregateType" ="-1"
    End
End
