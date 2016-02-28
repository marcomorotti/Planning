dbMemo "SQL" ="SELECT format([tblCOrdersStorico.DATA_ORDINE],\"yyyy mm\") AS Mese, Count(tblCOr"
    "dersStorico.COD_ART) AS TotaleTopP\015\012FROM tblCOrdersStorico\015\012GROUP BY"
    " format([tblCOrdersStorico.DATA_ORDINE],\"yyyy mm\"), tblCOrdersStorico.liv_urge"
    "nza\015\012HAVING (((tblCOrdersStorico.liv_urgenza)=2));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="TotaleTopP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Mese"
        dbLong "AggregateType" ="-1"
    End
End
