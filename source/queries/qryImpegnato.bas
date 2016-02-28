dbMemo "SQL" ="SELECT tblCOrders.COD_ART, sum(tblCOrders.qta_ord_umv) AS Impegnato\015\012FROM "
    "tblCOrders\015\012GROUP BY tblCOrders.COD_ART;\015\012"
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
        dbText "Name" ="tblCOrders.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Impegnato"
        dbLong "AggregateType" ="-1"
    End
End
