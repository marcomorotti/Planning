dbMemo "SQL" ="SELECT tblCOrders.NUMERO_DOC AS N_OrdC, tblPOrders.NUMERO_DOC AS N_OrdP, tblCOrd"
    "ers.COD_ART, Count(tblPOrders.NUMERO_DOC) AS Riga\015\012FROM tblCOrders INNER J"
    "OIN tblPOrders ON tblCOrders.COD_ART=tblPOrders.COD_ART\015\012GROUP BY tblCOrde"
    "rs.NUMERO_DOC, tblPOrders.NUMERO_DOC, tblCOrders.COD_ART\015\012ORDER BY tblCOrd"
    "ers.NUMERO_DOC;\015\012"
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
        dbText "Name" ="tblCOrders.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N_OrdC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N_OrdP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Riga"
        dbLong "AggregateType" ="-1"
    End
End
