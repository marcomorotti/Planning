dbMemo "SQL" ="SELECT tblArticoliCOrdersUnique.COD_ART, tblArticoliCOrdersUnique.DESCRIZIONE, t"
    "blArticoliCOrdersUnique.LEAD_TIME, tblArticoliCOrdersUnique.ROP, tblArticoliCOrd"
    "ersUnique.MAX_MINMAX_QUANTITY, tblArticoliCOrdersUnique.ROQ, tblArticoliCOrdersU"
    "nique.MAXIMUM_ORDER_QUANTITY, tblArticoliCOrdersUnique.CS_CSC\015\012FROM tblArt"
    "icoliCOrdersUnique\015\012WHERE ( ( (EXISTS (SELECT Cod_Art\015\012             "
    "         FROM tblArticoli\015\012                     WHERE tblArticoli.Cod_art "
    "= tblArticoliCOrdersUnique.COD_ART)) =\015\012             FALSE));\015\012"
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
        dbText "Name" ="tblArticoliCOrdersUnique.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliCOrdersUnique.DESCRIZIONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliCOrdersUnique.LEAD_TIME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliCOrdersUnique.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliCOrdersUnique.MAX_MINMAX_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliCOrdersUnique.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliCOrdersUnique.MAXIMUM_ORDER_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliCOrdersUnique.CS_CSC"
        dbLong "AggregateType" ="-1"
    End
End
