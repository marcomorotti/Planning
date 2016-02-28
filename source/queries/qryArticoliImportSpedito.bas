dbMemo "SQL" ="SELECT tblArticoliSpeditiUnique.COD_ART, tblArticoliSpeditiUnique.DESCRIZIONE, t"
    "blArticoliSpeditiUnique.LEAD_TIME, tblArticoliSpeditiUnique.ROP, tblArticoliSped"
    "itiUnique.MAX_MINMAX_QUANTITY, tblArticoliSpeditiUnique.ROQ, tblArticoliSpeditiU"
    "nique.MAXIMUM_ORDER_QUANTITY, tblArticoliSpeditiUnique.CS_CSC\015\012FROM tblArt"
    "icoliSpeditiUnique\015\012WHERE (((Exists (select Cod_Art\015\012from tblArticol"
    "i\015\012where tblArticoli.Cod_art = tblArticoliSpeditiUnique.COD_ART))=False));"
    "\015\012"
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
        dbText "Name" ="tblArticoliSpeditiUnique.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliSpeditiUnique.DESCRIZIONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliSpeditiUnique.LEAD_TIME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliSpeditiUnique.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliSpeditiUnique.MAX_MINMAX_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliSpeditiUnique.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliSpeditiUnique.MAXIMUM_ORDER_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliSpeditiUnique.CS_CSC"
        dbLong "AggregateType" ="-1"
    End
End
