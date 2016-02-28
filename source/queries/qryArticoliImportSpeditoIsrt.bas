dbMemo "SQL" ="INSERT INTO tblArticoli\015\012SELECT tblArticoliSpeditiUnique.COD_ART AS COD_AR"
    "T, tblArticoliSpeditiUnique.DESCRIZIONE AS DES_ART, tblArticoliSpeditiUnique.LEA"
    "D_TIME AS LEAD_TIME, tblArticoliSpeditiUnique.ROP AS ROP, tblArticoliSpeditiUniq"
    "ue.MAX_MINMAX_QUANTITY AS MAX_MINMAX_QUANTITY, tblArticoliSpeditiUnique.ROQ AS R"
    "OQ, tblArticoliSpeditiUnique.MAXIMUM_ORDER_QUANTITY AS MAXIMUM_ORDER_QUANTITY, t"
    "blArticoliSpeditiUnique.CS_CSC AS CS_CSC, tblArticoliSpeditiUnique.CATEG_MERC AS"
    " CATEG_MERC, tblArticoliSpeditiUnique.STATO AS STATO, tblArticoliSpeditiUnique.P"
    "ESO_NETTO AS PESO_NETTO, tblArticoliSpeditiUnique.PESO_LORDO AS PESO_LORDO\015\012"
    "FROM tblArticoliSpeditiUnique\015\012WHERE (((Exists\015\012           (select C"
    "od_Art\015\012                from tblArticoli\015\012               where tblAr"
    "ticoli.Cod_art = tblArticoliSpeditiUnique.COD_ART)) =\015\012          False));\015"
    "\012"
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
        dbText "Name" ="COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DESCRIZIONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LEAD_TIME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MAX_MINMAX_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MAXIMUM_ORDER_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CS_CSC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DES_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CATEG_MERC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="STATO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PESO_NETTO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PESO_LORDO"
        dbLong "AggregateType" ="-1"
    End
End
