dbMemo "SQL" ="INSERT INTO tblArticoliSpeditiUnique\015\012SELECT DISTINCT tblImportSpedito.COD"
    "_ART AS COD_ART, tblImportSpedito.DESCRIZIONE AS DESCRIZIONE, tblImportSpedito.L"
    "EAD_TIME AS LEAD_TIME, tblImportSpedito.ROP AS ROP, tblImportSpedito.MAX_MINMAX_"
    "QUANTITY AS MAX_MINMAX_QUANTITY, tblImportSpedito.ROQ AS ROQ, tblImportSpedito.M"
    "AXIMUM_ORDER_QUANTITY AS MAXIMUM_ORDER_QUANTITY, tblImportSpedito.CS_CSC AS CS_C"
    "SC, tblImportSpedito.CATEG_MERC AS CATEG_MERC, tblImportSpedito.STATO AS STATO, "
    "tblImportSpedito.PESO_NETTO AS PESO_NETTO, tblImportSpedito.PESO_LORDO AS PESO_L"
    "ORDO\015\012FROM tblImportSpedito;\015\012"
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
        dbText "Name" ="tblImportSpedito.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpedito.DESCRIZIONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpedito.LEAD_TIME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpedito.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpedito.MAX_MINMAX_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpedito.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpedito.MAXIMUM_ORDER_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpedito.CS_CSC"
        dbLong "AggregateType" ="-1"
    End
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
