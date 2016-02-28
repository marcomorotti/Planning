dbMemo "SQL" ="INSERT INTO tblArticoli\015\012SELECT tblImportGiacenza.CD_ART AS COD_ART, tblIm"
    "portGiacenza.DESCR_ART AS DES_ART, tblImportGiacenza.LEAD_TIME AS LEAD_TIME, tbl"
    "ImportGiacenza.ROP AS ROP, tblImportGiacenza.MAX_MINMAX_QUANTITY AS MAX_MINMAX_Q"
    "UANTITY, tblImportGiacenza.ROQ AS ROQ, tblImportGiacenza.MAXIMUM_ORDER_QUANTITY "
    "AS MAXIMUM_ORDER_QUANTITY, tblImportGiacenza.CS_CSC AS CS_CSC, tblImportGiacenza"
    ".QT_GIAC AS GIAC_MEDIA, tblImportGiacenza.CATEG_MERC AS CATEG_MERC, tblImportGia"
    "cenza.STATO AS STATO, tblImportGiacenza.PESO_NETTO AS PESO_NETTO, tblImportGiace"
    "nza.PESO_LORDO AS PESO_LORDO\015\012FROM tblImportGiacenza;\015\012"
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
        dbText "Name" ="qryArtIsrt.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryArtIsrt.DESCRIZIONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Des_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryArtIsrt.CS_CSC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryArtIsrt.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryArtIsrt.LEAD_TIME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryArtIsrt.MAX_MINMAX_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryArtIsrt.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryArtIsrt.MAXIMUM_ORDER_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CS_CSC"
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
        dbText "Name" ="GIAC_MEDIA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CATEG_MER"
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
    Begin
        dbText "Name" ="CATEG_MERC"
    End
End
