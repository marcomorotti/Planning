dbMemo "SQL" ="SELECT tblImportSpedito.COD_ART, tblImportSpedito.DESCRIZIONE, tblImportSpedito."
    "LEAD_TIME, tblImportSpedito.ROP, tblImportSpedito.MAX_MINMAX_QUANTITY, tblImport"
    "Spedito.ROQ, tblImportSpedito.MAXIMUM_ORDER_QUANTITY, tblImportSpedito.CS_CSC\015"
    "\012FROM tblImportSpedito\015\012UNION SELECT tblImportGiacenza.CD_ART, tblImpor"
    "tGiacenza.DESCR_ART, tblImportGiacenza.LEAD_TIME, tblImportGiacenza.ROP, tblImpo"
    "rtGiacenza.MAX_MINMAX_QUANTITY, tblImportGiacenza.ROQ, tblImportGiacenza.MAXIMUM"
    "_ORDER_QUANTITY, tblImportGiacenza.CS_CSC\015\012FROM tblImportGiacenza;\015\012"
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
        dbText "Name" ="tblImportSpedito.LEAD_TIME"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpedito.ROP"
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
        dbText "Name" ="tblImportSpedito.DESCRIZIONE"
        dbInteger "ColumnWidth" ="5550"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpedito.MAX_MINMAX_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
End
