dbMemo "SQL" ="SELECT tblImportSpedito.COD_ART, sum(QTA_OUT) AS Consumo, Count( * ) AS N_Spedit"
    "o_Mese\015\012FROM tblImportSpedito\015\012WHERE (((tblImportSpedito.DATABOLLA) "
    "Between [DataInizio] And [DataFine]))\015\012GROUP BY tblImportSpedito.COD_ART;\015"
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
        dbText "Name" ="tblImportSpedito.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N_Spedito_Mese"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2790"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="Consumo"
        dbLong "AggregateType" ="-1"
    End
End
