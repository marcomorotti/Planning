dbMemo "SQL" ="SELECT tblImportSpedito.COD_ART, sum(QTA_OUT) AS Consumo\015\012FROM tblImportSp"
    "edito\015\012WHERE (((tblImportSpedito.DATABOLLA) Between [DataInizio] And [Data"
    "Fine]))\015\012GROUP BY tblImportSpedito.COD_ART;\015\012"
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
        dbText "Name" ="Consumo"
        dbLong "AggregateType" ="-1"
    End
End
