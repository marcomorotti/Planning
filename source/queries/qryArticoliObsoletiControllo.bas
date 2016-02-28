dbMemo "SQL" ="SELECT tblArticoliStato.Cod_Art, tblArticoliObsoletiImport.TIPORELAZ, tblArticol"
    "iObsoletiImport.NEW_COD, tblArticoliStato.ScortaSicurezzaForzata\015\012FROM tbl"
    "ArticoliStato INNER JOIN tblArticoliObsoletiImport ON tblArticoliStato.Cod_Art=t"
    "blArticoliObsoletiImport.OLD_COD\015\012WHERE (((tblArticoliStato.Cod_Art_Correl"
    "ato) Is Null) AND ((tblArticoliObsoletiImport.TIPORELAZ)<>\"Correlato\"));\015\012"
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
        dbText "Name" ="tblArticoliStato.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliObsoletiImport.TIPORELAZ"
        dbInteger "ColumnWidth" ="3810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliObsoletiImport.NEW_COD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
End
