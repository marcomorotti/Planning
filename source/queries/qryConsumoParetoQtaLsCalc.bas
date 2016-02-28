dbMemo "SQL" ="SELECT tblConsumiParetoQtaLs.Cod_art AS Cod_Art, tblConsumiParetoQtaLs.SconsumoQ"
    "ta AS TotaleConsumo, Round(DSum(\"[SconsumoQta]\",\"tblConsumiParetoQtaLs\",\"[S"
    "consumoQta]>= \" & [TotaleConsumo] & \"\")/qryConsumoParetoLs100.TotaleConsumoQt"
    "a,2) AS CumPct\015\012FROM tblConsumiParetoQtaLs, qryConsumoParetoLs100\015\012O"
    "RDER BY tblConsumiParetoQtaLs.SconsumoQta DESC;\015\012"
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
        dbText "Name" ="TotaleConsumo"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="CumPct"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cod_Art"
        dbLong "AggregateType" ="-1"
    End
End
