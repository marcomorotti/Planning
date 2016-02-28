dbMemo "SQL" ="SELECT tblConsumiParetoValoreLs.Cod_art AS Cod_Art, tblConsumiParetoValoreLs.Sco"
    "nsumoValore AS TotaleConsumo, Round(DSum(\"[SconsumoValore]\",\"tblConsumiPareto"
    "ValoreLs\",\"[SconsumoValore]>= \" & [TotaleConsumo] & \"\")/qryConsumoParetoLs1"
    "00.TotaleConsumoValore,2) AS CumPct\015\012FROM tblConsumiParetoValoreLs, qryCon"
    "sumoParetoLs100\015\012ORDER BY tblConsumiParetoValoreLs.SconsumoValore DESC;\015"
    "\012"
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
