dbMemo "SQL" ="SELECT tblConsumi.COD_ART, Sum(tblConsumi.Consumo) AS SConsumo, Sum(tblConsumi.N"
    "_Spedito_Mese) AS SSpedito, Count(*) AS Num_Eventi, ((Avg([Consumo])/12)*[Num_Ev"
    "enti]) AS AvgConsumoMese\015\012FROM tblConsumi\015\012WHERE DateSerial([Anno],["
    "Mese],1)>=DateSerial([AnnoI],[MeseI],1) And DateSerial([Anno],[mese],1)<=DateSer"
    "ial([AnnoF],[MeseF],1)\015\012GROUP BY tblConsumi.COD_ART;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblConsumi.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Num_Eventi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SSpedito"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AvgConsumoMese"
        dbInteger "ColumnWidth" ="1890"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
