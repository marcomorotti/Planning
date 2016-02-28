dbMemo "SQL" ="SELECT tblArticoli.Cod_art, tblArticoli.Des_art, tblArticoli.SConsumo, tblArtico"
    "li.SSpedito, tblArticoli.ROP, tblArticoli.Punto_riordino, tblArticoli.ROQ, tblAr"
    "ticoli.Lotto_ec_acq, tblArticoli.Classe_Evento\015\012FROM tblArticoli\015\012WH"
    "ERE (((tblArticoli.SConsumo)=0) AND ((tblArticoli.ROP)>0))\015\012ORDER BY tblAr"
    "ticoli.SConsumo DESC;\015\012"
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
dbMemo "Filter" ="([tblArticoli Query].[Classe_Evento] Not In (\"Slow\",\"Very-Fast\",\"Very-Slow\""
    "))"
Begin
    Begin
        dbText "Name" ="tblArticoli.[Cod_art]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.[SConsumo]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.[ROP]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SSpedito"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Punto_riordino"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lotto_ec_acq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROP"
        dbLong "AggregateType" ="-1"
    End
End
