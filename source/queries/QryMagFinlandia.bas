dbMemo "SQL" ="SELECT MagFil.Number, MagFil.[Product in Finnish], MagFil.[Unit price], MagFil.q"
    "uantity, MagFil.[total value], tblArticoli.Cs_Csc, tblArticoli.Classe_Evento, tb"
    "lArticoli.SConsumo AS QCons36mesi, tblArticoli.SSpedito AS QSped36mesi, tblArtic"
    "oli.Num_Eventi AS N_Sped_36_Mesi, tblArticoli.SConsumo_12 AS QCons12mesi, tblArt"
    "icoli.SSpedito_12 AS QSped12mesi, tblArticoli.Num_Eventi_12 AS N_Sped_12_Mesi, t"
    "blArticoli.Giac_Media, tblArticoli.AvgConsumoMese, tblArticoli.Copertura, tblArt"
    "icoli.AbcGiacenza, tblArticoli.AbcConsumo\015\012FROM tblArticoli RIGHT JOIN Mag"
    "Fil ON tblArticoli.Cod_art=MagFil.Number;\015\012"
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
        dbText "Name" ="MagFil.Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcGiacenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MagFil.[total value]"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MagFil.[Unit price]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MagFil.quantity"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MagFil.[Product in Finnish]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Copertura"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AvgConsumoMese"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QCons36mesi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QSped36mesi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N_Sped_36_Mesi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QCons12mesi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N_Sped_12_Mesi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="QSped12mesi"
        dbLong "AggregateType" ="-1"
    End
End
