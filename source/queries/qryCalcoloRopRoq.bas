dbMemo "SQL" ="SELECT tblArticoli.cod_art, tblArticoli.Cs_Csc, tblArticoli.LivelloServizio, tbl"
    "Articoli.Lead_time, tblArticoli.AvgConsumoMese, tblArticoli.DevStdConsumoMese, R"
    "OP([Lead_time],[AvgConsumoMese]) AS Rop, ScortaSicurezza([LivelloServizio],[Lead"
    "_time],[AvgConsumoMese],[DevStdConsumoMese]) AS ScortaSicurezza, ROQ([MesiCopert"
    "ura],[AvgConsumoMese],[Cs_Csc]) AS Roq, IIf(GiacenzaMediaMese>0,Round([SConsumo_"
    "12]/[GiacenzaMediaMese],2),0) AS Ind_Rotaz, Round([Giac_Media]/[AvgConsumoMese],"
    "2) AS Copertura, tblArticoliStato.ScortaSicurezzaForzata, tblArticoliStato.ID_St"
    "atoArticolo AS StatoArticolo, tblArticoliStato.lotto_min, tblArticoliStato.lotto"
    "_multiplo\015\012FROM tblArticoli LEFT JOIN tblArticoliStato ON tblArticoli.Cod_"
    "art=tblArticoliStato.Cod_Art\015\012WHERE (((tblArticoli.AvgConsumoMese)>0) AND "
    "((tblArticoli.Cs_Csc)>0));\015\012"
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
        dbText "Name" ="tblArticoli.cod_art"
        dbInteger "ColumnWidth" ="1365"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Rop"
        dbInteger "ColumnWidth" ="780"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.LivelloServizio"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lead_time"
        dbInteger "ColumnWidth" ="1395"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AvgConsumoMese"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.DevStdConsumoMese"
        dbInteger "ColumnWidth" ="2460"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Roq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ind_Rotaz"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Copertura"
        dbInteger "ColumnWidth" ="1470"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScortaSicurezza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StatoArticolo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.lotto_min"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.lotto_multiplo"
        dbLong "AggregateType" ="-1"
    End
End
