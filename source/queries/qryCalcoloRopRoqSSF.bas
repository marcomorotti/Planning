dbMemo "SQL" ="SELECT tblArticoli.cod_art, tblArticoli.Cs_Csc, tblArticoli.LivelloServizio, tbl"
    "Articoli.Lead_time, tblArticoli.AvgConsumoMese, tblArticoli.DevStdConsumoMese, R"
    "OP([Lead_time],[AvgConsumoMese]) AS Rop, ScortaSicurezza([LivelloServizio],[Lead"
    "_time],[AvgConsumoMese],[DevStdConsumoMese]) AS ScortaSicurezza, ROQ([MesiCopert"
    "ura],[AvgConsumoMese],[Cs_Csc]) AS Roq, IIf(Giac_Media>0,Round([SConsumo_12]/[Gi"
    "ac_Media],2),0) AS Ind_Rotaz, Round([Giac_Media]/[AvgConsumoMese],2) AS Copertur"
    "a, tblArticoliStato.ScortaSicurezzaForzata, tblArticoliStato.ID_StatoArticolo AS"
    " StatoArticolo\015\012FROM tblArticoli LEFT JOIN tblArticoliStato ON tblArticoli"
    ".Cod_art=tblArticoliStato.Cod_Art\015\012WHERE (((tblArticoli.Cs_Csc)=0) And tbl"
    "ArticoliStato.ScortaSicurezzaForzata>=0);\015\012"
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
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbInteger "ColumnWidth" ="2880"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="StatoArticolo"
        dbLong "AggregateType" ="-1"
    End
End
