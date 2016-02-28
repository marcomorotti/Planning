dbMemo "SQL" ="SELECT tblArticoli.cod_art, tblArticoli.Lead_time, tblArticoli.AvgConsumoMese, t"
    "blArticoli.devStdConsumoMese, tblArticoli.cs_csc, ScortaSicurezza(99.99,tblArtic"
    "oli.Lead_time,tblArticoli.AvgConsumoMese,tblArticoli.devStdConsumoMese) AS Scort"
    "aS\015\012FROM tblArticoli\015\012WHERE (((tblArticoli.Classe_Evento)=\"Very-Fas"
    "t\") AND ((tblArticoli.AbcConsumoValoreLs)=\"C1\"));\015\012"
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
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lead_time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AvgConsumoMese"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1920"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblArticoli.devStdConsumoMese"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScortaS"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.cs_csc"
        dbLong "AggregateType" ="-1"
    End
End
