dbMemo "SQL" ="SELECT tblArticoli.Cod_art, tblArticoli.Des_art, qryStatoArticoli.Stato, tblArti"
    "coli.LivelloServizio, tblArticoli.ROP AS RopAct, iif(isnull([Punto_riordino] + ["
    "ScortaSicurezza]), 0,   [Punto_riordino] + [ScortaSicurezza]) AS RopProp, qrySta"
    "toArticoli.ScortaSicurezzaForzata, tblArticoli.ROQ AS RoqAct, iif(ISNULL(tblArti"
    "coli.Lotto_ec_acq), 0, tblArticoli.Lotto_ec_acq) AS RoqProp, tblArticoli.AbcCons"
    "umoValoreLs, tblArticoli.Classe_Evento, tblArticoli.Giac_Media, tblArticoli.SCon"
    "sumo_12, tblArticoli.Cs_Csc, tblArticoli.AbcGiacenza, tblArticoli.AbcConsumo, tb"
    "lArticoli.Categ_Merc\015\012FROM tblArticoli LEFT JOIN qryStatoArticoli ON tblAr"
    "ticoli.Cod_art = qryStatoArticoli.Cod_Art\015\012WHERE tblArticoli.Categ_Merc no"
    "t in ('L70204', 'L70205');\015\012"
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
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStatoArticoli.Stato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RopAct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RopProp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RoqAct"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RoqProp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcConsumoValoreLs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStatoArticoli.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SConsumo_12"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1536"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcGiacenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.LivelloServizio"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Categ_Merc"
        dbLong "AggregateType" ="-1"
    End
End
