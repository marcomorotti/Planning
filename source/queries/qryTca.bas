Operation =1
Option =0
Begin InputTables
    Name ="tblArticoli"
    Name ="qryStatoArticoli"
End
Begin OutputColumns
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.Des_art"
    Expression ="qryStatoArticoli.Stato"
    Alias ="RopAct"
    Expression ="tblArticoli.ROP"
    Alias ="RopProp"
    Expression ="[Punto_riordino]+[ScortaSicurezza]"
    Expression ="qryStatoArticoli.ScortaSicurezzaForzata"
    Alias ="RoqAct"
    Expression ="tblArticoli.ROQ"
    Alias ="RoqProp"
    Expression ="tblArticoli.Lotto_ec_acq"
    Alias ="TcaAct"
    Expression ="IIf(tblArticoli.ROQ>0,((tblArticoli.ROQ/2+tblArticoli.ScortaSicurezza)*(tblArtic"
        "oli.Cs_Csc*0.68*0.21)+((tblArticoli.SConsumo_12/tblArticoli.ROQ)*20)),0)"
    Alias ="TcaProp"
    Expression ="IIf(tblArticoli.Lotto_ec_acq>0,((tblArticoli.Lotto_ec_acq/2+tblArticoli.ScortaSi"
        "curezza)*(tblArticoli.Cs_Csc*0.68*0.21)+((tblArticoli.SConsumo_12/tblArticoli.Lo"
        "tto_ec_acq)*20)),0)"
    Alias ="Classe_Movi"
    Expression ="tblArticoli.Classe_Evento"
    Alias ="Giacenza"
    Expression ="tblArticoli.Giac_Media"
    Alias ="ConsAnnuo"
    Expression ="tblArticoli.SConsumo_12"
    Expression ="tblArticoli.AbcGiacenza"
    Expression ="tblArticoli.AbcConsumo"
    Expression ="tblArticoli.AbcConsumoValoreLs"
    Expression ="tblArticoli.Cs_Csc"
End
Begin Joins
    LeftTable ="tblArticoli"
    RightTable ="qryStatoArticoli"
    Expression ="tblArticoli.Cod_art=qryStatoArticoli.Cod_Art"
    Flag =2
End
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
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcConsumoValoreLs"
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
        dbText "Name" ="qryStatoArticoli.ScortaSicurezzaForzata"
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
        dbText "Name" ="Classe_Movi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Giacenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TcaAct"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TcaProp"
        dbInteger "ColumnWidth" ="1815"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcGiacenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ConsAnnuo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcConsumo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =52
    Top =224
    Right =1363
    Bottom =932
    Left =-1
    Top =-1
    Right =1279
    Bottom =218
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =242
        Bottom =231
        Top =0
        Name ="tblArticoli"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="qryStatoArticoli"
        Name =""
    End
End
