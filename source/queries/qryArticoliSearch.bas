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
    Expression ="tblArticoli.Classe_Evento"
    Expression ="tblArticoli.LivelloServizio"
    Expression ="tblArticoli.Cs_Csc"
    Expression ="tblArticoli.ClasseCosto"
    Expression ="tblArticoli.Lead_time"
    Expression ="tblArticoli.ROP"
    Expression ="tblArticoli.Punto_riordino"
    Expression ="tblArticoli.ROQ"
    Expression ="tblArticoli.Lotto_ec_acq"
    Expression ="qryStatoArticoli.ScortaSicurezzaForzata"
    Expression ="tblArticoli.Giac_Media"
    Expression ="tblArticoli.Copertura"
    Expression ="tblArticoli.SConsumo"
    Expression ="tblArticoli.SSpedito"
    Expression ="tblArticoli.SConsumo_12"
    Expression ="tblArticoli.SSpedito_12"
    Expression ="tblArticoli.AbcGiacenza"
    Expression ="tblArticoli.AbcConsumo"
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
        dbText "Name" ="tblArticoli.Lead_time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Punto_riordino"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lotto_ec_acq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStatoArticoli.Stato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Copertura"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SConsumo_12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SSpedito_12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.LivelloServizio"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ClasseCosto"
        dbLong "AggregateType" ="-1"
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
        dbText "Name" ="tblArticoli.SConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SSpedito"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryStatoArticoli.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-7
    Top =146
    Right =947
    Bottom =665
    Left =-1
    Top =-1
    Right =922
    Bottom =190
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblArticoli"
        Name =""
    End
    Begin
        Left =269
        Top =10
        Right =413
        Bottom =154
        Top =0
        Name ="qryStatoArticoli"
        Name =""
    End
End
