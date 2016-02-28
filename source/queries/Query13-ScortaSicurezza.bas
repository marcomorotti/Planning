Operation =1
Option =0
Where ="(((tblArticoliStato.ScortaSicurezzaForzata)>=0))"
Begin InputTables
    Name ="tblArticoli"
    Name ="tblArticoliStato"
End
Begin OutputColumns
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoliStato.ScortaSicurezzaForzata"
    Expression ="tblArticoli.Punto_riordino"
    Expression ="tblArticoli.Lotto_ec_acq"
    Expression ="tblArticoli.ROP"
    Expression ="tblArticoli.ROQ"
End
Begin Joins
    LeftTable ="tblArticoli"
    RightTable ="tblArticoliStato"
    Expression ="tblArticoli.Cod_art = tblArticoliStato.Cod_Art"
    Flag =1
End
Begin OrderBy
    Expression ="tblArticoliStato.ScortaSicurezzaForzata"
    Flag =0
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
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
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
        dbText "Name" ="tblArticoli.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROQ"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =30
    Top =100
    Right =951
    Bottom =777
    Left =-1
    Top =-1
    Right =883
    Bottom =322
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
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tblArticoliStato"
        Name =""
    End
End
