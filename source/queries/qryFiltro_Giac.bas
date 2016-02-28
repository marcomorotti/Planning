Operation =1
Option =0
Where ="((([Anno]+[Mese])>=[AnnoG]+[MeseG]))"
Begin InputTables
    Name ="tblRifCalcolo"
    Name ="tblGiacenze"
End
Begin OutputColumns
    Expression ="tblGiacenze.Cod_Art"
    Expression ="tblGiacenze.Anno"
    Expression ="tblGiacenze.Mese"
    Expression ="tblGiacenze.Giacenza"
    Alias ="Expr1"
    Expression ="[Anno]+[Mese]"
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
        dbText "Name" ="Expr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGiacenze.Anno"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGiacenze.Mese"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGiacenze.Giacenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGiacenze.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =134
    Top =150
    Right =921
    Bottom =679
    Left =-1
    Top =-1
    Right =755
    Bottom =173
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblRifCalcolo"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tblGiacenze"
        Name =""
    End
End
