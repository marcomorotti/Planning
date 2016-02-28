Operation =6
Option =0
Begin InputTables
    Name ="tblGiacenze"
End
Begin OutputColumns
    Expression ="tblGiacenze.[Mese]"
    GroupLevel =2
    Expression ="tblGiacenze.[Anno]"
    GroupLevel =1
    Alias ="SommaDiGiacenza"
    Expression ="Sum(tblGiacenze.[Giacenza])"
    Alias ="Totale di Giacenza"
    Expression ="Sum(tblGiacenze.[Giacenza])"
    GroupLevel =2
End
Begin Groups
    Expression ="tblGiacenze.[Mese]"
    GroupLevel =2
    Expression ="tblGiacenze.[Anno]"
    GroupLevel =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
Begin
    Begin
        dbText "Name" ="[Mese]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Totale di Giacenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2013"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2014"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2015"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SommaDiGiacenza"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1414
    Bottom =852
    Left =-1
    Top =-1
    Right =1382
    Bottom =271
    Left =0
    Top =0
    ColumnsShown =559
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblGiacenze"
        Name =""
    End
End
