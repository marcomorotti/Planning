Operation =6
Option =0
Begin InputTables
    Name ="Query22"
End
Begin OutputColumns
    Expression ="Query22.[Mese]"
    GroupLevel =2
    Expression ="Query22.[Anno]"
    GroupLevel =1
    Alias ="SommaDiCosto"
    Expression ="Sum(Query22.[Costo])"
    Alias ="Totale di Costo"
    Expression ="Sum(Query22.[Costo])"
    GroupLevel =2
End
Begin Groups
    Expression ="Query22.[Mese]"
    GroupLevel =2
    Expression ="Query22.[Anno]"
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
        dbText "Name" ="Totale di Costo"
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
        dbText "Name" ="SommaDiCosto"
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
        Name ="Query22"
        Name =""
    End
End
