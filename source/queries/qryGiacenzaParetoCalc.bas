Operation =1
Option =0
Begin InputTables
    Name ="tblGiacenzePareto"
End
Begin OutputColumns
    Alias ="Cod_Art"
    Expression ="tblGiacenzePareto.Cod_art"
    Alias ="TotaleGiacenza"
    Expression ="tblGiacenzePareto.SGiacenzaValore"
    Alias ="CumPct"
    Expression ="Round(DSum(\"[SGiacenzaValore]\",\"tblGiacenzePareto\",\"[SGiacenzaValore]>=\" &"
        " [TotaleGiacenza] & \"\")/DSum(\"[SGiacenzaValore]\",\"tblGiacenzePareto\"),2)"
End
Begin OrderBy
    Expression ="tblGiacenzePareto.SGiacenzaValore"
    Flag =1
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
        dbText "Name" ="CumPct"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotaleGiacenza"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =7
    Top =25
    Right =1024
    Bottom =589
    Left =-1
    Top =-1
    Right =985
    Bottom =45
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblGiacenzePareto"
        Name =""
    End
End
