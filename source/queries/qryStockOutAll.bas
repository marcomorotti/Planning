Operation =1
Option =0
Where ="((([qryStockOut].[Evadibile])='StockOut'))"
Begin InputTables
    Name ="qryStockOut"
End
Begin OutputColumns
    Alias ="Espr1"
    Expression ="qryStockOut.DATA_ORDINE"
    Alias ="TotStockOutAll"
    Expression ="Count(*)"
End
Begin Groups
    Expression ="qryStockOut.DATA_ORDINE"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="240"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="qryStockOut.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotStockOutAll"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Espr1"
    End
End
Begin
    State =0
    Left =21
    Top =18
    Right =951
    Bottom =462
    Left =-1
    Top =-1
    Right =898
    Bottom =130
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =49
        Top =-8
        Right =193
        Bottom =136
        Top =0
        Name ="qryStockOut"
        Name =""
    End
End
