Operation =1
Option =0
Begin InputTables
    Name ="tblCOrders"
End
Begin OutputColumns
    Expression ="tblCOrders.UTENTE"
End
Begin OrderBy
    Expression ="tblCOrders.UTENTE"
    Flag =0
End
Begin Groups
    Expression ="tblCOrders.UTENTE"
    GroupLevel =0
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
        dbText "Name" ="tblCOrders.UTENTE"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =999
    Bottom =604
    Left =-1
    Top =-1
    Right =967
    Bottom =228
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblCOrders"
        Name =""
    End
End
