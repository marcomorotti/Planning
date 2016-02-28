Operation =1
Option =0
Having ="(((tblCOrdersStorico.liv_urgenza)=2))"
Begin InputTables
    Name ="tblCOrdersStorico"
End
Begin OutputColumns
    Expression ="tblCOrdersStorico.DATA_ORDINE"
    Alias ="TotaleTopP"
    Expression ="Count(tblCOrdersStorico.COD_ART)"
End
Begin Groups
    Expression ="tblCOrdersStorico.DATA_ORDINE"
    GroupLevel =0
    Expression ="tblCOrdersStorico.liv_urgenza"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCOrdersStorico.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotaleTopP"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =993
    Bottom =604
    Left =-1
    Top =-1
    Right =961
    Bottom =191
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblCOrdersStorico"
        Name =""
    End
End
