Operation =1
Option =0
Where ="(((tblArticoli.InsManualmente)='S'))"
Begin InputTables
    Name ="tblArticoli"
    Name ="tblArticoliStato"
End
Begin OutputColumns
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.InsManualmente"
    Expression ="tblArticoliStato.ID_StatoArticolo"
End
Begin Joins
    LeftTable ="tblArticoli"
    RightTable ="tblArticoliStato"
    Expression ="tblArticoli.Cod_art = tblArticoliStato.Cod_Art"
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
        dbText "Name" ="tblArticoli.InsManualmente"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ID_StatoArticolo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =76
    Top =171
    Right =1359
    Bottom =822
    Left =-1
    Top =-1
    Right =1251
    Bottom =317
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
        Right =459
        Bottom =198
        Top =0
        Name ="tblArticoliStato"
        Name =""
    End
End
