Operation =1
Option =0
Where ="(((tblArticoli.Cod_art) Is Null))"
Begin InputTables
    Name ="tblArticoliStato"
    Name ="tblArticoli"
End
Begin OutputColumns
    Expression ="tblArticoliStato.Cod_Art"
    Expression ="tblArticoliStato.ID_ArticoliStato"
End
Begin Joins
    LeftTable ="tblArticoliStato"
    RightTable ="tblArticoli"
    Expression ="tblArticoliStato.Cod_Art = tblArticoli.Cod_art"
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
        dbText "Name" ="[tblArticoliStato].[Cod_Art]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[tblArticoliStato].[ID_ArticoliStato]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ID_ArticoliStato"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-9
    Top =-36
    Right =1589
    Bottom =834
    Left =-1
    Top =-1
    Right =1405
    Bottom =199
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblArticoliStato"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tblArticoli"
        Name =""
    End
End
