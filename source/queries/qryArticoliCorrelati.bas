Operation =1
Option =0
Where ="(((tblArticoli.ROP)>0))"
Begin InputTables
    Name ="tblArticoli"
    Name ="tblArticoliStato"
    Name ="tblConsumi"
End
Begin OutputColumns
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.ROP"
    Expression ="tblArticoli.ROQ"
    Expression ="tblArticoliStato.Cod_Art_Correlato"
End
Begin Joins
    LeftTable ="tblArticoli"
    RightTable ="tblArticoliStato"
    Expression ="tblArticoli.Cod_art = tblArticoliStato.Cod_Art"
    Flag =1
    LeftTable ="tblArticoliStato"
    RightTable ="tblConsumi"
    Expression ="tblArticoliStato.Cod_Art = tblConsumi.Cod_Art"
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
        dbText "Name" ="tblArticoli.Cod_art"
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
    Begin
        dbText "Name" ="tblArticoliStato.Cod_Art_Correlato"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =14
    Top =16
    Right =1031
    Bottom =580
    Left =-1
    Top =-1
    Right =985
    Bottom =189
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
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tblConsumi"
        Name =""
    End
End
