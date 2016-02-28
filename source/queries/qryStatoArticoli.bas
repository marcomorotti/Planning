Operation =1
Option =0
Begin InputTables
    Name ="tblArticoliStato"
    Name ="tblStatoArticolo"
End
Begin OutputColumns
    Expression ="tblArticoliStato.Cod_Art"
    Expression ="tblStatoArticolo.Stato"
    Expression ="tblArticoliStato.ScortaSicurezzaForzata"
    Expression ="tblArticoliStato.Note"
End
Begin Joins
    LeftTable ="tblArticoliStato"
    RightTable ="tblStatoArticolo"
    Expression ="tblArticoliStato.ID_StatoArticolo=tblStatoArticolo.ID_StatoArticolo"
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
        dbText "Name" ="tblArticoliStato.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStatoArticolo.Stato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Note"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1443
    Bottom =852
    Left =-1
    Top =-1
    Right =1411
    Bottom =222
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
        Name ="tblStatoArticolo"
        Name =""
    End
End
