Operation =1
Option =0
Where ="(((tblArticoli.Categ_Merc) Like 'A4*' Or (tblArticoli.Categ_Merc) Like 'B6*' Or "
    "(tblArticoli.Categ_Merc) Like 'N8*' Or (tblArticoli.Categ_Merc) Like 'D4*' Or (t"
    "blArticoli.Categ_Merc)='S10101') AND ((tblArticoliStato.ID_StatoArticolo)=9))"
Begin InputTables
    Name ="tblArticoli"
    Name ="tblArticoliStato"
End
Begin OutputColumns
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.Categ_Merc"
    Expression ="tblArticoliStato.ID_StatoArticolo"
End
Begin Joins
    LeftTable ="tblArticoli"
    RightTable ="tblArticoliStato"
    Expression ="tblArticoli.Cod_art = tblArticoliStato.Cod_Art"
    Flag =1
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
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Categ_Merc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ID_StatoArticolo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =190
    Top =94
    Right =1410
    Bottom =728
    Left =-1
    Top =-1
    Right =1182
    Bottom =276
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =195
        Top =0
        Name ="tblArticoli"
        Name =""
    End
    Begin
        Left =300
        Top =15
        Right =480
        Bottom =195
        Top =0
        Name ="tblArticoliStato"
        Name =""
    End
End
