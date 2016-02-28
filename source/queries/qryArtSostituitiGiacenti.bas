Operation =1
Option =0
Where ="(((tblArticoliStato.Cod_Art_Correlato)<>\"\") AND ((tblArticoli.Giac_Media)>0))"
Begin InputTables
    Name ="tblArticoliStato"
    Name ="tblArticoli"
    Name ="tblStatoArticolo"
End
Begin OutputColumns
    Expression ="tblArticoliStato.Cod_Art"
    Expression ="tblArticoliStato.DES_ART"
    Expression ="tblArticoliStato.Cod_Art_Correlato"
    Expression ="tblArticoli.Giac_Media"
    Expression ="tblArticoliStato.Note"
    Expression ="tblStatoArticolo.Stato"
    Expression ="tblArticoliStato.ScortaSicurezzaForzata"
    Expression ="tblArticoli.Cs_Csc"
End
Begin Joins
    LeftTable ="tblArticoliStato"
    RightTable ="tblArticoli"
    Expression ="tblArticoliStato.Cod_Art = tblArticoli.Cod_art"
    Flag =1
    LeftTable ="tblArticoliStato"
    RightTable ="tblStatoArticolo"
    Expression ="tblArticoliStato.ID_StatoArticolo = tblStatoArticolo.ID_StatoArticolo"
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
        dbText "Name" ="tblArticoli.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.DES_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Cod_Art_Correlato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Note"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStatoArticolo.Stato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =79
    Top =56
    Right =1214
    Bottom =671
    Left =-1
    Top =-1
    Right =1097
    Bottom =276
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =300
        Top =15
        Right =480
        Bottom =195
        Top =0
        Name ="tblArticoliStato"
        Name =""
    End
    Begin
        Left =547
        Top =16
        Right =727
        Bottom =196
        Top =0
        Name ="tblArticoli"
        Name =""
    End
    Begin
        Left =799
        Top =37
        Right =979
        Bottom =217
        Top =0
        Name ="tblStatoArticolo"
        Name =""
    End
End
