Operation =1
Option =0
Where ="(((tblArticoli.AbcConsumoValoreLs)<>\"\") AND ((tblArticoli.Classe_Evento)<>\"\""
    "))"
Begin InputTables
    Name ="tblArticoli"
End
Begin OutputColumns
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.AbcConsumoValoreLs"
    Expression ="tblArticoli.Classe_Evento"
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
        dbText "Name" ="tblArticoli.AbcConsumoValoreLs"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2610"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =63
    Top =160
    Right =1008
    Bottom =812
    Left =-1
    Top =-1
    Right =913
    Bottom =279
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =315
        Bottom =156
        Top =0
        Name ="tblArticoli"
        Name =""
    End
End
