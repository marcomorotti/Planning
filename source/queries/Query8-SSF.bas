Operation =1
Option =0
Where ="(((tblArticoliStato.ScortaSicurezzaForzata)>0) AND ((tblArticoli.Classe_Evento)="
    "'Very-Slow') AND ((tblArticoli.AbcConsumoValoreLs)='A1'))"
Begin InputTables
    Name ="tblArticoli"
    Name ="tblArticoliStato"
End
Begin OutputColumns
    Alias ="SSFValo"
    Expression ="Sum(tblArticoli.Cs_Csc*tblArticoliStato.ScortaSicurezzaForzata)"
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
        dbText "Name" ="SSFValo"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =17
    Top =43
    Right =1004
    Bottom =791
    Left =-1
    Top =-1
    Right =949
    Bottom =274
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
End
