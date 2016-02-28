Operation =1
Option =0
Begin InputTables
    Name ="tblDatiGenerali"
    Name ="qryFiltro_Giac"
End
Begin OutputColumns
    Expression ="qryFiltro_Giac.Cod_Art"
    Alias ="StockMedio"
    Expression ="Sum([Giacenza]/[mesi_giacenze])"
End
Begin Groups
    Expression ="qryFiltro_Giac.Cod_Art"
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
        dbText "Name" ="StockMedio"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2535"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="qryFiltro_Giac.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =102
    Top =83
    Right =889
    Bottom =688
    Left =-1
    Top =-1
    Right =755
    Bottom =197
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblDatiGenerali"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="qryFiltro_Giac"
        Name =""
    End
End
