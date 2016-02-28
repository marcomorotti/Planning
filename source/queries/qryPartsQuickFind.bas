Operation =1
Option =0
Where ="(((tblArticoli.Cod_art) Like forms!frmPartsQuickFind!txtBuildName & \"*\" Or (tb"
    "lArticoli.Cod_art) Is Null) And ((tblArticoli.Des_art) Like forms!frmPartsQuickF"
    "ind!txt1BuildName & \"*\" Or (tblArticoli.Des_art) Is Null) And ((tblArticoli.RO"
    "P) Like forms!frmPartsQuickFind!txt2BuildName & \"*\" Or (tblArticoli.ROP) Is Nu"
    "ll))"
Begin InputTables
    Name ="tblArticoli"
End
Begin OutputColumns
    Expression ="tblArticoli.ID_Articoli"
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.Des_art"
    Expression ="tblArticoli.Lead_time"
    Expression ="tblArticoli.Cs_Csc"
    Expression ="tblArticoli.Giac_Media"
    Expression ="tblArticoli.AvgConsumoMese"
    Expression ="tblArticoli.ROP"
    Expression ="tblArticoli.Punto_riordino"
    Expression ="tblArticoli.ROQ"
    Expression ="tblArticoli.Lotto_ec_acq"
    Expression ="tblArticoli.AbcGiacenza"
    Expression ="tblArticoli.AbcConsumo"
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
        dbText "Name" ="tblArticoli.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lead_time"
        dbLong "AggregateType" ="-1"
        dbByte "DecimalPlaces" ="0"
    End
    Begin
        dbText "Name" ="tblArticoli.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AvgConsumoMese"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Punto_riordino"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lotto_ec_acq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ID_Articoli"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcGiacenza"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =61
    Top =286
    Right =1119
    Bottom =770
    Left =-1
    Top =-1
    Right =1026
    Bottom =-1
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
End
