Operation =1
Option =0
Begin InputTables
    Name ="tblArticoli"
    Name ="tblObsoleti"
    Name ="tblArticoliStato"
End
Begin OutputColumns
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.Des_art"
    Expression ="tblObsoleti.TIPO_ARTICOLO"
    Expression ="tblObsoleti.STATO_ARTICOLO"
    Expression ="tblObsoleti.CLASSE_MERCEOLOGICA"
    Expression ="tblObsoleti.AZIENDA_COSTRUTTRICE"
    Expression ="tblObsoleti.TIPO_APP_M_P"
    Expression ="tblArticoli.Cs_Csc"
    Expression ="tblArticoli.Giac_Media"
    Expression ="tblArticoli.ROP"
    Expression ="tblArticoli.ROQ"
    Expression ="tblArticoli.Punto_riordino"
    Expression ="tblArticoli.ScortaSicurezza"
    Expression ="tblArticoliStato.ScortaSicurezzaForzata"
    Expression ="tblArticoli.Lotto_ec_acq"
    Expression ="tblArticoli.Copertura"
    Expression ="tblArticoli.SConsumo"
    Expression ="tblArticoli.SConsumo_12"
End
Begin Joins
    LeftTable ="tblObsoleti"
    RightTable ="tblArticoli"
    Expression ="tblObsoleti.ARTICOLO = tblArticoli.Cod_art"
    Flag =3
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
        dbText "Name" ="tblArticoli.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Giac_Media"
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
        dbText "Name" ="tblArticoli.SConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SConsumo_12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblObsoleti.TIPO_ARTICOLO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblObsoleti.STATO_ARTICOLO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblObsoleti.CLASSE_MERCEOLOGICA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblObsoleti.AZIENDA_COSTRUTTRICE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblObsoleti.TIPO_APP_M_P"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Punto_riordino"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ScortaSicurezza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lotto_ec_acq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Copertura"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-8
    Top =-30
    Right =1460
    Bottom =860
    Left =-1
    Top =-1
    Right =780
    Bottom =364
    Left =96
    Top =0
    ColumnsShown =539
    Begin
        Left =1
        Top =17
        Right =277
        Bottom =273
        Top =0
        Name ="tblArticoli"
        Name =""
    End
    Begin
        Left =368
        Top =41
        Right =512
        Bottom =185
        Top =0
        Name ="tblObsoleti"
        Name =""
    End
    Begin
        Left =566
        Top =52
        Right =780
        Bottom =196
        Top =0
        Name ="tblArticoliStato"
        Name =""
    End
End
