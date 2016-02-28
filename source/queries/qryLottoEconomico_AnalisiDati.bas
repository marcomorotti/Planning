Operation =1
Option =0
Begin InputTables
    Name ="tblArticoli"
    Name ="tblArticoliStato"
End
Begin OutputColumns
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.Des_art"
    Expression ="tblArticoli.Lead_time"
    Expression ="tblArticoli.Giac_Media"
    Expression ="tblArticoli.Cs_Csc"
    Expression ="tblArticoli.Categ_Merc"
    Expression ="tblArticoli.LivelloServizio"
    Expression ="tblArticoli.Punto_riordino"
    Expression ="tblArticoli.ScortaSicurezza"
    Expression ="tblArticoli.Lotto_ec_acq"
    Expression ="tblArticoli.ROP"
    Expression ="tblArticoli.ROQ"
    Expression ="tblArticoli.ClasseCosto"
    Expression ="tblArticoli.Classe_Evento"
    Expression ="tblArticoliStato.ID_StatoArticolo"
    Expression ="tblArticoliStato.ScortaSicurezzaForzata"
    Expression ="tblArticoliStato.Lotto_min"
    Expression ="tblArticoliStato.Lotto_multiplo"
    Expression ="tblArticoli.MesiCopertura"
    Expression ="tblArticoli.SConsumo"
    Expression ="tblArticoli.SSpedito"
End
Begin Joins
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
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Categ_Merc"
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
        dbText "Name" ="tblArticoli.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ClasseCosto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ID_StatoArticolo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Lotto_min"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Lotto_multiplo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.MesiCopertura"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.LivelloServizio"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SSpedito"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lead_time"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-9
    Top =-36
    Right =1607
    Bottom =853
    Left =-1
    Top =-1
    Right =1578
    Bottom =466
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =336
        Bottom =464
        Top =0
        Name ="tblArticoli"
        Name =""
    End
    Begin
        Left =564
        Top =31
        Right =805
        Bottom =343
        Top =0
        Name ="tblArticoliStato"
        Name =""
    End
End
