Operation =1
Option =0
Where ="(((A.SConsumo_12)>0) AND ((A.ROP)=0))"
Begin InputTables
    Name ="tblArticoli"
    Alias ="A"
    Name ="tblArticoliStato"
    Name ="tblStatoArticolo"
End
Begin OutputColumns
    Expression ="A.Cod_art"
    Expression ="A.Des_art"
    Expression ="A.SConsumo_12"
    Expression ="A.SSpedito_12"
    Expression ="A.Giac_Media"
    Expression ="A.Cs_Csc"
    Expression ="A.ROP"
    Expression ="A.Punto_riordino"
    Expression ="A.ROQ"
    Expression ="A.Lotto_ec_acq"
    Expression ="tblArticoliStato.ScortaSicurezzaForzata"
    Expression ="A.Classe_Evento"
    Expression ="tblStatoArticolo.Stato"
    Expression ="tblArticoliStato.Note"
End
Begin Joins
    LeftTable ="A"
    RightTable ="tblArticoliStato"
    Expression ="A.Cod_art = tblArticoliStato.Cod_Art"
    Flag =2
    LeftTable ="tblArticoliStato"
    RightTable ="tblStatoArticolo"
    Expression ="tblArticoliStato.ID_StatoArticolo = tblStatoArticolo.ID_StatoArticolo"
    Flag =2
End
Begin OrderBy
    Expression ="A.SConsumo"
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
        dbText "Name" ="A.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.Punto_riordino"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.Lotto_ec_acq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.SConsumo_12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.SSpedito_12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.Cs_Csc"
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
    Left =29
    Top =49
    Right =1273
    Bottom =638
    Left =-1
    Top =-1
    Right =1212
    Bottom =201
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="A"
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
        Name ="tblStatoArticolo"
        Name =""
    End
End
