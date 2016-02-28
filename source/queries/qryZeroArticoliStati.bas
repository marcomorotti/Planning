dbMemo "SQL" ="DELETE *\015\012FROM tblArticoliStato\015\012WHERE tblArticoliStato.ID_StatoArti"
    "colo=5 And tblArticoliStato.Cod_Art_Correlato Is Null And tblArticoliStato.Scort"
    "aSicurezzaForzata Is Null And tblArticoliStato.Lotto_min Is Null And tblArticoli"
    "Stato.Lotto_multiplo Is Null;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblArticoliStato.Note"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Cod_Art_Correlato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Data_Modifica"
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
        dbText "Name" ="tblArticoliStato.ID_ArticoliStato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ID_StatoArticolo"
        dbInteger "ColumnWidth" ="2190"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Des_art"
        dbLong "AggregateType" ="-1"
    End
End
