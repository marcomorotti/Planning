dbMemo "SQL" ="SELECT tblArticoli.Cod_art, tblArticoli.ClasseCosto, tblArticoli.MesiCopertura, "
    "tblArticoliStato.Cod_Art, tblArticoliStato.Lotto_min, tblArticoli.Classe_Evento\015"
    "\012FROM tblArticoli LEFT JOIN tblArticoliStato ON tblArticoli.Cod_art = tblArti"
    "coliStato.Cod_Art\015\012WHERE (((tblArticoli.ClasseCosto) In ('F2', 'F3')) AND\015"
    "\012       ((tblArticoli.Classe_Evento) Is Not Null));\015\012"
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
        dbText "Name" ="tblArticoli.ClasseCosto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.MesiCopertura"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Lotto_min"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
End
