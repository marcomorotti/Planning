dbMemo "SQL" ="SELECT Sum(tblArticoli.Cs_Csc*tblArticoliStato.ScortaSicurezzaForzata) AS SSFVal"
    "o\015\012FROM tblArticoli INNER JOIN tblArticoliStato ON tblArticoli.Cod_art=tbl"
    "ArticoliStato.Cod_Art\015\012WHERE tblArticoliStato.ScortaSicurezzaForzata>0 And"
    " tblArticoli.Classe_Evento='Very-Fast' And AbcConsumoValoreLs='A1';\015\012"
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
        dbText "Name" ="SSFValo"
        dbLong "AggregateType" ="-1"
    End
End
