dbMemo "SQL" ="select NullId as ID_StatoArticolo, NullId as Stato, NullId as SequenzaStato from"
    " tblDummy UNION SELECT tblStatoArticolo.ID_StatoArticolo, tblStatoArticolo.Stato"
    ", tblStatoArticolo.SequenzaStato FROM tblStatoArticolo\015\012ORDER BY SequenzaS"
    "tato;\015\012"
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
        dbText "Name" ="ID_StatoArticolo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Stato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SequenzaStato"
        dbLong "AggregateType" ="-1"
    End
End
