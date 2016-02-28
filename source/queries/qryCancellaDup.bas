dbMemo "SQL" ="DELETE *\015\012FROM tblArticoliStato\015\012WHERE ID_ArticoliStato not in\015\012"
    "(select ID_ArticoliStato from tblArticoliStato1 T2\015\012 where T2.ID_ArticoliS"
    "tato=tblArticoliStato.ID_ArticoliStato);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
End
