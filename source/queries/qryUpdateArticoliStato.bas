﻿dbMemo "SQL" ="UPDATE tblArticoliStato INNER JOIN tblArticoliStatoDescr ON tblArticoliStato.Cod"
    "_art=tblArticoliStatoDescr.Cod_Art SET tblArticoliStato.Des_art = tblArticoliSta"
    "toDescr.Descr_Art\015\012WHERE tblArticoliStato.Cod_art=tblArticoliStatoDescr.Co"
    "d_Art;\015\012"
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
