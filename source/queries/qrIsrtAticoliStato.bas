﻿dbMemo "SQL" ="INSERT INTO tblArticoli ( cod_art, des_art )\015\012SELECT cod_art, des_art\015\012"
    "FROM tblArticoliStato\015\012WHERE (((Exists (select Cod_Art\015\012            "
    "           from tblArticoli\015\012                      where tblArticoliStato."
    "Cod_art =\015\012                            tblArticoli.Cod_art)) = False));\015"
    "\012"
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
