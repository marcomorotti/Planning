﻿dbMemo "SQL" ="UPDATE tblArticoliStato INNER JOIN tblArticoli ON tblArticoli.Cod_art=tblArticol"
    "iStato.Cod_Art SET tblArticoliStato.ID_StatoArticolo = 9\015\012WHERE (((tblArti"
    "coli.Categ_Merc) Like 'A4*' Or (tblArticoli.Categ_Merc) Like 'B6*' Or (tblArtico"
    "li.Categ_Merc) Like 'N8*' Or (tblArticoli.Categ_Merc) Like 'D4*' Or (tblArticoli"
    ".Categ_Merc)='S10101')) And tblArticoliStato.ID_StatoArticolo Is Null;\015\012"
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
