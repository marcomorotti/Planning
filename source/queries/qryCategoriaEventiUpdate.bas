﻿dbMemo "SQL" ="UPDATE tblArticoli SET tblArticoli.Classe_Evento = [iClasse_Evento], tblArticoli"
    ".ClasseCosto = [iClasseCosto], tblArticoli.MesiCopertura = [iMesiCopertura]\015\012"
    "WHERE (((tblArticoli.Cod_Art)=[iCod_Art]));\015\012"
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
