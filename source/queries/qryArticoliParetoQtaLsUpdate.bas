﻿dbMemo "SQL" ="UPDATE tblArticoli SET tblArticoli.AbcConsumoQtaLs = [iAbcConsumoQtaLs], tblArti"
    "coli.PctConsumoQtaLs = [iPctConsumoQtaLs]\015\012WHERE (((tblArticoli.Cod_Art)=["
    "iCod_Art])) And tblArticoli.Cs_Csc>0 And tblArticoli.SConsumo_12>0;\015\012"
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
