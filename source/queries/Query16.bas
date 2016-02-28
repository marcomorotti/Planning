dbMemo "SQL" ="UPDATE tblArticoli SET tblArticoli.Des_art_En = (select tblArticoliDescrEstesa.D"
    "es_Art_Estesa\015\012\011\011\011\011\011\011\011 from tblArticoliDescrEstesa\015"
    "\012\011\011\011\011\011\011\011 INNER JOIN tblArticoli \015\012\011\011\011\011"
    "\011\011\011\011ON tblArticoli.Cod_art=tblArticoliDescrEstesa.Cod_Art);\015\012"
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
