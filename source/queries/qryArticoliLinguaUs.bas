dbMemo "SQL" ="UPDATE tblArticoli INNER JOIN tblArticoliDescrEstesa ON tblArticoli.Cod_art=tblA"
    "rticoliDescrEstesa.Cod_Art SET tblArticoli.Des_art_En = tblArticoliDescrEstesa.D"
    "es_Art_Estesa\015\012WHERE tblArticoli.Cod_art=tblArticoliDescrEstesa.Cod_Art;\015"
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
