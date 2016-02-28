dbMemo "SQL" ="SELECT tblArticoli.Cod_art, tblArticoli.Punto_riordino, tblArticoli.ScortaSicure"
    "zza, tblArticoli.Lotto_ec_acq, tblArticoliStato.ScortaSicurezzaForzata\015\012FR"
    "OM tblArticoli INNER JOIN tblArticoliStato ON tblArticoli.Cod_art=tblArticoliSta"
    "to.Cod_Art\015\012WHERE (((tblArticoliStato.ScortaSicurezzaForzata)=0));\015\012"
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
