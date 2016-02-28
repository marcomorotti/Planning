dbMemo "SQL" ="SELECT Sum((tblArticoli.Lotto_ec_acq/2)*(tblArticoli.Cs_Csc)) AS Espr1\015\012FR"
    "OM tblArticoli\015\012WHERE (((tblArticoli.Classe_Evento)='Very-Fast') And ((tbl"
    "Articoli.AbcConsumoValoreLs)='C2'));\015\012"
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
        dbText "Name" ="Espr1"
        dbLong "AggregateType" ="-1"
    End
End
