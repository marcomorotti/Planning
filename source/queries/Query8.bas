dbMemo "SQL" ="SELECT count(tblArticoli.Cod_art)\015\012FROM tblArticoli\015\012WHERE (((tblArt"
    "icoli.Classe_Evento)='Very-Fast') And ((tblArticoli.AbcConsumoValoreLs)='A1') An"
    "d lotto_ec_acq<>0);\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
End
