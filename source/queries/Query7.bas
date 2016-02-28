dbMemo "SQL" ="SELECT Count(tblArticoli.Cod_art) AS Qta\015\012FROM tblArticoli LEFT JOIN tblAr"
    "ticoliStato ON tblArticoli.Cod_art=tblArticoliStato.Cod_Art\015\012WHERE (((tblA"
    "rticoli.Classe_Evento)='Very-Fast') And ((tblArticoli.AbcConsumoValoreLs)='A1') "
    "And ((tblArticoliStato.ScortaSicurezzaForzata) Is Null Or (tblArticoliStato.Scor"
    "taSicurezzaForzata)=0) And Lotto_ec_acq<>0);\015\012"
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
        dbText "Name" ="Qta"
        dbLong "AggregateType" ="-1"
    End
End
