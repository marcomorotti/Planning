dbMemo "SQL" ="SELECT Sum(SConsumo_12*Cs_Csc) AS TotaleConsumoValore, Sum(SConsumo_12) AS Total"
    "eConsumoQta\015\012FROM tblArticoli\015\012WHERE (((tblArticoli.SConsumo_12)>0) "
    "AND ((tblArticoli.Cs_Csc)>0));\015\012"
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
        dbText "Name" ="TotaleConsumoValore"
        dbInteger "ColumnWidth" ="3390"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TotaleConsumoQta"
        dbInteger "ColumnWidth" ="3195"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
