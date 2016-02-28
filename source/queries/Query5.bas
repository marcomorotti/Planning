dbMemo "SQL" ="UPDATE tblArticoli SET tblArticoli.AbcGiacenza = 'D'\015\012WHERE (((tblArticoli"
    ".AbcGiacenza) Is Null));\015\012"
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
        dbText "Name" ="tblArticoli.AbcGiacenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
End
