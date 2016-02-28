dbMemo "SQL" ="DELETE tblStockOutCause.evadibile, *\015\012FROM tblStockOutCause;\015\012"
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
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="tblStockOutCause.[evadibile]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblStockOutCause.evadibile"
        dbLong "AggregateType" ="-1"
    End
End
