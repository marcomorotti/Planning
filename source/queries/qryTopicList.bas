dbMemo "SQL" ="SELECT DISTINCTROW tblTopics.pkeyQNumber, tblTopics.strTopic AS Topic\015\012FRO"
    "M tblTopics\015\012ORDER BY tblTopics.strTopic;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Topic"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblTopics.pkeyQNumber"
        dbLong "AggregateType" ="-1"
    End
End
