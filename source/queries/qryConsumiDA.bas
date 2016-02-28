dbMemo "SQL" ="SELECT tblConsumi.Cod_Art, tblConsumi.Anno, tblConsumi.Mese, tblConsumi.Consumo "
    "AS ConsSpDa, tblConsumi1.Consumo AS ConsSp\015\012FROM tblConsumi INNER JOIN tbl"
    "Consumi1 ON (tblConsumi.Cod_Art=tblConsumi1.Cod_Art) AND (tblConsumi.Anno=tblCon"
    "sumi1.Anno) AND (tblConsumi.Mese=tblConsumi1.Mese)\015\012WHERE (((tblConsumi.co"
    "nsumo)<>[tblConsumi1].[Consumo]));\015\012"
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
        dbText "Name" ="tblConsumi.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblConsumi.Anno"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblConsumi.Mese"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ConsSpDa"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ConsSp"
        dbLong "AggregateType" ="-1"
    End
End
