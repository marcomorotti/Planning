dbMemo "SQL" ="SELECT tblConsumi.Cod_Art, [Anno]+[mese] AS Per, tblConsumi.Consumo, tblArticoli"
    ".Rop, tblArticoli.Giac_Media\015\012FROM tblConsumi INNER JOIN tblArticoli ON tb"
    "lConsumi.Cod_Art=tblArticoli.Cod_art\015\012GROUP BY tblConsumi.Cod_Art, [Anno]+"
    "[mese], tblConsumi.Consumo, tblArticoli.Rop, tblArticoli.Giac_Media;\015\012"
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
        dbText "Name" ="tblConsumi.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Per"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblConsumi.Consumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Rop"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
End
