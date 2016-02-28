dbMemo "SQL" ="TRANSFORM Nz(Sum([tblConsumi].[Consumo]),0) AS Consumo\015\012SELECT tblConsumi."
    "Cod_Art\015\012FROM tblConsumi\015\012WHERE (((tblConsumi.Cod_Art)=\"0000107048A"
    "\") AND (([Anno]+Format([Mese],\"00\"))>=\"201203\"))\015\012GROUP BY tblConsumi"
    ".Cod_Art\015\012PIVOT [Anno]+Format([Mese],\"00\");\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "OrderByOn" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblConsumi.[Cod_Art]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SommaDiConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Espr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201010"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201011"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201012"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="20106"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="20107"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="20108"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="20109"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="20111"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="20112"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="20113"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="20114"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="20115"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblConsumi.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Consumo$"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Consumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="01/01/1905"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="01/12/1905"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="190501"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="190512"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201006"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201007"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201008"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201009"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201101"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201102"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201103"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201104"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201105"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201204"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201205"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201206"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201207"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201208"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201209"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201210"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201211"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201212"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201301"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201302"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="[Anno]+Format([Mese],\"00\")"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="201203"
        dbLong "AggregateType" ="-1"
    End
End
