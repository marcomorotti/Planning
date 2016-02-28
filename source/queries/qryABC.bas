dbMemo "SQL" ="SELECT tblArticoli.Cod_art, tblArticoli.Des_art, tblArticoli.Giac_Media*tblArtic"
    "oli.Cs_Csc AS GiacVal, tblArticoli.Cs_Csc*tblArticoli.SConsumo_12 AS ConsVal, tb"
    "lArticoli.AbcGiacenza, tblArticoli.AbcConsumo, tblArticoli.AbcGiacenza & tblArti"
    "coli.AbcConsumo AS AbcGiacCons\015\012FROM tblArticoli\015\012ORDER BY tblArtico"
    "li.Cs_Csc*tblArticoli.SConsumo_12 DESC;\015\012"
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
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GiacVal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ConsVal"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcGiacenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AbcGiacCons"
        dbLong "AggregateType" ="-1"
    End
End
