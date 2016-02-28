dbMemo "SQL" ="UPDATE tblArticoli SET tblArticoli.Giac_Media = [iGiac_Media]\015\012WHERE (((tb"
    "lArticoli.Cod_Art)=[iCod_Art]));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblArticoli.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Exists (select * from tblImportGiacenza where tblArticoli.Cod_Art = tblImportGia"
            "cenza.CD_Art)"
        dbLong "AggregateType" ="-1"
    End
End
