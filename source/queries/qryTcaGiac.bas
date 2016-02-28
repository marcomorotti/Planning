dbMemo "SQL" ="SELECT sum((tblArticoli.Lotto_ec_acq/2+tblArticoli.ScortaSicurezza)*(tblArticoli"
    ".Cs_Csc*0.68*0.21)) AS TcaGiac\015\012FROM tblArticoli;\015\012"
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
        dbText "Name" ="TcaGiac"
        dbLong "AggregateType" ="-1"
    End
End
