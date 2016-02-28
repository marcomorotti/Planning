dbMemo "SQL" ="SELECT A.Cod_Art, B.Des_art, B.Cs_Csc, B.Giac_Media\015\012FROM tblConsumi AS A,"
    " tblArticoli AS B\015\012WHERE (((Exists (SELECT tblConsumi.Cod_Art FROM tblCons"
    "umi   WHERE DateSerial([Anno],[Mese],1)>=DateSerial([AnnoF],[MeseF],1) and A.Cod"
    "_art = tblConsumi.Cod_art))=False)\015\012AND ((Exists (Select * from tblArticol"
    "i where A.Cod_Art = tblArticoli.Cod_Art and tblArticoli.Giac_Media > 0))<>False)"
    ")\015\012AND B.Cod_art = A.Cod_Art\015\012ORDER BY A.Cod_Art;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="A.Cod_Art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="B.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="B.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="B.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
End
