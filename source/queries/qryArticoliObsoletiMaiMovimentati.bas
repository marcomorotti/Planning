dbMemo "SQL" ="SELECT A.Cod_art, A.Des_art, A.Giac_Media, A.Cs_Csc\015\012FROM tblArticoli AS A"
    "\015\012WHERE (((A.Giac_Media)>0) AND ((Exists (select *\015\012from tblConsumi "
    "B where A.Cod_art = B.Cod_art))=False));\015\012"
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
        dbText "Name" ="A.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="A.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
End
