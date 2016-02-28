dbMemo "SQL" ="SELECT tblArticoliManuali.COD_ART, tblArticoliManuali.DES_Art, tblArticoliManual"
    "i.InsManualmente\015\012FROM tblArticoliManuali\015\012WHERE (((Exists (SELECT C"
    "od_Art\015\012                      FROM tblArticoli\015\012                    "
    " WHERE tblArticoli.Cod_art = tblArticoliManuali.COD_ART))=False));\015\012"
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
        dbText "Name" ="tblArticoliManuali.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliManuali.InsManualmente"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="3105"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblArticoliManuali.DES_Art"
        dbLong "AggregateType" ="-1"
    End
End
