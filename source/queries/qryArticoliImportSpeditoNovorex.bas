dbMemo "SQL" ="SELECT tblImportSpeditoNovorex.COD_ART, tblImportSpeditoNovorex.DESCRIZIONE, tbl"
    "ImportSpeditoNovorex.UM, tblImportSpeditoNovorex.Qta_out, tblImportSpeditoNovore"
    "x.CS_CSC\015\012FROM tblImportSpeditoNovorex\015\012WHERE (((Exists (select Cd_A"
    "rt\015\012from tblGiacenzaNovorex\015\012where tblGiacenzaNovorex.Cd_art = tblIm"
    "portSpeditoNovorex.COD_ART))=False));\015\012"
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
        dbText "Name" ="tblImportSpeditoNovorex.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpeditoNovorex.DESCRIZIONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpeditoNovorex.UM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpeditoNovorex.Qta_out"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportSpeditoNovorex.CS_CSC"
        dbLong "AggregateType" ="-1"
    End
End
