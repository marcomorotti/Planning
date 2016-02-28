dbMemo "SQL" ="SELECT tblImportContoLavoro.CD_ART, tblImportContoLavoro.DESCR_ART, tblImportCon"
    "toLavoro.TB_UBIC, NZ([tblImportContoLavoroGiacSp.QT_GIAC],0) AS GiacSp, tblImpor"
    "tContoLavoro.QT_GIAC, tblImportContoLavoro.QT_IMP, tblImportContoLavoro.UpdateDa"
    "te\015\012FROM tblImportContoLavoro LEFT JOIN tblImportContoLavoroGiacSp ON tblI"
    "mportContoLavoro.CD_ART=tblImportContoLavoroGiacSp.CD_ART\015\012WHERE tblImport"
    "ContoLavoro.QT_IMP-tblImportContoLavoro.QT_GIAC<>0;\015\012"
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
        dbText "Name" ="tblImportContoLavoro.CD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportContoLavoro.QT_GIAC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GiacSp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportContoLavoro.DESCR_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportContoLavoro.TB_UBIC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportContoLavoro.QT_IMP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblImportContoLavoro.UpdateDate"
        dbLong "AggregateType" ="-1"
    End
End
