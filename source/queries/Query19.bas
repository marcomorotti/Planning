dbMemo "SQL" ="INSERT INTO tblArticoliStato\015\012SELECT tblArticoliObsoletiImport.OLD_COD AS "
    "COD_ART, 4 AS ID_StatoArticolo, tblArticoliObsoletiImport.NOTE & \"Inserito \" &"
    " now() AS [NOTE], tblArticoliObsoletiImport.NEW_COD AS Cod_Art_Correlato, tblArt"
    "icoliObsoletiImport.LAST_UPDATE_DATE AS Data_Modifica, 0 AS ScortaSicurezzaForza"
    "ta\015\012FROM tblArticoliObsoletiImport\015\012WHERE (((Exists (select Cod_Art\015"
    "\012from tblArticoliStato\015\012where tblArticoliStato.Cod_art = tblArticoliObs"
    "oletiImport.OLD_COD))=False));\015\012"
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
        dbText "Name" ="COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ID_StatoArticolo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliObsoletiImport.NOTE"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cod_Art_Correlato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Data_Modifica"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NOTE"
        dbLong "AggregateType" ="-1"
    End
End
