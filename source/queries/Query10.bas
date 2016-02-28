dbMemo "SQL" ="INSERT INTO tblArticoliStato\015\012SELECT tblArticoli.COD_ART AS COD_ART, 9 AS "
    "ID_StatoArticolo, \"Inserito \" & now() AS [NOTE]\015\012FROM tblArticoli\015\012"
    "WHERE (((Exists\015\012           (select Cod_Art\015\012                from tb"
    "lArticoliStato\015\012               where tblArticoli.Cod_art = tblArticoliStat"
    "o.Cod_art)) =\015\012          False))\015\012      AND (((tblArticoli.Categ_Mer"
    "c) Like 'A4*' Or\015\012          (tblArticoli.Categ_Merc) Like 'B6*' Or\015\012"
    "          (tblArticoli.Categ_Merc) Like 'N8*' Or\015\012          (tblArticoli.C"
    "ateg_Merc) Like 'D4*' Or\015\012          (tblArticoli.Categ_Merc) = 'S10101'));"
    "\015\012"
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
    End
    Begin
        dbText "Name" ="ID_StatoArticolo"
    End
    Begin
        dbText "Name" ="NOTE"
    End
End
