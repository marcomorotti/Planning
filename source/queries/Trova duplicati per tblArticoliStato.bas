dbMemo "SQL" ="SELECT tblArticoliStato.Cod_art, tblArticoliStato.ID_ArticoliStato\015\012FROM t"
    "blArticoliStato\015\012WHERE (((tblArticoliStato.Cod_art) In (SELECT [Cod_art] F"
    "ROM [tblArticoliStato] As Tmp GROUP BY [Cod_art] HAVING Count(*)> 1 )))\015\012O"
    "RDER BY tblArticoliStato.Cod_art;\015\012"
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
dbMemo "Filter" ="([Trova duplicati per tblArticoliStato].[Cod_art]=\"0001307032A\")"
Begin
    Begin
        dbText "Name" ="tblArticoliStato.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ID_ArticoliStato"
        dbLong "AggregateType" ="-1"
    End
End
