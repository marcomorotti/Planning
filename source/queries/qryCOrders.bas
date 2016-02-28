dbMemo "SQL" ="SELECT tblCOrders.liv_urgenza, tblCOrders.UTENTE, tblCOrders.NUMERO_DOC, tblCOrd"
    "ers.COD_CLI, tblCOrders.DS_RAG_SOC, tblCOrders.COD_ART, tblCOrders.Descrizione, "
    "tblCOrders.DATA_ORDINE, tblCOrders.Data_Prev_Cons, tblCOrders.qta_ord_umv, tblCO"
    "rders.qta_cons_umv, tblOnHandsSp.DISP, tblOnHandsStefani.DISP AS DispAh, qryImpe"
    "gnato.Impegnato\015\012FROM ((tblCOrders LEFT JOIN tblOnHandsSp ON tblCOrders.CO"
    "D_ART=tblOnHandsSp.COD_ART) LEFT JOIN tblOnHandsStefani ON tblCOrders.COD_ART=tb"
    "lOnHandsStefani.COD_ART) LEFT JOIN qryImpegnato ON tblCOrders.COD_ART=qryImpegna"
    "to.COD_ART\015\012WHERE (((tblCOrders.UTENTE) Like [forms]![frmOcsaMst]![txtUten"
    "te] & \"*\") AND ((tblCOrders.NUMERO_DOC) Like [forms]![frmOcsaMst]![txtNumero_D"
    "oc] & \"*\") AND ((tblCOrders.DS_RAG_SOC) Like [forms]![frmOcsaMst]![txtRag_Soc]"
    " & \"*\")) and not exists (Select *                           from tblCOrdersSpe"
    "diz                           where tblCOrders.NUMERO_DOC = tblCOrdersSpediz.NUM"
    "ERO_DOC_ORD                          AND tblCOrders.RIGA_DOC = tblCOrdersSpediz."
    "RIGA_DOC_ORD)\015\012ORDER BY tblCOrders.liv_urgenza, tblCOrders.UTENTE, tblCOrd"
    "ers.NUMERO_DOC;\015\012"
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
        dbText "Name" ="tblCOrders.COD_CLI"
        dbInteger "ColumnWidth" ="1065"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.Descrizione"
        dbInteger "ColumnWidth" ="5475"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.NUMERO_DOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.DS_RAG_SOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.UTENTE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.liv_urgenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.Data_Prev_Cons"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.qta_ord_umv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.qta_cons_umv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblOnHandsSp.DISP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryImpegnato.Impegnato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DispAh"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
End
