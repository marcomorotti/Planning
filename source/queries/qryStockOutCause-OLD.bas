dbMemo "SQL" ="SELECT tblCOrders.liv_urgenza, tblCOrders.UTENTE, tblCOrders.NUMERO_DOC, tblCOrd"
    "ers.COD_CLI, tblCOrders.DS_RAG_SOC, tblCOrders.COD_ART, tblCOrders.Descrizione, "
    "tblCOrders.Data_Ord_cli, tblCOrders.Data_Prev_Cons, tblCOrders.qta_ord_umv, tblC"
    "Orders.qta_cons_umv, tblOnHandsSp.DISP AS DispSp, qryImpegnato.Impegnato, IIf((N"
    "Z(tblOnHandsSp.Disp,0)-([Qta_ord_umv]-[Qta_cons_umv]))<0,1,0)\015\012FROM (tblCO"
    "rders LEFT JOIN tblOnHandsSp ON tblCOrders.COD_ART=tblOnHandsSp.COD_ART) LEFT JO"
    "IN qryImpegnato ON tblCOrders.COD_ART=qryImpegnato.COD_ART\015\012WHERE NOT EXIS"
    "TS (Select * from tblCOrdersSpediz where tblCOrders.NUMERO_DOC = tblCOrdersSpedi"
    "z.NUMERO_DOC_ORD AND tblCOrders.RIGA_DOC = tblCOrdersSpediz.RIGA_DOC_ORD) AND (("
    "NZ([tblOnHandsSp].[Disp],0)-([Qta_ord_umv]-[Qta_cons_umv]))<0 Or (NZ([tblOnHands"
    "Sp].[Disp],0)-[Impegnato])<0) AND tblCOrders.liv_urgenza = 2\015\012ORDER BY tbl"
    "COrders.liv_urgenza, tblCOrders.UTENTE, tblCOrders.NUMERO_DOC;\015\012"
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
        dbText "Name" ="tblCOrders.liv_urgenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.UTENTE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.NUMERO_DOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.COD_CLI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.DS_RAG_SOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.Descrizione"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.Data_Ord_cli"
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
        dbText "Name" ="DispSp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryImpegnato.Impegnato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1013"
        dbLong "AggregateType" ="-1"
    End
End
