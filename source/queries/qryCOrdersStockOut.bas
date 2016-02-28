Operation =1
Option =0
Where ="(((tblCOrders.UTENTE) Like [forms]![frmOcsaMstStockOut]![txtUtente] & \"*\") AND"
    " ((tblCOrders.NUMERO_DOC) Like [forms]![frmOcsaMstStockOut]![txtNumero_Doc] & \""
    "*\") AND ((tblCOrders.DS_RAG_SOC) Like [forms]![frmOcsaMstStockOut]![txtRag_Soc]"
    " & \"*\") AND ((NZ([tblOnHandsSp].[Disp],0)-([Qta_ord_umv]-[Qta_cons_umv]))<0) A"
    "ND ((Exists (Select *                           from tblCOrdersSpediz           "
    "                where tblCOrders.NUMERO_DOC = tblCOrdersSpediz.NUMERO_DOC_ORD   "
    "                       AND tblCOrders.RIGA_DOC = tblCOrdersSpediz.RIGA_DOC_ORD))"
    "=False)) OR (((tblCOrders.UTENTE) Like [forms]![frmOcsaMstStockOut]![txtUtente] "
    "& \"*\") AND ((tblCOrders.NUMERO_DOC) Like [forms]![frmOcsaMstStockOut]![txtNume"
    "ro_Doc] & \"*\") AND ((tblCOrders.DS_RAG_SOC) Like [forms]![frmOcsaMstStockOut]!"
    "[txtRag_Soc] & \"*\") AND ((Exists (Select *                           from tblC"
    "OrdersSpediz                           where tblCOrders.NUMERO_DOC = tblCOrdersS"
    "pediz.NUMERO_DOC_ORD                          AND tblCOrders.RIGA_DOC = tblCOrde"
    "rsSpediz.RIGA_DOC_ORD))=False) AND ((NZ([tblOnHandsSp].[Disp],0)-[Impegnato])<0)"
    ")"
Begin InputTables
    Name ="tblCOrders"
    Name ="tblOnHandsSp"
    Name ="tblOnHandsStefani"
    Name ="qryImpegnato"
End
Begin OutputColumns
    Expression ="tblCOrders.liv_urgenza"
    Expression ="tblCOrders.UTENTE"
    Expression ="tblCOrders.NUMERO_DOC"
    Expression ="tblCOrders.RIGA_DOC"
    Expression ="tblCOrders.COD_CLI"
    Expression ="tblCOrders.DS_RAG_SOC"
    Expression ="tblCOrders.COD_ART"
    Expression ="tblCOrders.Descrizione"
    Expression ="tblCOrders.DATA_ORDINE"
    Expression ="tblCOrders.Data_Prev_Cons"
    Expression ="tblCOrders.qta_ord_umv"
    Expression ="tblCOrders.qta_cons_umv"
    Alias ="DISP"
    Expression ="nz(tblOnHandsSp.DISP,0)"
    Alias ="DispAh"
    Expression ="NZ(tblOnHandsStefani.DISP,0)"
    Expression ="qryImpegnato.Impegnato"
End
Begin Joins
    LeftTable ="tblCOrders"
    RightTable ="tblOnHandsSp"
    Expression ="tblCOrders.COD_ART=tblOnHandsSp.COD_ART"
    Flag =2
    LeftTable ="tblCOrders"
    RightTable ="tblOnHandsStefani"
    Expression ="tblCOrders.COD_ART=tblOnHandsStefani.COD_ART"
    Flag =2
    LeftTable ="tblCOrders"
    RightTable ="qryImpegnato"
    Expression ="tblCOrders.COD_ART=qryImpegnato.COD_ART"
    Flag =2
End
Begin OrderBy
    Expression ="tblCOrders.liv_urgenza"
    Flag =1
    Expression ="tblCOrders.UTENTE"
    Flag =0
    Expression ="tblCOrders.NUMERO_DOC"
    Flag =0
End
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
    Begin
        dbText "Name" ="DISP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.RIGA_DOC"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =88
    Top =8
    Right =1153
    Bottom =592
    Left =-1
    Top =-1
    Right =1033
    Bottom =170
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblCOrders"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tblOnHandsSp"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tblOnHandsStefani"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="qryImpegnato"
        Name =""
    End
End
