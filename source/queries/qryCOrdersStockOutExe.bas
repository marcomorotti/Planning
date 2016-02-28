Operation =1
Option =0
Where ="(((NZ([tblOnHandsSp].[Disp],0)-([Qta_ord_umv]-[Qta_cons_umv]))<0) AND ((Exists ("
    "Select * from tblCOrdersSpediz  where tblCOrders.NUMERO_DOC = tblCOrdersSpediz.N"
    "UMERO_DOC_ORD AND tblCOrders.RIGA_DOC = tblCOrdersSpediz.RIGA_DOC_ORD))=False)) "
    "OR (((Exists (Select * from tblCOrdersSpediz  where tblCOrders.NUMERO_DOC = tblC"
    "OrdersSpediz.NUMERO_DOC_ORD AND tblCOrders.RIGA_DOC = tblCOrdersSpediz.RIGA_DOC_"
    "ORD))=False) AND ((NZ([tblOnHandsSp].[Disp],0)-[Impegnato])<0))"
Begin InputTables
    Name ="tblCOrders"
    Name ="tblOnHandsSp"
    Name ="tblOnHandsStefani"
    Name ="qryImpegnato"
    Name ="tblPOrders"
End
Begin OutputColumns
    Expression ="tblCOrders.liv_urgenza"
    Expression ="tblCOrders.UTENTE"
    Expression ="tblCOrders.NUMERO_DOC"
    Alias ="DataOrdine"
    Expression ="tblCOrders.DATA_ORDINE"
    Expression ="tblCOrders.COD_CLI"
    Expression ="tblCOrders.DS_RAG_SOC"
    Expression ="tblCOrders.COD_ART"
    Expression ="tblCOrders.Descrizione"
    Expression ="tblCOrders.Data_Ord_cli"
    Expression ="tblCOrders.Data_Prev_Cons"
    Expression ="tblCOrders.qta_ord_umv"
    Expression ="tblCOrders.qta_cons_umv"
    Alias ="DispSp"
    Expression ="tblOnHandsSp.DISP"
    Alias ="DispAh"
    Expression ="tblOnHandsStefani.DISP"
    Expression ="qryImpegnato.Impegnato"
    Alias ="Ord_Acq"
    Expression ="tblPOrders.NUMERO_DOC"
    Expression ="tblPOrders.COD_FORN"
    Expression ="tblPOrders.RAG_SOC_FORN"
    Expression ="tblPOrders.DATA_ORDINE"
    Expression ="tblPOrders.QTA_ORDINE"
    Expression ="tblPOrders.QTA_RESIDUA"
    Expression ="tblPOrders.DATA_RIC"
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
    LeftTable ="tblCOrders"
    RightTable ="tblPOrders"
    Expression ="tblCOrders.COD_ART=tblPOrders.COD_ART"
    Flag =2
End
Begin OrderBy
    Expression ="tblCOrders.liv_urgenza"
    Flag =0
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
        dbText "Name" ="qryImpegnato.Impegnato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DispAh"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Ord_Acq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblPOrders.COD_FORN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblPOrders.RAG_SOC_FORN"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblPOrders.QTA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblPOrders.QTA_RESIDUA"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="2025"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="tblPOrders.DATA_RIC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DispSp"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblPOrders.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DataOrdine"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =11
    Top =9
    Right =989
    Bottom =517
    Left =-1
    Top =-1
    Right =946
    Bottom =226
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
    Begin
        Left =48
        Top =156
        Right =192
        Bottom =300
        Top =0
        Name ="tblPOrders"
        Name =""
    End
End
