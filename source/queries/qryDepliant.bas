Operation =1
Option =0
Where ="(((tblCOrders.UTENTE)='SPMANAGER'))"
Begin InputTables
    Name ="tblCOrders"
    Name ="tblPOrders"
End
Begin OutputColumns
    Expression ="tblCOrders.NUMERO_DOC"
    Expression ="tblCOrders.COD_CLI"
    Expression ="tblCOrders.DS_RAG_SOC"
    Expression ="tblCOrders.COD_ART"
    Expression ="tblCOrders.Descrizione"
    Expression ="tblCOrders.DATA_ORDINE"
    Expression ="tblCOrders.qta_ord_umv"
    Expression ="tblPOrders.QTA_ORDINE"
    Expression ="tblPOrders.COD_FORN"
    Expression ="tblPOrders.RAG_SOC_FORN"
    Expression ="tblPOrders.DATA_RIC"
End
Begin Joins
    LeftTable ="tblCOrders"
    RightTable ="tblPOrders"
    Expression ="tblCOrders.COD_ART=tblPOrders.COD_ART"
    Flag =1
    LeftTable ="tblCOrders"
    RightTable ="tblPOrders"
    Expression ="tblCOrders.DATA_ORDINE=tblPOrders.DATA_ORDINE"
    Flag =1
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
        dbInteger "ColumnWidth" ="2220"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrders.qta_ord_umv"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblPOrders.QTA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblPOrders.COD_FORN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblPOrders.RAG_SOC_FORN"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblPOrders.DATA_RIC"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =36
    Top =51
    Right =1011
    Bottom =535
    Left =-1
    Top =-1
    Right =943
    Bottom =255
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =224
        Top =0
        Name ="tblCOrders"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =241
        Top =0
        Name ="tblPOrders"
        Name =""
    End
End
