Operation =1
Option =0
Begin InputTables
    Name ="tblCOrdersStorico"
End
Begin OutputColumns
    Expression ="tblCOrdersStorico.NUMERO_DOC"
    Expression ="tblCOrdersStorico.RIGA_DOC"
    Expression ="tblCOrdersStorico.COD_ART"
    Expression ="tblCOrdersStorico.Descrizione"
    Expression ="tblCOrdersStorico.DATA_ORDINE"
    Expression ="tblCOrdersStorico.Giorni_Lt"
    Expression ="tblCOrdersStorico.STATO"
    Expression ="tblCOrdersStorico.liv_urgenza"
    Alias ="Evadibile"
    Expression ="StockOut([Giorni_Lt],[Liv_Urgenza],[Tipo_Doc])"
    Expression ="tblCOrdersStorico.FLAG_STOCK_OUT"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="240"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="tblCOrdersStorico.[DATA_ORDINE]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrdersStorico.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrdersStorico.Giorni_Lt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrdersStorico.STATO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrdersStorico.liv_urgenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrdersStorico.FLAG_STOCK_OUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Evadibile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrdersStorico.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrdersStorico.Descrizione"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrdersStorico.NUMERO_DOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblCOrdersStorico.RIGA_DOC"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =18
    Right =951
    Bottom =462
    Left =-1
    Top =-1
    Right =898
    Bottom =63
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =82
        Top =-13
        Right =226
        Bottom =131
        Top =0
        Name ="tblCOrdersStorico"
        Name =""
    End
End
