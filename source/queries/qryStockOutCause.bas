Operation =1
Option =1
Where ="((([s].[riga_doc]) Not In (select c.riga_doc\015\012                            "
    " from tblStockoutCause c)) AND (([s].[Evadibile])='StockOut') AND (([s].[data_or"
    "dine])>#5/1/2011#) AND (([s].[liv_urgenza])=2))"
Begin InputTables
    Name ="qryStockOut"
    Alias ="s"
End
Begin OutputColumns
    Alias ="Espr1"
    Expression ="s.liv_urgenza"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="s.tblCOrdersStorico.NUMERO_DOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.tblCOrdersStorico.RIGA_DOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.tblCOrdersStorico.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.tblCOrdersStorico.Descrizione"
        dbLong "AggregateType" ="-1"
        dbInteger "ColumnWidth" ="4785"
        dbBoolean "ColumnHidden" ="0"
    End
    Begin
        dbText "Name" ="s.tblCOrdersStorico.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.tblCOrdersStorico.Giorni_Lt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.tblCOrdersStorico.STATO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.Evadibile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.tblCOrdersStorico.FLAG_STOCK_OUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="s.liv_urgenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Espr1"
    End
End
Begin
    State =0
    Left =8
    Top =22
    Right =969
    Bottom =506
    Left =-1
    Top =-1
    Right =929
    Bottom =207
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="s"
        Name =""
    End
End
