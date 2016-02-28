Operation =1
Option =0
Where ="(((Exists (SELECT *\015\012FROM tblArticoli A\015\012WHERE S.COD_ART = A.COD_ART"
    "\015\012AND A.ROP > 0))<>False))"
Begin InputTables
    Name ="tblCOrdersStorico"
    Alias ="S"
End
Begin OutputColumns
    Expression ="S.NUMERO_DOC"
    Expression ="S.RIGA_DOC"
    Expression ="S.COD_ART"
    Expression ="S.Descrizione"
    Expression ="S.DATA_ORDINE"
    Expression ="S.Giorni_Lt"
    Expression ="S.STATO"
    Expression ="S.liv_urgenza"
    Alias ="Evadibile"
    Expression ="StockOut([Giorni_Lt],[Liv_Urgenza],[Tipo_Doc])"
    Expression ="S.FLAG_STOCK_OUT"
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
        dbText "Name" ="Evadibile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.NUMERO_DOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.Descrizione"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.Giorni_Lt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.STATO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.liv_urgenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.FLAG_STOCK_OUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="S.RIGA_DOC"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =19
    Top =108
    Right =949
    Bottom =552
    Left =-1
    Top =-1
    Right =898
    Bottom =63
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="S"
        Name =""
    End
End
