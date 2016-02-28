dbMemo "SQL" ="SELECT sum(Importo_Netto)\015\012FROM tblCOrdersStorico\015\012WHERE tblCOrdersS"
    "torico.DATA_ORDINE>=#10/1/2011# And tblCOrdersStorico.DATA_ORDINE<=#10/31/2011# "
    "And tblCOrdersStorico.COD_CLI Not In ('FD','FE','FF','FGB','FOLD','FPOL','TE') A"
    "nd tblCOrdersStorico.soc_ord=' SP';\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
End
