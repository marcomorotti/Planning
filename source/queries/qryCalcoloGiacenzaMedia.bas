dbMemo "SQL" ="SELECT tblGiacenze.COD_ART, Sum(tblGiacenze.Giacenza) AS SGiacenza, Count(*) AS "
    "Num_Mesi_Giac, ((SGiacenza*30)/365) AS GiacenzaMediaMese\015\012FROM tblGiacenze"
    "\015\012WHERE DateSerial([Anno],[Mese],1)>=DateSerial([AnnoI],[MeseI],1) And Dat"
    "eSerial([Anno],[mese],1)<=DateSerial([AnnoF],[MeseF],1)\015\012GROUP BY tblGiace"
    "nze.COD_ART;\015\012"
dbMemo "Connect" =""
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
        dbText "Name" ="SGiacenza"
        dbInteger "ColumnWidth" ="1350"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Num_Mesi_Giac"
        dbInteger "ColumnWidth" ="2130"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="GiacenzaMediaMese"
        dbInteger "ColumnWidth" ="2340"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblGiacenze.COD_ART"
        dbLong "AggregateType" ="-1"
    End
End
