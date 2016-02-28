dbMemo "SQL" ="TRANSFORM Nz(Sum([Cs_Csc]*[SConsumo_12]),0) AS inventory\015\012SELECT Count(tbl"
    "Articoli.COD_ART) AS [N_di Cod_art], tblArticoli.Classe_Evento, Sum([Cs_Csc]*[SC"
    "onsumo_12]) AS [Valore_Consumo$]\015\012FROM tblArticoli\015\012GROUP BY tblArti"
    "coli.Classe_Evento\015\012PIVOT tblArticoli.MAXIMUM_ORDER_QUANTITY;\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "OrderByOn" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
Begin
    Begin
        dbText "Name" ="tblArticoli.[Classe_Evento]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ConteggioDiCod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Totale di Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.MAXIMUM_ORDER_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="0"
        dbBoolean "ColumnHidden" ="-1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PIVOT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ConteggioDiClasse_Evento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbInteger "ColumnOrder" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ConteggioDiMAXIMUM_ORDER_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="inventory1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Valore_Consumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="N_di Cod_art"
        dbInteger "ColumnOrder" ="2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="inventory"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Valore_Consumo$"
        dbInteger "ColumnWidth" ="1950"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="<>"
        dbLong "AggregateType" ="-1"
    End
End
