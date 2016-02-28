Operation =1
Option =0
Begin InputTables
    Name ="tblVenduto"
    Name ="tblArticoli"
    Name ="tblArticoliStato"
End
Begin OutputColumns
    Expression ="tblVenduto.*"
    Expression ="tblArticoli.Giac_Media"
    Expression ="tblArticoli.ROP"
    Expression ="tblArticoli.ROQ"
    Expression ="tblArticoli.Punto_riordino"
    Expression ="tblArticoli.ScortaSicurezza"
    Expression ="tblArticoli.Lotto_ec_acq"
    Expression ="tblArticoliStato.ScortaSicurezzaForzata"
    Expression ="tblArticoli.Classe_Evento"
    Expression ="tblArticoli.Num_Eventi_12"
    Expression ="tblArticoli.LivelloServizio"
    Expression ="tblArticoli.MesiCopertura"
    Expression ="tblArticoli.AbcConsumo"
End
Begin Joins
    LeftTable ="tblVenduto"
    RightTable ="tblArticoli"
    Expression ="tblVenduto.COD_ART = tblArticoli.Cod_art"
    Flag =2
    LeftTable ="tblArticoli"
    RightTable ="tblArticoliStato"
    Expression ="tblArticoli.Cod_art = tblArticoliStato.Cod_Art"
    Flag =2
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
        dbText "Name" ="tblVenduto.COD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.DESCRIZIONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Punto_riordino"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ScortaSicurezza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lotto_ec_acq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Num_Eventi_12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.LivelloServizio"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.MesiCopertura"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.SOCIETA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.UTENTE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.STATO_CALCOLATO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.ORDINE_WB"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.NUMERO_DOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.PRIORITÀ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.DATA_ORDINE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.DATA_DDT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.LIVELLO_SERVIZIO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.RITARDO_SPEDIZIONE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.NUM_ORD_CLI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.CLIENTEINTERNO_FILIALI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.COD_CLI"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.DS_RAG_SOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.DESCRIZIONE_CAUSALE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.QTA_ORD_UMV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.QTA_CONS_UMV"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.FUNZ_DOC"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.FLAG_STOCK_OUT"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.EVASIONE_COMPLETA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.ANNO_ORD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.MESE_ORD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.GIORNO_ORD"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.IMPORTO_NETTO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.COSTO_STAND"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.DS_NAZ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.CLASSE_MERCEOLOGICA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.RESPONSABILE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.STATO_CONSEGNA"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.COD_LISTINO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.PREZZO_LISTINO"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblVenduto.RESPONSABILE_PROMO"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =38
    Top =86
    Right =1434
    Bottom =852
    Left =-1
    Top =-1
    Right =1364
    Bottom =483
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblVenduto"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =410
        Bottom =425
        Top =0
        Name ="tblArticoli"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="tblArticoliStato"
        Name =""
    End
End
