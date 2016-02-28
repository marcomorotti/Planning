dbMemo "SQL" ="SELECT tblArticoli.ID_Articoli, tblArticoli.Cod_art, tblArticoli.Des_art, tblArt"
    "icoli.Des_art_En, tblArticoli.Cs_Csc, tblArticoli.Categ_Merc, tblArticoli.Stato,"
    " tblArticoli.Peso_Netto, tblArticoli.Peso_Lordo, tblArticoli.Costo_stoc_perc, tb"
    "lArticoli.Giorni_cop, tblArticoli.Lead_time, tblArticoli.ScortaSicurezza AS Scor"
    "taSicurezza, tblArticoli.Anno_calcolo, tblArticoli.Mese_calcolo, tblArticoli.Mes"
    "i_consumo, tblArticoli.Scorta_min, tblArticoli.Punto_riordino, tblArticoli.Lotto"
    "_ec_acq, tblArticoli.Copertura, tblArticoli.Giac_Media, tblArticoli.GiacenzaMedi"
    "aMese, tblArticoli.Num_Mesi_Giac, tblArticoli.Ind_Rotaz, tblArticoli.Ind_Durata,"
    " tblArticoli.ROP, tblArticoli.MAX_MINMAX_QUANTITY, tblArticoli.ROQ, tblArticoli."
    "MAXIMUM_ORDER_QUANTITY, tblArticoli.SConsumo, tblArticoli.SSpedito, tblArticoli."
    "Num_Eventi, tblArticoli.Num_Eventi_12, tblArticoli.SConsumo_12, tblArticoli.SSpe"
    "dito_12, tblArticoli.VfNe, tblArticoli.FNe, tblArticoli.MfNe, tblArticoli.MNe, t"
    "blArticoli.MsNe, tblArticoli.SNe, tblArticoli.VsNe, tblArticoli.Classe_Evento, t"
    "blArticoli.LivelloServizio, tblArticoli.ClasseCosto, tblArticoli.MesiCopertura, "
    "tblArticoli.AvgConsumoMese, tblArticoli.DevStdConsumoMese, tblArticoli.AbcConsum"
    "o, tblArticoli.PctConsumo, tblArticoli.AbcGiacenza, tblArticoli.PctGiacenza, tbl"
    "ArticoliStato.ID_StatoArticolo, tblArticoliStato.Cod_Art_Correlato, tblArticoliS"
    "tato.Note, tblArticoliStato.ScortaSicurezzaForzata, tblArticoli.AbcConsumoValore"
    "Ls, tblArticoliStato.Data_Modifica, tblArticoliStato.Lotto_min, tblArticoliStato"
    ".Lotto_multiplo, (tblArticoli.ROQ/tblArticoli.SConsumo_12)*52 AS TBO_Act, (tblAr"
    "ticoli.Lotto_ec_acq/tblArticoli.SConsumo_12)*52 AS TBO_Prop\015\012FROM tblArtic"
    "oli LEFT JOIN tblArticoliStato ON tblArticoli.Cod_art=tblArticoliStato.Cod_Art;\015"
    "\012"
dbMemo "Connect" =""
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
        dbText "Name" ="tblArticoli.PctGiacenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ID_StatoArticolo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Note"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ID_Articoli"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Copertura"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Cod_Art_Correlato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ScortaSicurezza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Costo_stoc_perc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Giorni_cop"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lead_time"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Anno_calcolo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Mese_calcolo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Mesi_consumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Scorta_min"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Punto_riordino"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Lotto_ec_acq"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Giac_Media"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Ind_Rotaz"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Ind_Durata"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROP"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.MAX_MINMAX_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ROQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.MAXIMUM_ORDER_QUANTITY"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SSpedito"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Num_Eventi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Num_Eventi_12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SConsumo_12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SSpedito_12"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.VfNe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.FNe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.MfNe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.MNe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.MsNe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.SNe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.VsNe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Classe_Evento"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.LivelloServizio"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ClasseCosto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.MesiCopertura"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AvgConsumoMese"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.DevStdConsumoMese"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.PctConsumo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcGiacenza"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.ScortaSicurezzaForzata"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.AbcConsumoValoreLs"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Des_art_En"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Num_Mesi_Giac"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Data_Modifica"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Lotto_min"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoliStato.Lotto_multiplo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.GiacenzaMediaMese"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Categ_Merc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Stato"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Peso_Netto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Peso_Lordo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TBO_Act"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TBO_Prop"
        dbLong "AggregateType" ="-1"
    End
End
