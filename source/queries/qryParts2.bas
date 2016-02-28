Operation =1
Option =0
Begin InputTables
    Name ="tblArticoli"
    Name ="tblArticoliStato"
    Name ="qryInventoryControl"
    Name ="tblDashboardGraphics"
End
Begin OutputColumns
    Expression ="tblArticoli.ID_Articoli"
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.Des_art"
    Expression ="tblArticoli.Des_art_En"
    Expression ="tblArticoli.Cs_Csc"
    Expression ="tblArticoli.Categ_Merc"
    Expression ="tblArticoli.Stato"
    Expression ="tblArticoli.Peso_Netto"
    Expression ="tblArticoli.Peso_Lordo"
    Expression ="tblArticoli.Costo_stoc_perc"
    Expression ="tblArticoli.Giorni_cop"
    Expression ="tblArticoli.Lead_time"
    Alias ="ScortaSicurezza"
    Expression ="tblArticoli.ScortaSicurezza"
    Expression ="tblArticoli.Anno_calcolo"
    Expression ="tblArticoli.Mese_calcolo"
    Expression ="tblArticoli.Mesi_consumo"
    Expression ="tblArticoli.Scorta_min"
    Expression ="tblArticoli.Punto_riordino"
    Expression ="tblArticoli.Lotto_ec_acq"
    Expression ="tblArticoli.Copertura"
    Expression ="tblArticoli.Giac_Media"
    Expression ="tblArticoli.GiacenzaMediaMese"
    Expression ="tblArticoli.Num_Mesi_Giac"
    Expression ="tblArticoli.Ind_Rotaz"
    Expression ="tblArticoli.Ind_Durata"
    Expression ="tblArticoli.ROP"
    Expression ="tblArticoli.MAX_MINMAX_QUANTITY"
    Expression ="tblArticoli.ROQ"
    Expression ="tblArticoli.MAXIMUM_ORDER_QUANTITY"
    Expression ="tblArticoli.SConsumo"
    Expression ="tblArticoli.SSpedito"
    Expression ="tblArticoli.Num_Eventi"
    Expression ="tblArticoli.Num_Eventi_12"
    Expression ="tblArticoli.SConsumo_12"
    Expression ="tblArticoli.SSpedito_12"
    Expression ="tblArticoli.VfNe"
    Expression ="tblArticoli.FNe"
    Expression ="tblArticoli.MfNe"
    Expression ="tblArticoli.MNe"
    Expression ="tblArticoli.MsNe"
    Expression ="tblArticoli.SNe"
    Expression ="tblArticoli.VsNe"
    Expression ="tblArticoli.Classe_Evento"
    Expression ="tblArticoli.LivelloServizio"
    Expression ="tblArticoli.ClasseCosto"
    Expression ="tblArticoli.MesiCopertura"
    Expression ="tblArticoli.AvgConsumoMese"
    Expression ="tblArticoli.DevStdConsumoMese"
    Expression ="tblArticoli.AbcConsumo"
    Expression ="tblArticoli.PctConsumo"
    Expression ="tblArticoli.AbcGiacenza"
    Expression ="tblArticoli.PctGiacenza"
    Expression ="tblArticoliStato.ID_StatoArticolo"
    Expression ="tblArticoliStato.Cod_Art_Correlato"
    Expression ="tblArticoliStato.Note"
    Expression ="tblArticoliStato.ScortaSicurezzaForzata"
    Expression ="tblArticoli.AbcConsumoValoreLs"
    Expression ="tblArticoliStato.Data_Modifica"
    Expression ="tblArticoliStato.Lotto_min"
    Expression ="tblArticoliStato.Lotto_multiplo"
    Alias ="TBO_Act"
    Expression ="(tblArticoli.ROQ/tblArticoli.SConsumo_12)*52"
    Alias ="TBO_Prop"
    Expression ="(tblArticoli.Lotto_ec_acq/tblArticoli.SConsumo_12)*52"
    Expression ="qryInventoryControl.QTY_SALE"
    Expression ="qryInventoryControl.QTY_STOCK"
    Expression ="qryInventoryControl.QTY_ACQ"
    Expression ="qryInventoryControl.Disponibile"
    Expression ="qryInventoryControl.Lotto_Acquisto"
    Expression ="qryInventoryControl.Pcnt"
    Expression ="tblDashboardGraphics.GaugesHiGood"
    Expression ="tblDashboardGraphics.GaugesLowGood"
    Expression ="tblDashboardGraphics.ProgressBarHiGood"
    Expression ="tblDashboardGraphics.ProgressBarLowGood"
    Expression ="tblDashboardGraphics.ProgressColumnHiGood"
    Expression ="tblDashboardGraphics.ProgressColumnLowGood"
End
Begin Joins
    LeftTable ="tblArticoli"
    RightTable ="tblArticoliStato"
    Expression ="tblArticoli.Cod_art=tblArticoliStato.Cod_Art"
    Flag =2
    LeftTable ="tblArticoli"
    RightTable ="qryInventoryControl"
    Expression ="tblArticoli.Cod_art=qryInventoryControl.CD_ART"
    Flag =2
    LeftTable ="qryInventoryControl"
    RightTable ="tblDashboardGraphics"
    Expression ="qryInventoryControl.Pcnt=tblDashboardGraphics.ValuePcnt"
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
dbMemo "Filter" ="((qryParts2.Lotto_Acquisto>=0)) And (qryParts2.Lotto_Acquisto>=0)"
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
    Begin
        dbText "Name" ="qryInventoryControl.QTY_SALE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryInventoryControl.QTY_STOCK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryInventoryControl.QTY_ACQ"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryInventoryControl.Disponibile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryInventoryControl.Lotto_Acquisto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDashboardGraphics.GaugesHiGood"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryInventoryControl.Pcnt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDashboardGraphics.GaugesLowGood"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDashboardGraphics.ProgressBarHiGood"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDashboardGraphics.ProgressBarLowGood"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDashboardGraphics.ProgressColumnHiGood"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDashboardGraphics.ProgressColumnLowGood"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =59
    Top =94
    Right =1424
    Bottom =791
    Left =-1
    Top =-1
    Right =1327
    Bottom =218
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="tblArticoli"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="tblArticoliStato"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =652
        Bottom =175
        Top =0
        Name ="qryInventoryControl"
        Name =""
    End
    Begin
        Left =712
        Top =15
        Right =892
        Bottom =195
        Top =0
        Name ="tblDashboardGraphics"
        Name =""
    End
End
