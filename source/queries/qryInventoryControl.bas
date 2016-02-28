Operation =1
Option =0
Begin InputTables
    Name ="tblArticoli"
    Name ="tblInventoryControl"
End
Begin OutputColumns
    Expression ="tblInventoryControl.CD_ART"
    Expression ="tblArticoli.Des_art"
    Expression ="tblInventoryControl.QTY_SALE"
    Expression ="tblInventoryControl.QTY_STOCK"
    Expression ="tblInventoryControl.QTY_ACQ"
    Expression ="tblArticoli.Punto_riordino"
    Expression ="tblArticoli.ScortaSicurezza"
    Expression ="tblArticoli.Lotto_ec_acq"
    Alias ="Disponibile"
    Expression ="(tblInventoryControl.QTY_ACQ+tblInventoryControl.QTY_STOCK-tblInventoryControl.Q"
        "TY_SALE-tblArticoli.ScortaSicurezza)"
    Alias ="Lotto_Acquisto"
    Expression ="IIf((tblInventoryControl.QTY_ACQ+tblInventoryControl.QTY_STOCK-tblInventoryContr"
        "ol.QTY_SALE)<tblArticoli.Punto_riordino,tblArticoli.Lotto_ec_acq-(tblInventoryCo"
        "ntrol.QTY_ACQ+tblInventoryControl.QTY_STOCK-tblInventoryControl.QTY_SALE-tblArti"
        "coli.ScortaSicurezza),0)"
    Alias ="Pcnt"
    Expression ="IIf(tblArticoli.Lotto_ec_acq<=0,0.01,IIf(IsNull(tblArticoli.Lotto_ec_acq),0.01,I"
        "If(tblInventoryControl.QTY_STOCK=0,0.001,IIf((tblInventoryControl.QTY_STOCK/tblA"
        "rticoli.Lotto_ec_acq)>=1,1,Round(tblInventoryControl.QTY_STOCK/tblArticoli.Lotto"
        "_ec_acq,2)))))"
    Expression ="tblArticoli.Cs_Csc"
End
Begin Joins
    LeftTable ="tblArticoli"
    RightTable ="tblInventoryControl"
    Expression ="tblArticoli.Cod_art=tblInventoryControl.CD_ART"
    Flag =3
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
        dbText "Name" ="tblInventoryControl.CD_ART"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Des_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInventoryControl.QTY_SALE"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInventoryControl.QTY_STOCK"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblInventoryControl.QTY_ACQ"
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
        dbText "Name" ="Disponibile"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Lotto_Acquisto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Pcnt"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-259
    Top =121
    Right =1478
    Bottom =833
    Left =-1
    Top =-1
    Right =1699
    Bottom =177
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =60
        Top =15
        Right =240
        Bottom =195
        Top =0
        Name ="tblArticoli"
        Name =""
    End
    Begin
        Left =300
        Top =15
        Right =480
        Bottom =195
        Top =0
        Name ="tblInventoryControl"
        Name =""
    End
End
