Operation =1
Option =0
Begin InputTables
    Name ="tblDashBoard"
    Name ="tblDashboardGraphics"
End
Begin OutputColumns
    Expression ="tblDashBoard.Indice"
    Expression ="tblDashBoard.Actual"
    Expression ="tblDashBoard.Target"
    Expression ="tblDashBoard.PcntIndice"
    Expression ="tblDashboardGraphics.GaugesHiGood"
End
Begin Joins
    LeftTable ="tblDashBoard"
    RightTable ="tblDashboardGraphics"
    Expression ="tblDashBoard.PcntIndice=tblDashboardGraphics.ValuePcnt"
    Flag =1
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
        dbText "Name" ="tblDashboardGraphics.GaugesHiGood"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDashBoard.Indice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDashBoard.Actual"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDashBoard.Target"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblDashBoard.PcntIndice"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1065
    Bottom =555
    Left =-1
    Top =-1
    Right =1033
    Bottom =123
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =89
        Top =13
        Right =233
        Bottom =157
        Top =0
        Name ="tblDashBoard"
        Name =""
    End
    Begin
        Left =292
        Top =12
        Right =436
        Bottom =156
        Top =0
        Name ="tblDashboardGraphics"
        Name =""
    End
End
