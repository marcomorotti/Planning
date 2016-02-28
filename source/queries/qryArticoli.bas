Operation =1
Option =0
Where ="(((tblArticoli.Num_Eventi_12)>0))"
Begin InputTables
    Name ="tblArticoli"
End
Begin OutputColumns
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.Num_Eventi"
    Expression ="tblArticoli.Num_Eventi_12"
    Expression ="tblArticoli.VfNe"
    Expression ="tblArticoli.FNe"
    Expression ="tblArticoli.MfNe"
    Expression ="tblArticoli.MNe"
    Expression ="tblArticoli.MsNe"
    Expression ="tblArticoli.SNe"
    Expression ="tblArticoli.VsNe"
    Expression ="tblArticoli.Cs_Csc"
    Alias ="ClasseCosto"
    Expression ="ClasseCosto([Cs_Csc])"
    Alias ="MesiCopertura"
    Expression ="MesiCopertura([Cs_Csc])"
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
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Num_Eventi"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.VsNe"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.Num_Eventi_12"
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
        dbText "Name" ="tblArticoli.Cs_Csc"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClasseCosto"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="MesiCopertura"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =80
    Right =1060
    Bottom =604
    Left =-1
    Top =-1
    Right =1007
    Bottom =45
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
End
