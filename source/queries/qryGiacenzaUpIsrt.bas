Operation =3
Name ="tblImportGiacenzaUpdate"
Option =0
Begin InputTables
End
Begin OutputColumns
    Alias ="Espr1"
    Name ="CD_ART"
    Expression ="iCD_ART"
    Alias ="Espr2"
    Name ="qt_giac"
    Expression ="iqt_giac"
    Alias ="Espr3"
    Name ="UpdateDate"
    Expression ="iUpdateDate"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Espr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Espr2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Espr3"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-220
    Top =180
    Right =845
    Bottom =764
    Left =-1
    Top =-1
    Right =1033
    Bottom =-1
    Left =0
    Top =0
    ColumnsShown =651
End
