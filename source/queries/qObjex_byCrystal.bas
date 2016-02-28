Operation =1
Option =0
Where ="(((Left([Name],1))<>\"~\") AND ((Left([Name],4))<>\"MSys\"))"
Begin InputTables
    Name ="MSysObjects"
End
Begin OutputColumns
    Alias ="ObjectType"
    Expression ="GetObjectType([Type])"
    Expression ="MSysObjects.Name"
    Alias ="Type_"
    Expression ="MSysObjects.[Type]"
End
Begin OrderBy
    Expression ="GetObjectType([Type])"
    Flag =0
    Expression ="MSysObjects.Name"
    Flag =0
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
        dbText "Name" ="ObjectType"
    End
    Begin
        dbText "Name" ="Type_"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1443
    Bottom =852
    Left =-1
    Top =-1
    Right =1411
    Bottom =529
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="MSysObjects"
        Name =""
    End
End
