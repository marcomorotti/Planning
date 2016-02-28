Operation =1
Option =0
Where ="(((tblArticoli.Cod_art) In (SELECT [Cod_art] FROM [tblArticoli] As Tmp GROUP BY "
    "[Cod_art] HAVING Count(*)>1 )))"
Begin InputTables
    Name ="tblArticoli"
End
Begin OutputColumns
    Expression ="tblArticoli.Cod_art"
    Expression ="tblArticoli.ID_Articoli"
End
Begin OrderBy
    Expression ="tblArticoli.Cod_art"
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
        dbText "Name" ="tblArticoli.Cod_art"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="tblArticoli.ID_Articoli"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =2
    Left =-9
    Top =-36
    Right =1401
    Bottom =834
    Left =-1
    Top =-1
    Right =923
    Bottom =145
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
