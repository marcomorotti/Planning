Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5340
    DatasheetFontHeight =10
    ItemSuffix =47
    Left =9360
    Top =1875
    Right =14955
    Bottom =8400
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3955d11af5b9e340
    End
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            Width =850
            Height =850
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
        End
        Begin Section
            Height =6292
            BackColor =13229799
            Name ="Corpo"
            AlternateBackColor =13229799
            Begin
                Begin Label
                    SpecialEffect =4
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =2
                    OverlapFlags =85
                    TextAlign =2
                    Left =-60
                    Width =4695
                    Height =555
                    FontSize =20
                    FontWeight =700
                    BackColor =12311007
                    ForeColor =13209
                    Name ="lblTitolo"
                    Caption ="IMPORT - CALCOLI"
                    FontName ="Verdana"
                    LayoutCachedLeft =-60
                    LayoutCachedWidth =4635
                    LayoutCachedHeight =555
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =1605
                    Top =1485
                    Width =2244
                    Height =620
                    FontSize =10
                    FontWeight =700
                    ForeColor =1845071
                    Name ="cmdAggiornaConsumi"
                    Caption ="1 - AGGIORNA \015\012   DATI"
                    OnClick ="[Event Procedure]"
                    FontName ="Gill Sans MT"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1605
                    LayoutCachedTop =1485
                    LayoutCachedWidth =3849
                    LayoutCachedHeight =2105
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =4755
                    Top =15
                    Width =576
                    Height =576
                    TabIndex =1
                    ForeColor =0
                    Name ="cmdClose"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadad0dadadadadaadad00adadadadaddad030dadadadada ,
                        0xad0330adadadadad0033300000000adaa03330ff0dadadadd03300ff0adad4da ,
                        0xa03330ff0dad44add03330ff0ad44444a03330ff0d444444d03330ff0ad44444 ,
                        0xa0330fff0dad44add030ffff0adad4daa00fffff0dadadadd00000000adadada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Tag ="ShiftLeftNewUnit,ShiftLeftPermLocn"
                    ControlTipText ="Close the form"

                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =1605
                    Top =2225
                    Width =2244
                    Height =620
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    ForeColor =1845071
                    Name ="cmdEventi"
                    Caption ="2 - CALCOLO\015\012   CONSUMI"
                    OnClick ="[Event Procedure]"
                    FontName ="Gill Sans MT"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1605
                    LayoutCachedTop =2225
                    LayoutCachedWidth =3849
                    LayoutCachedHeight =2845
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1590
                    Top =4648
                    Width =2244
                    Height =620
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    ForeColor =1845071
                    Name ="cmdABC"
                    Caption ="5 - CALCOLO\015\012   ABC"
                    OnClick ="[Event Procedure]"
                    FontName ="Gill Sans MT"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1590
                    LayoutCachedTop =4648
                    LayoutCachedWidth =3834
                    LayoutCachedHeight =5268
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =1605
                    Top =2985
                    Width =2244
                    Height =620
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    ForeColor =1845071
                    Name ="cmdRopRoq"
                    Caption ="3 - CALCOLO\015\012   ROP ROQ"
                    OnClick ="[Event Procedure]"
                    FontName ="Gill Sans MT"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1605
                    LayoutCachedTop =2985
                    LayoutCachedWidth =3849
                    LayoutCachedHeight =3605
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =1590
                    Top =735
                    Width =2244
                    Height =620
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    ForeColor =1845071
                    Name ="cmdImportaConsumi"
                    Caption ="0 - IMPORTA \015\012   CONSUMI"
                    OnClick ="[Event Procedure]"
                    FontName ="Gill Sans MT"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1590
                    LayoutCachedTop =735
                    LayoutCachedWidth =3834
                    LayoutCachedHeight =1355
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1587
                    Top =5442
                    Width =2244
                    Height =620
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    ForeColor =1845071
                    Name ="cmdOnHand"
                    Caption ="6 - AGGIORNA\015\012 DISPONIBILE"
                    OnClick ="[Event Procedure]"
                    FontName ="Gill Sans MT"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1587
                    LayoutCachedTop =5442
                    LayoutCachedWidth =3831
                    LayoutCachedHeight =6062
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =795
                    Width =113
                    Height =114
                    BackColor =10040879
                    Name ="box01"
                    LayoutCachedLeft =227
                    LayoutCachedTop =795
                    LayoutCachedWidth =340
                    LayoutCachedHeight =909
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =225
                    Top =962
                    Width =113
                    Height =114
                    BackColor =10040879
                    Name ="box02"
                    LayoutCachedLeft =225
                    LayoutCachedTop =962
                    LayoutCachedWidth =338
                    LayoutCachedHeight =1076
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =225
                    Top =1142
                    Width =113
                    Height =114
                    BackColor =10040879
                    Name ="box03"
                    LayoutCachedLeft =225
                    LayoutCachedTop =1142
                    LayoutCachedWidth =338
                    LayoutCachedHeight =1256
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =1530
                    Width =113
                    Height =114
                    BackColor =5026082
                    Name ="box11"
                    LayoutCachedLeft =227
                    LayoutCachedTop =1530
                    LayoutCachedWidth =340
                    LayoutCachedHeight =1644
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =225
                    Top =1697
                    Width =113
                    Height =114
                    BackColor =5026082
                    Name ="box12"
                    LayoutCachedLeft =225
                    LayoutCachedTop =1697
                    LayoutCachedWidth =338
                    LayoutCachedHeight =1811
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =225
                    Top =1877
                    Width =113
                    Height =114
                    BackColor =5026082
                    Name ="box13"
                    LayoutCachedLeft =225
                    LayoutCachedTop =1877
                    LayoutCachedWidth =338
                    LayoutCachedHeight =1991
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =2325
                    Width =113
                    Height =114
                    BackColor =9974127
                    Name ="box21"
                    LayoutCachedLeft =227
                    LayoutCachedTop =2325
                    LayoutCachedWidth =340
                    LayoutCachedHeight =2439
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =225
                    Top =2492
                    Width =113
                    Height =114
                    BackColor =9974127
                    Name ="box22"
                    LayoutCachedLeft =225
                    LayoutCachedTop =2492
                    LayoutCachedWidth =338
                    LayoutCachedHeight =2606
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =225
                    Top =2672
                    Width =113
                    Height =114
                    BackColor =9974127
                    Name ="box23"
                    LayoutCachedLeft =225
                    LayoutCachedTop =2672
                    LayoutCachedWidth =338
                    LayoutCachedHeight =2786
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =3060
                    Width =113
                    Height =114
                    BackColor =2366701
                    Name ="box31"
                    LayoutCachedLeft =227
                    LayoutCachedTop =3060
                    LayoutCachedWidth =340
                    LayoutCachedHeight =3174
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =3227
                    Width =113
                    Height =114
                    BackColor =2366701
                    Name ="box32"
                    LayoutCachedLeft =227
                    LayoutCachedTop =3227
                    LayoutCachedWidth =340
                    LayoutCachedHeight =3341
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =3407
                    Width =113
                    Height =114
                    BackColor =2366701
                    Name ="box33"
                    LayoutCachedLeft =227
                    LayoutCachedTop =3407
                    LayoutCachedWidth =340
                    LayoutCachedHeight =3521
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =3
                    OverlapFlags =215
                    Left =1417
                    Top =680
                    Width =2608
                    Height =2948
                    BorderColor =2366701
                    Name ="Casella30"
                    LayoutCachedLeft =1417
                    LayoutCachedTop =680
                    LayoutCachedWidth =4025
                    LayoutCachedHeight =3628
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =3855
                    Width =113
                    Height =114
                    BackColor =10040879
                    Name ="box41"
                    LayoutCachedLeft =227
                    LayoutCachedTop =3855
                    LayoutCachedWidth =340
                    LayoutCachedHeight =3969
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =4022
                    Width =113
                    Height =114
                    BackColor =5026082
                    Name ="box42"
                    LayoutCachedLeft =227
                    LayoutCachedTop =4022
                    LayoutCachedWidth =340
                    LayoutCachedHeight =4136
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =4202
                    Width =113
                    Height =114
                    BackColor =2366701
                    Name ="box43"
                    LayoutCachedLeft =227
                    LayoutCachedTop =4202
                    LayoutCachedWidth =340
                    LayoutCachedHeight =4316
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =4755
                    Width =113
                    Height =114
                    BackColor =62207
                    Name ="box51"
                    LayoutCachedLeft =227
                    LayoutCachedTop =4755
                    LayoutCachedWidth =340
                    LayoutCachedHeight =4869
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =4922
                    Width =113
                    Height =114
                    BackColor =62207
                    Name ="box52"
                    LayoutCachedLeft =227
                    LayoutCachedTop =4922
                    LayoutCachedWidth =340
                    LayoutCachedHeight =5036
                End
                Begin Rectangle
                    SpecialEffect =0
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =5102
                    Width =113
                    Height =114
                    BackColor =62207
                    BorderColor =62207
                    Name ="box53"
                    LayoutCachedLeft =227
                    LayoutCachedTop =5102
                    LayoutCachedWidth =340
                    LayoutCachedHeight =5216
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =5550
                    Width =113
                    Height =114
                    BackColor =-2147483617
                    Name ="box61"
                    LayoutCachedLeft =227
                    LayoutCachedTop =5550
                    LayoutCachedWidth =340
                    LayoutCachedHeight =5664
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =5717
                    Width =113
                    Height =114
                    BackColor =-2147483617
                    Name ="box62"
                    LayoutCachedLeft =227
                    LayoutCachedTop =5717
                    LayoutCachedWidth =340
                    LayoutCachedHeight =5831
                End
                Begin Rectangle
                    SpecialEffect =2
                    BackStyle =1
                    OverlapFlags =85
                    Left =227
                    Top =5897
                    Width =113
                    Height =114
                    BackColor =-2147483617
                    Name ="box63"
                    LayoutCachedLeft =227
                    LayoutCachedTop =5897
                    LayoutCachedWidth =340
                    LayoutCachedHeight =6011
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1590
                    Top =3795
                    Width =2244
                    Height =620
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    ForeColor =1845071
                    Name ="cmdAggiornaAll"
                    Caption ="4 - AGGIORNA\015\012   TUTTO"
                    OnClick ="[Event Procedure]"
                    FontName ="Gill Sans MT"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1590
                    LayoutCachedTop =3795
                    LayoutCachedWidth =3834
                    LayoutCachedHeight =4415
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =73
                    Left =4755
                    Top =570
                    Width =577
                    Height =577
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    ForeColor =0
                    Name ="cmdSearch"
                    Caption ="R&icerca"
                    OnClick ="=aprimaschere(\"frmFasiElaboraDati\")"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Search ..."
                    UnicodeAccessKey =105
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000b0a0900000000000b0a09000a0705000 ,
                        0xa070501000000000a07860d0a07860ff906040ff905840ff905840ff804020f0 ,
                        0x8040201000000000b0806000b08060f0906850e0a07050ff905840ff905030ff ,
                        0x703010f000000000b09080fffff8ffffe0c8c0ffd0a090ffc08060ff804020ff ,
                        0x0000000000000000b0806010b08060fff0e8e0ffe0c8c0ffd0a890ffb07850ff ,
                        0x804820ff00000000b09080fffff8ffffe0c8c0ffd0a090ffc08060ff804020ff ,
                        0x000000000000000000000000b08060fff0e8e0fff0e0e0ffe0c0b0ffc08870ff ,
                        0x804830ff00000000b09080fffff8ffffe0c8c0ffd0a090ffc08060ff804020ff ,
                        0x000000000000000000000000b08060fff0e8e0fff0e0e0ffe0c0b0ffc08870ff ,
                        0x804830ff00000000c09880fffffffffff0e8e0ffe0c8c0ffd0a080ff804020ff ,
                        0x000000000000000000000000b08060fff0e8e0fff0e0e0ffe0c0b0ffc08870ff ,
                        0x804830ff00000000c0a090f0b08870ffa06850ff905030ff804830ff804820ff ,
                        0x803810ff803810e0b08870ffa06850ff905830ff904830ff804020ff703810ff ,
                        0x905830ff00000000c0a09080b08870ffffffffffe0d0c0ffd0a090ffa07050ff ,
                        0x804010ffb0907080b09070ffe0d8d0fff0d8d0ffd0a090ffb07850ff803820ff ,
                        0xb090708000000000c0a09000c09080fff0f0f0fffff8f0fff0d8c0ffb08060ff ,
                        0x804820ff803810d0b09070fffffffffffff8f0fff0d0c0ffb07850ff804820ff ,
                        0x000000000000000000000000b0908010c09880ffb08060ffa06850ff905030ff ,
                        0x905840ff905830e0b07860ffb08870ffa07050ff804830ff804820ffb0806010 ,
                        0x00000000000000000000000000000000c09880fffff8ffffe0c0b0ffc09070ff ,
                        0x804820ff70301000c09880fffff8ffffe0c8b0ffd0a080ff804820ff00000000 ,
                        0x00000000000000000000000000000000c0a090c0b09080ffa06850ff905030ff ,
                        0x804830f000000000c0a890ffb09080ffa06850ff905030ff804830f000000000 ,
                        0x0000000000000000000000000000000000000000905840b0fff8f0ff703010e0 ,
                        0x803810000000000080381000905840b0fff8f0ff703010e00000000000000000 ,
                        0x0000000000000000000000000000000000000000c0a090e0b08870f0905830e0 ,
                        0x000000000000000000000000c0a090c0b08870f0905830e00000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =4755
                    LayoutCachedTop =570
                    LayoutCachedWidth =5332
                    LayoutCachedHeight =1147
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub cmdABC_Click()
'DoCmd.SetWarnings False
'DoCmd.Hourglass True

'Definizioni
'*************************************************************************
    Dim db0 As Database
    Dim ClasseAPerc As Variant
    Dim ClasseBPerc As Variant
    Dim ClasseCPerc As Variant
    Dim DatiGen As New ADODB.Recordset
    Dim strSQL As String
    Dim Msg As String
    Dim Style As String
    Dim Title As String
    Dim Response As String
    Dim qr1 As QueryDef
    Dim brs As DAO.Recordset
    Dim bqry As DAO.QueryDef
    Dim conn As ADODB.Connection
    Dim intI As Variant
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    '*************************************************************************

    'memorizzo le soglie di Classi ABC da tblDatiGenerali
    DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
    ClasseAPerc = DatiGen.Fields("ClasseAPerc")
    ClasseBPerc = DatiGen.Fields("ClasseBPerc")
    ClasseCPerc = DatiGen.Fields("ClasseCPerc")
    DatiGen.Close


    ' Vuoi calcolare i dati?
    Msg = "Vuoi Calcolare le Classi ABC di Consumo e Giacenza?"    ' Define message.
    Style = vbYesNo + vbQuestion + vbDefaultButton2    ' Define buttons.
    Title = "MsgBox Import"    ' Define title.
    Response = MsgBox(Msg, Style, Title)
    If Response = vbYes Then GoTo INIZIO Else    ' User chose Yes.
    Exit Sub

INIZIO:
    ' 201602
  ' Aggiorna tblFasiElaboraDati cob data inizio
  Call Tempofase("51", True) 'Scrive inizio

    WriteToLog ("'Inizio: 4 - CALCOLO ABC'")
    WriteToLog ("'  Inizio Calcolo Abc Consumo'")


    ' ************************************************************************
    ' CALCOLO ABC CONSUMO
    ' ************************************************************************


    'Cancella i dati tblConsumiPareto
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * from  tblConsumiPareto"


    strSQL = "INSERT INTO tblConsumiPareto " & "SELECT Cod_Art, " & _
             "SConsumo_12 * Cs_Csc As SConsumoValore " & "FROM [tblArticoli] " & _
             "WHERE tblArticoli.SConsumo_12 > 0 " & "AND tblArticoli.Cs_Csc > 0 " & _
             "Order By (SConsumo_12 * Cs_Csc) desc"

    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True

    ' *** Memorizzo Classe ABC in tblArticoli
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("AGGIORNO CLASSE ABC CONSUMI", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryConsumoParetoCalc")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
        If Not brs.BOF Then    'se ci sono record nel recordset
            brs.MoveLast    ' necessario per determinare l'attuale numero di record
            intMassimo = brs.RecordCount
            brs.MoveFirst
        End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art"), brs.Fields("SConsumo"), brs.Fields("SSpedito"), brs.Fields("Num_Eventi")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryArticoliParetoConsUpdate")
        qr1!iCOD_ART = brs![Cod_art]
        Select Case brs![CumPct]
        Case Is <= ClasseAPerc / 100
            qr1!iAbcConsumo = "A"
            qr1!iPctConsumo = brs![CumPct]
        Case Is <= (ClasseAPerc + ClasseBPerc) / 100
            qr1!iAbcConsumo = "B"
            qr1!iPctConsumo = brs![CumPct]
        Case Else
            qr1!iAbcConsumo = "C"
            qr1!iPctConsumo = brs![CumPct]
        End Select
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    bqry.Close
    
    ' 201602 Chiude Fase 51
    Call FaseEseguita("51")
    
 '20160216 Puntini Avanzamento
    Me.box51.Visible = True
    
    ' 201602 Inizia Fase 02
    Call Tempofase("52", True)
    

    WriteToLog ("'  Inizio Calcolo Abc Giacenza'")

    ' Inizio:
    ' ************************************************************************
    ' CALCOLO ABC GIACENZA
    ' ************************************************************************



    'Cancella i dati tblGiacenzePareto
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * from  tblGiacenzePareto"


    strSQL = "INSERT INTO tblGiacenzePareto " & "SELECT Cod_Art, " & _
             "Giac_Media * Cs_Csc As SGiacenzaValore " & "FROM [tblArticoli] " & _
             "WHERE tblArticoli.Giac_Media > 0 " & "AND tblArticoli.Cs_Csc > 0 " & _
             "Order By Giac_Media * Cs_Csc desc;"

    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True

    ' *** Memorizzo Classe ABC in tblArticoli
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("CLASSE ABC GIACENZE", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    '
    Set bqry = db0.QueryDefs("qryGiacenzaParetoCalc")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
        If Not brs.BOF Then    'se ci sono record nel recordset
            brs.MoveLast    ' necessario per determinare l'attuale numero di record
            intMassimo = brs.RecordCount
            brs.MoveFirst
        End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art"), brs.Fields("SConsumo"), brs.Fields("SSpedito"), brs.Fields("Num_Eventi")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryArticoliParetoGiacUpdate")
        qr1!iCOD_ART = brs![Cod_art]
        Select Case brs![CumPct]
        Case Is <= ClasseAPerc / 100
            qr1!iAbcGiacenza = "A"
            qr1!iPctGiacenza = brs![CumPct]
        Case Is <= (ClasseAPerc + ClasseBPerc) / 100
            qr1!iAbcGiacenza = "B"
            qr1!iPctGiacenza = brs![CumPct]
        Case Else
            qr1!iAbcGiacenza = "C"
            qr1!iPctGiacenza = brs![CumPct]
        End Select
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    bqry.Close
    MsgBox "Finish", vbInformation, "CLASSE ABC GIACENZA UPDATE"

    ' *** AGGIORNA CLASSE D per GIACENZA***
    db0.Execute _
            "Update tblArticoli set  tblArticoli.AbcGiacenza = 'D' WHERE tblArticoli.AbcGiacenza Is Null; ", _
            dbFailOnError
    ' *** AGGIORNA CLASSE D per CONSUMO***
    db0.Execute _
            "Update tblArticoli set  tblArticoli.AbcConsumo = 'D' WHERE tblArticoli.AbcConsumo Is Null; ", _
            dbFailOnError

    ' MsgBox "Pareto calculation done!", vbOKOnly, ""

    WriteToLog ("'Fine: 4 - CALCOLO ABC'")
    
    FaseEseguita ("52")

    MsgBox "Calcolo ABC finito!"

    'DoCmd.SetWarnings True
    'DoCmd.Hourglass False
End Sub

Private Sub CmdAggiornaAll_Click()
' 20160216
    Response = MsgBox("Vuoi importare i Dati nel DATA BASE ", vbYesNo, "Continue")
    If Response = vbYes Then
        ' Imposta parametro Calcolo_All su tblParametri
        Call scrivichiave("AggiornaAll", "SI")
        Response = MsgBox("Vuoi partire da VOCE - 1 - AGGIORNA DATI ", vbYesNo, "Continua")
        If Response = vbNo Then
            ' Chiama pulsante 0
            ' 201602
            ' Aggiorna tblFasiElaboraDati cob data inizio
            Call Tempofase("41", True)    'Scrive inizio
            Call cmdImportaConsumi_Click
            '20160216 Puntini avanzamento
            FaseEseguita ("41")
            Me.box41.Visible = True

        End If
        ' Se non passa per importa Consumi aggiorna Eseguito = False
        If Response = vbNo Then
            DoCmd.SetWarnings False
            strSQL = "UPDATE [tblFasiElaboraDati] SET [tblFasiElaboraDati].Eseguita = False " & _
                     "WHERE [tblFasiElaboraDati].num_fase not in ('01', '02', '03'); "
            DoCmd.RunSQL strSQL
            DoCmd.SetWarnings True
        End If

        ' Chiama pulsante 1
        Call Tempofase("42", True)    'Scrive inizio
        Call cmdAggiornaConsumi_Click
        '20160216 Puntini avanzamento
        FaseEseguita ("42")
        Me.box42.Visible = True

        ' Chiama pulsante 2 **** Non Gestiti pnt. avanzamento
        Call cmdEventi_Click

        ' Chiama pulsante 3
        Call Tempofase("43", True)    'Scrive inizio
        Call cmdRopRoq_Click
        '20160216 Puntini avanzamento
        FaseEseguita ("43")
        Me.box43.Visible = True
        Call scrivichiave("AggiornaAll", "NO")

    Else
        '        Me.box5.Visible = False
        Exit Sub
    End If
End Sub

Private Sub cmdClose_Click()
    ' Close me
    DoCmd.Close acForm, Me.name
End Sub
Private Sub cmdEventi_Click()
    Dim ws As Workspace
    Dim Db As Database
    Dim db0 As Database    ' ******* DATA BASE CORRENTE ***********
    Dim Lconnect As String
    Dim MyQuery As String
    Dim qm As QueryDef
    Dim rst As Object
    Dim rs As Recordset
    Dim Response As Integer
    Dim VarWhereIn As Variant
    Dim conn As ADODB.Connection
    Dim intI As Double
    ' Verificare i Dim
    Dim intMonth As Integer, intYear As Integer, intStartDate As Variant, _
        intStartDateS As Variant, intEndDate As Variant, intLastDay As Integer, NumMesiC As Integer
    Dim DatiGen As New ADODB.Recordset
    Dim N_Ordini_Evasi_12_mesi, N_Ultimi_Mesi As Integer
    ' Dim CatEventi As New adodb.Recordset
    ' Dim CalMag As New adodb.Recordset
    ' Dim CalGiac As New adodb.Recordset
    Dim Art As DAO.Recordset
    Dim AnnoC As String
    Dim MeseC As String
    Dim AnnoI As String
    Dim MeseI As String
    Dim AnnoF As String
    Dim MeseF As String
    Dim Cmd As New ADODB.Command
    Dim bqry As DAO.QueryDef
    Dim brs As DAO.Recordset
    Dim qr1 As QueryDef
    Dim Classe_Evento As String
    Dim strSQL As String
    Dim ConsMeseT(1 To 12) As Variant
    Dim intCounter As Integer


    ' Dim nRecords As Integer 'Long
    On Error GoTo Err_cmdImportDataOracle_Click:

    ' 20160216 Gestione Aggiona ALL
    If leggiChiave("aggiornaAll") = "NO" Then
        Msg = "Vuoi Calcolare i Consumi ?"    ' Definisce titolo messaggio.
        Style = vbYesNo + vbQuestion + vbDefaultButton2    ' Definisce pulsante.
        Title = "MsgBox Aggiorna Data Base"    ' Definisce Titolo.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            GoTo INIZIO
        Else
            Exit Sub
        End If
    End If

INIZIO:
    ' 201602
    ' Aggiorna tblFasiElaboraDati cob data inizio
    Call Tempofase("21", True)    'Scrive inizio
    WriteToLog ("'Inizio: 2 - CALCOLA CONSUMI'")

    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    NumMesiC = DatiGen.Fields("Mesi_consumo")
    intYear = DatiGen.Fields("Anno_Calcolo")
    intMonth = DatiGen.Fields("Mese_Calcolo")
    intLastDay = Day(DateAdd("m", 1, DateSerial(intYear, intMonth, 1)) - 1)
    intEndDate = DateSerial(intYear, intMonth, intLastDay)
    intEndDate = Format(intEndDate, "dd/mmm/yyyy")
    intStartDate = DateAdd("m", -NumMesiC + 1, DateSerial(intYear, intMonth, 1))
    intStartDate = Format(intStartDate, "dd/mmm/yyyy")
    AnnoC = Year(intStartDate)
    MeseC = Month(intStartDate)
    AnnoF = Year(intEndDate)
    MeseF = Format(Month(intEndDate), "00")

    AnnoI = Year(DateAdd("m", -11, DateSerial(intYear, intMonth, 1)))
    MeseI = Format(Month(DateAdd("m", -11, DateSerial(intYear, intMonth, 1))), "00")
    ' GoTo start

    ' qui
    ' ***   CALCOLO GIACENZA MEDIA ***********************************************************************
    'Visualizza lo Status Meter
    Call acbInitMeter("GIACENZA MEDIA", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media giacenza
    Set bqry = db0.QueryDefs("qryCalcoloGiacenzaMedia")
    ' Imposta i parametri qry parametrica
    bqry.Parameters("AnnoI").Value = AnnoI
    bqry.Parameters("MeseI").Value = MeseI
    bqry.Parameters("AnnoF").Value = AnnoF
    bqry.Parameters("MeseF").Value = MeseF
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Set qr1 = db0.QueryDefs("qryGiacenzaMediaUpdate")
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art")
        'Memorizzo i dati nella tblarticoli
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iNum_Mesi_Giac = brs![Num_Mesi_Giac]
        qr1!iGiacenzaMediaMese = brs![GiacenzaMediaMese]
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    ' qui fine
    ' 201602 Chiude Fase 01
    Call FaseEseguita("21")

    '20160216 Puntini Avanzamento
    Me.box21.Visible = True

    ' 201602 Inizia Fase 02
    Call Tempofase("22", True)



    ' ***   CALCOLO CATEGORIA EVENTI ***********************************************************************
    ' ******************************************************************************************************
    ' EVENTI 12 MESI - Calcolo tblArticolo.Num_Eventi_12 (Numero eventi negli ultimi 12 mese
    ' tblArticolo.Num_Eventi viene calcolato su tutto il periodo di calcolo come dalla Tabella Dati Generali
    ' ******************************************************************************************************
    N_Ordini_Evasi_12_mesi = DLookup("N_Ordini_Evasi_12_mesi", "tblCategoriaEventi", _
                                     "Classe_Evento = 'Very-Fast'")
    N_Ultimi_Mesi = DLookup("N_Ultimi_Mesi", "tblCategoriaEventi", _
                            "Classe_Evento = 'Very-Fast'")
    ' VfNe = DLookup("N_Ordini_IN_N_Ultimi_Mesi", "tblCategoriaEventi", "Classe_Evento = 'Very-Fast'")
    'AnnoI = Year(DateAdd("m", -11, DateSerial(intYear, intMonth, 1)))
    'MeseI = Format(Month(DateAdd("m", -11, DateSerial(intYear, intMonth, 1))), "00")
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("SPEDIZIONI 12 MESI UPDATE", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloEventi")
    ' Imposta i parametri qry parametrica
    bqry.Parameters("AnnoI").Value = AnnoI
    bqry.Parameters("MeseI").Value = MeseI
    bqry.Parameters("AnnoF").Value = AnnoF
    bqry.Parameters("MeseF").Value = MeseF
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryEventiUpdate12")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iNum_Eventi_12 = brs![Num_Eventi]
        qr1!iSConsumo_12 = brs![SConsumo]
        qr1!iSSpedito_12 = brs![SSpedito]
        qr1!iAvgConsumoMese = brs![AvgConsumoMese]
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    ' ************************************************************************************************************
    ' * VERY-FAST EVENTI N MESI - Calcolo tblArticolo.VfNe nei mesi n specificati in tbl.CapegoriaEventi
    ' ************************************************************************************************************

    Dim VfNe As Integer
    VfNe = DLookup("N_Ultimi_Mesi", "tblCategoriaEventi", _
                   "Classe_Evento = 'Very-Fast'")
    AnnoI = Year(DateAdd("m", -VfNe, DateSerial(intYear, intMonth, 1)))
    MeseI = Format(Month(DateAdd("m", -VfNe, DateSerial(intYear, intMonth, 1))), _
                   "00")
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("SPEDIZIONI VERY-FAST", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloEventi")
    ' Imposta i parametri qry parametrica
    bqry.Parameters("AnnoI").Value = AnnoI
    bqry.Parameters("MeseI").Value = MeseI
    bqry.Parameters("AnnoF").Value = AnnoF
    bqry.Parameters("MeseF").Value = MeseF
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryEventiUpdateVf")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iVfNe = brs![Num_Eventi]
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close

    ' start:
    ' ************************************************************************************************************
    ' *  FAST EVENTI N MESI - Calcolo tblArticolo.FNe nei mesi n specificati in tbl.CapegoriaEventi
    ' ************************************************************************************************************
    Dim FNe As Integer
    FNe = DLookup("N_Ultimi_Mesi", "tblCategoriaEventi", "Classe_Evento = 'Fast'")
    AnnoI = Year(DateAdd("m", -FNe, DateSerial(intYear, intMonth, 1)))
    MeseI = Format(Month(DateAdd("m", -FNe, DateSerial(intYear, intMonth, 1))), _
                   "00")
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("SPEDIZIONI FAST", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloEventi")
    ' Imposta i parametri qry parametrica
    bqry.Parameters("AnnoI").Value = AnnoI
    bqry.Parameters("MeseI").Value = MeseI
    bqry.Parameters("AnnoF").Value = AnnoF
    bqry.Parameters("MeseF").Value = MeseF
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryEventiUpdateF")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iFNe = brs![Num_Eventi]
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close

    ' ************************************************************************************************************
    ' * MEDIUM-FAST EVENTI N MESI - Calcolo tblArticolo.MfNe nei mesi n specificati in tbl.CapegoriaEventi
    ' ************************************************************************************************************
    Dim MfNe As Integer
    MfNe = DLookup("N_Ultimi_Mesi", "tblCategoriaEventi", _
                   "Classe_Evento = 'Medium-Fast'")
    AnnoI = Year(DateAdd("m", -MfNe, DateSerial(intYear, intMonth, 1)))
    MeseI = Format(Month(DateAdd("m", -MfNe, DateSerial(intYear, intMonth, 1))), _
                   "00")
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("SPEDIZIONI MEDIUM-FAST", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloEventi")
    ' Imposta i parametri qry parametrica
    bqry.Parameters("AnnoI").Value = AnnoI
    bqry.Parameters("MeseI").Value = MeseI
    bqry.Parameters("AnnoF").Value = AnnoF
    bqry.Parameters("MeseF").Value = MeseF
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If

    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryEventiUpdateMf")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iMfNe = brs![Num_Eventi]
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    ' ************************************************************************************************************
    ' * MEDIUM EVENTI N MESI - Calcolo tblArticolo.MNe nei mesi n specificati in tbl.CapegoriaEventi
    ' ************************************************************************************************************
    Dim MNe As Integer
    MNe = DLookup("N_Ultimi_Mesi", "tblCategoriaEventi", "Classe_Evento = 'Medium'")
    AnnoI = Year(DateAdd("m", -MNe, DateSerial(intYear, intMonth, 1)))
    MeseI = Format(Month(DateAdd("m", -MNe, DateSerial(intYear, intMonth, 1))), _
                   "00")
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("SPEDIZIONI MEDIUM", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloEventi")
    ' Imposta i parametri qry parametrica
    bqry.Parameters("AnnoI").Value = AnnoI
    bqry.Parameters("MeseI").Value = MeseI
    bqry.Parameters("AnnoF").Value = AnnoF
    bqry.Parameters("MeseF").Value = MeseF
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryEventiUpdateM")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iMNe = brs![Num_Eventi]
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close

    ' 20160216 Puntini avanzamento
    Me.box22.Visible = True

    ' ************************************************************************************************************
    ' * MEDIUM-SLOW EVENTI N MESI - Calcolo tblArticolo.MsNe nei mesi n specificati in tbl.CapegoriaEventi
    ' ************************************************************************************************************
    Dim MsNe As Integer
    MsNe = DLookup("N_Ultimi_Mesi", "tblCategoriaEventi", _
                   "Classe_Evento = 'Medium-Slow'")
    AnnoI = Year(DateAdd("m", -MsNe, DateSerial(intYear, intMonth, 1)))
    MeseI = Format(Month(DateAdd("m", -MsNe, DateSerial(intYear, intMonth, 1))), _
                   "00")
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("SPEDIZIONI MEDIUM-SLOW", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloEventi")
    ' Imposta i parametri qry parametrica
    bqry.Parameters("AnnoI").Value = AnnoI
    bqry.Parameters("MeseI").Value = MeseI
    bqry.Parameters("AnnoF").Value = AnnoF
    bqry.Parameters("MeseF").Value = MeseF
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryEventiUpdateMs")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iMsNe = brs![Num_Eventi]
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    ' start:
    ' ************************************************************************************************************
    ' * SLOW EVENTI N MESI - Calcolo tblArticolo.sNe nei mesi n specificati in tbl.CapegoriaEventi
    ' ************************************************************************************************************
    Dim SNe As Integer
    SNe = DLookup("N_Ultimi_Mesi", "tblCategoriaEventi", "Classe_Evento = 'Slow'")
    AnnoI = Year(DateAdd("m", -SNe, DateSerial(intYear, intMonth, 1)))
    MeseI = Format(Month(DateAdd("m", -SNe, DateSerial(intYear, intMonth, 1))), _
                   "00")
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("SPEDIZIONI SLOW", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloEventi")
    ' Imposta i parametri qry parametrica
    bqry.Parameters("AnnoI").Value = AnnoI
    bqry.Parameters("MeseI").Value = MeseI
    bqry.Parameters("AnnoF").Value = AnnoF
    bqry.Parameters("MeseF").Value = MeseF
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryEventiUpdateS")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iSNe = brs![Num_Eventi]
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close

    ' ************************************************************************************************************
    ' * VERY-SLOW EVENTI N MESI - Calcolo tblArticolo.VsNe nei mesi n specificati in tbl.CapegoriaEventi
    ' ************************************************************************************************************
    Dim VsNe As Integer
    VsNe = DLookup("N_Ultimi_Mesi", "tblCategoriaEventi", _
                   "Classe_Evento = 'Very-Slow'")
    AnnoI = Year(DateAdd("m", -VsNe, DateSerial(intYear, intMonth, 1)))
    MeseI = Format(Month(DateAdd("m", -VsNe, DateSerial(intYear, intMonth, 1))), _
                   "00")
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("SPEDIZIONI VERY-SLOW", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloEventi")
    ' Imposta i parametri qry parametrica
    bqry.Parameters("AnnoI").Value = AnnoI
    bqry.Parameters("MeseI").Value = MeseI
    bqry.Parameters("AnnoF").Value = AnnoF
    bqry.Parameters("MeseF").Value = MeseF
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryEventiUpdateVS")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iVSNe = brs![Num_Eventi]
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    ' 201602 Chiude Fase 21
    Call FaseEseguita("22")

    '20160216 Puntini Avanzamento
    Me.box22.Visible = True

    ' 201602 Inizia Fase 23
    Call Tempofase("23", True)

    ' start:
    ' *** UPDATE CLASSE_EVENTI - LIVELLO_SERVIZIO - CLASSE_COSTO
    Dim VfNe12 As Integer
    Dim FNe12 As Integer
    Dim MfNe12 As Integer
    Dim MNe12 As Integer
    Dim MsNe12 As Integer
    Dim sNe12 As Integer
    Dim VsNe12 As Integer

    Dim VfNeTab As Long
    Dim FNeTab As Long
    Dim MfNeTab As Long
    Dim MNeTab As Long
    Dim MsNeTab As Long
    Dim sNeTab As Long
    Dim VsNeTab As Long



    VfNe12 = DLookup("N_Ordini_Evasi_12_mesi", "tblCategoriaEventi", _
                     "Classe_Evento = 'Very-Fast'")
    FNe12 = DLookup("N_Ordini_Evasi_12_mesi", "tblCategoriaEventi", _
                    "Classe_Evento = 'Fast'")
    MfNe12 = DLookup("N_Ordini_Evasi_12_mesi", "tblCategoriaEventi", _
                     "Classe_Evento = 'Medium-Fast'")
    MNe12 = DLookup("N_Ordini_Evasi_12_mesi", "tblCategoriaEventi", _
                    "Classe_Evento = 'Medium'")
    MsNe12 = DLookup("N_Ordini_Evasi_12_mesi", "tblCategoriaEventi", _
                     "Classe_Evento = 'Medium-Slow'")
    sNe12 = DLookup("N_Ordini_Evasi_12_mesi", "tblCategoriaEventi", _
                    "Classe_Evento = 'Slow'")
    VsNe12 = DLookup("N_Ordini_Evasi_12_mesi", "tblCategoriaEventi", _
                     "Classe_Evento = 'Very-Slow'")

    VfNeTab = DLookup("N_Ordini_IN_N_Ultimi_Mesi", "tblCategoriaEventi", _
                      "Classe_Evento = 'Very-Fast'")
    FNeTab = DLookup("N_Ordini_IN_N_Ultimi_Mesi", "tblCategoriaEventi", _
                     "Classe_Evento = 'Fast'")
    MfNeTab = DLookup("N_Ordini_IN_N_Ultimi_Mesi", "tblCategoriaEventi", _
                      "Classe_Evento = 'Medium-Fast'")
    MNeTab = DLookup("N_Ordini_IN_N_Ultimi_Mesi", "tblCategoriaEventi", _
                     "Classe_Evento = 'Medium'")
    MsNeTab = DLookup("N_Ordini_IN_N_Ultimi_Mesi", "tblCategoriaEventi", _
                      "Classe_Evento = 'Medium-Slow'")
    sNeTab = DLookup("N_Ordini_IN_N_Ultimi_Mesi", "tblCategoriaEventi", _
                     "Classe_Evento = 'Slow'")
    VsNeTab = DLookup("N_Ordini_IN_N_Ultimi_Mesi", "tblCategoriaEventi", _
                      "Classe_Evento = 'Very-Slow'")

    'VfLs = DLookup("LivelloServizio", "tblCategoriaEventi", _
     '    "Classe_Evento = 'Very-Fast'")
    'FLs = DLookup("LivelloServizio", "tblCategoriaEventi", "Classe_Evento = 'Fast'")
    'MfLs = DLookup("LivelloServizio", "tblCategoriaEventi", _
     '    "Classe_Evento = 'Medium-Fast'")
    'MLs = DLookup("LivelloServizio", "tblCategoriaEventi", _
     '    "Classe_Evento = 'Medium'")
    'MsLs = DLookup("LivelloServizio", "tblCategoriaEventi", _
     '    "Classe_Evento = 'Medium-Slow'")
    'sLs = DLookup("LivelloServizio", "tblCategoriaEventi", "Classe_Evento = 'Slow'")
    'VsLs = DLookup("LivelloServizio", "tblCategoriaEventi", _
     '    "Classe_Evento = 'Very-Slow'")
    'Visualizza lo Status Meter
    Call acbInitMeter("AGG. EVENTI CLS_COSTO", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    Set bqry = db0.QueryDefs("qryArticoli")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    VfNe = DLookup("N_Ordini_IN_N_Ultimi_Mesi", "tblCategoriaEventi", _
                   "Classe_Evento = 'Very-Fast'")

    Do While Not brs.EOF
        'Debug.Print brs.Fields("Cod_art"), brs![Num_Eventi_12]
        Select Case brs![Num_Eventi_12]
        Case Is >= VfNe12    'Calcolo Very-Fast
            If brs![VfNe] > VfNeTab Then
                Classe_Evento = "Very-Fast"
                '            LivelloServizio = VfLs
            End If
        Case Is >= FNe12    'Calcolo Fast
            If brs![FNe] >= FNeTab Then
                Classe_Evento = "Fast"
                '            LivelloServizio = FLs
            End If
        Case Is >= MfNe12    'Calcolo Medium-Fast
            If brs![MfNe] >= MfNeTab Then
                Classe_Evento = "Medium-Fast"
                '            LivelloServizio = MfLs
            End If
        Case Is >= MNe12    'Calcolo Medium
            If brs![MNe] >= MNeTab Then
                Classe_Evento = "Medium"
                '            LivelloServizio = MLs
            End If
        Case Is >= MsNe12    'Calcolo Medium-Slow
            If brs![MsNe] >= MsNeTab Then
                Classe_Evento = "Medium-Slow"
                '            LivelloServizio = MsLs
            End If
        Case Is >= sNe12    'Calcolo Slow
            If brs![SNe] >= sNeTab Then
                Classe_Evento = "Slow"
                '            LivelloServizio = sLs
            End If
        Case Else    'Calcolo Very-Slow
            Classe_Evento = "Very-Slow"
            '            LivelloServizio = VsLs
        End Select
        Set qr1 = db0.QueryDefs("qryCategoriaEventiUpdate")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iClasse_Evento = Classe_Evento
        '        qr1!iLivelloServizio = LivelloServizio
        qr1!iClasseCosto = brs![ClasseCosto]
        qr1!iMesiCopertura = brs![MesiCopertura]
        qr1.Execute
        ' Debug.Print Classe_Evento
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    WriteToLog ("'Fine: 2 - CALCOLA CONSUMI'")
    
    ' 20160216 Puntini avanzamento
    FaseEseguita ("23")
    Me.box23.Visible = True
    ' 20160216 AggiornaAll
    If leggiChiave("aggiornaAll") = "NO" Then
        MsgBox "Finish", vbInformation, "STEP 2 ULTIMATO !!! "
    End If


Exit_cmdImportDataOracle_Click:
    Exit Sub

Err_cmdImportDataOracle_Click:
    MsgBox Err.Description
    Resume Exit_cmdImportDataOracle_Click

End Sub

Private Sub cmdImportaConsumi_Click()
    Dim ws As Workspace
    Dim Db As Database
    Dim db0 As Database    ' ******* DATA BASE CORRENTE ***********
    Dim Lconnect As String
    Dim MyQuery As String
    Dim qm As QueryDef
    Dim rst As Object
    Dim rs As Recordset
    Dim Response As Integer
    Dim conn As ADODB.Connection
    Dim intI As Double
    Dim OrigineDati As String
    ' Verificare i Dim
    Dim intMonth As Integer, intYear As Integer, intStartDate As Variant, _
        intEndDate As Variant, intLastDay As Integer, NumMesiC As Integer
    Dim DatiGen As New ADODB.Recordset
    ' Dim CatEventi As New adodb.Recordset
    ' Dim CalMag As New adodb.Recordset
    ' Dim CalGiac As New adodb.Recordset
    Dim Art As DAO.Recordset
    Dim AnnoC As String
    Dim MeseC As String
    Dim AnnoC12 As String
    Dim MeseC12 As String
    Dim AnnoF As String
    Dim MeseF As String
    Dim DataInizio As Date

    Dim Cmd As New ADODB.Command
    Dim bqry As DAO.QueryDef
    Dim brs As DAO.Recordset
    Dim ConsMeseT(1 To 12) As Variant
    Dim intCounter As Integer
    Dim TempKeyValue As Variant, NextKeyValue As Variant
    Dim VarWhereIn As Variant
    ' Dim nRecords As Integer 'Long
    On Error GoTo Err_cmdImportaConsumi_Click:

    ' 20160216 Gestione Aggiona ALL
    If leggiChiave("aggiornaAll") = "NO" Then
        Msg = "Vuoi importare i Dati nel DATA BASE ?"    ' Definisce titolo messaggio.
        Style = vbYesNo + vbQuestion + vbDefaultButton2    ' Definisce pulsante.
        Title = "MsgBox Rop Roq"    ' Definisce Titolo.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            GoTo INIZIO
        Else
            Exit Sub
        End If
    End If
INIZIO:
    ' Pulisce campo Eseguita  della tblFasiElaboraDati
    DoCmd.SetWarnings False
    DoCmd.RunSQL "UPDATE [tblFasiElaboraDati] SET [tblFasiElaboraDati].Eseguita = False;"
    DoCmd.SetWarnings True
    ' 201602
    ' Aggiorna tblFasiElaboraDati cob data inizio
    Call Tempofase("01", True) 'Scrive inizio
    
    Me.box01.Visible = True
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    NumMesiC = DatiGen.Fields("Mesi_consumo")
    intYear = DatiGen.Fields("Anno_Calcolo")
    intMonth = DatiGen.Fields("Mese_Calcolo")
    intLastDay = Day(DateAdd("m", 1, DateSerial(intYear, intMonth, 1)) - 1)
    intEndDate = DateSerial(intYear, intMonth, intLastDay)
    intEndDate = Format(intEndDate, "dd/mmm/yyyy")
    intStartDate = DateAdd("m", -NumMesiC + 1, DateSerial(intYear, intMonth, 1))
    intStartDate = Format(intStartDate, "dd/mmm/yyyy")
    AnnoC = Year(intStartDate)
    MeseC = Month(intStartDate)
    AnnoF = Year(intEndDate)
    MeseF = Month(intEndDate)
    ' Data inizio 12 mesi:
    NumMesiAbc = 12
    IntStartDateAbc = DateAdd("m", -NumMesiAbc + 1, DateSerial(intYear, intMonth, 1))
    IntStartDateAbc = Format(IntStartDateAbc, "dd/mmm/yyyy")
    AnnoC12 = Year(IntStartDateAbc)
    MeseC12 = Month(IntStartDateAbc)
    Dim qr1 As QueryDef
    Dim qr2 As QueryDef

    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media giacenze
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media giacenze
    NumMesiG = DatiGen.Fields("Mesi_giacenze") - 1
    NumAnniG = IIf(Int(NumMesiG / 12) = 1, Int(NumMesiG / 12), Int(NumMesiG / 12) + 1)
    AnnoG = Trim(Str((Val(DatiGen.Fields("Anno_calcolo")) - NumAnniG)))
    MeseG = Trim(Str(12 * NumAnniG - NumMesiG + Val(DatiGen.Fields("Mese_calcolo"))))
    MeseG = IIf(MeseG < 10, "0", "") + MeseG

    DatiGen.Close

    Response = MsgBox("Il periodo di elaborazione dati è dal " & intStartDate & _
                      " al " & intEndDate & " PER MODIFICARE ESCI E LANCIA FORM DATI GENERALI", _
                      vbYesNo, "Continua")
    If Response = vbYes Then
        ' Aggiorna tblRifCalcolo
        Cmd.ActiveConnection = conn
        Cmd.CommandText = "DELETE * FROM tblRifCalcolo"
        Cmd.Execute
        Cmd.CommandText = "INSERT INTO tblRifCalcolo ( AnnoC, MeseC, AnnoG, MeseG  ) SELECT " & _
                          AnnoC & "," & MeseC & "," & AnnoG & "," & MeseG
        Cmd.Execute
        'Aggiorna la data di ins. Dati SPOSTARE ALLA FINE
        db0.Execute "UPDATE tblDatiGenerali SET UpdateDate = #" & Format(Date, _
                                                                         "mm/dd/yyyy") & "#, " & "UpdateAnno_calcolo = " & AnnoC & ", " & _
                                                                         "UpdateMese_calcolo = " & MeseC & ", UpdateMesi_consumo = " & NumMesiC, dbFailOnError


        ' ******** Modifica 13-01-2015 per importare da Excel
        ' Verifica se importare Dati da Oracle o Excel
        strOrigineDati = DLookup("[OrigineDati]", "tblDatiGenerali")

        If strOrigineDati = "Oracle" Then
            Call ImportFromOracle
            WriteToLog ("'1 - Fine Importa da eBS'")
            ' 20160216 AggiornaAll
            If leggiChiave("aggiornaAll") = "NO" Then
                MsgBox "Finish", vbInformation, "ARTICOLI IMPORT DA ORACLE "
            End If
        ElseIf strOrigineDati = "Excel" Then
            Call ImportFromExcel
        End If
        '20160216 Chiude Fase 03
            Call FaseEseguita("03")
            Me.box03.Visible = True
    Else
        Exit Sub
    End If



Exit_cmdImportaConsumi_Click:
    Exit Sub

Err_cmdImportaConsumi_Click:
    MsgBox Err.Description

    Resume Exit_cmdImportaConsumi_Click

End Sub

Private Sub cmdAggiornaConsumi_Click()
    Dim ws As Workspace
    Dim Db As Database
    Dim db0 As Database    ' ******* DATA BASE CORRENTE ***********
    Dim Lconnect As String
    Dim MyQuery As String
    Dim qm As QueryDef
    Dim rst As Object
    Dim rs As Recordset
    Dim Response As Integer
    Dim conn As ADODB.Connection
    Dim intI As Double
    Dim OrigineDati As String
    ' Verificare i Dim
    Dim intMonth As Integer, intYear As Integer, intStartDate As Variant, _
        intEndDate As Variant, intLastDay As Integer, NumMesiC As Integer
    Dim DatiGen As New ADODB.Recordset
    ' Dim CatEventi As New adodb.Recordset
    ' Dim CalMag As New adodb.Recordset
    ' Dim CalGiac As New adodb.Recordset
    Dim Art As DAO.Recordset
    Dim AnnoC As String
    Dim MeseC As String
    Dim AnnoC12 As String
    Dim MeseC12 As String
    Dim AnnoF As String
    Dim MeseF As String
    Dim DataInizio As Date

    Dim Cmd As New ADODB.Command
    Dim bqry As DAO.QueryDef
    Dim brs As DAO.Recordset
    Dim ConsMeseT(1 To 12) As Variant
    Dim intCounter As Integer
    Dim TempKeyValue As Variant, NextKeyValue As Variant
    Dim VarWhereIn As Variant
    Dim sngPercento As Single

    ' Dim nRecords As Integer 'Long
    On Error GoTo Err_cmdImportDataOracle_Click:

    ' 20160216 Gestione Aggiona ALL
    If leggiChiave("aggiornaAll") = "NO" Then
        Msg = "Vuoi aggiornare il DATA BASE ?"    ' Definisce titolo messaggio.
        Style = vbYesNo + vbQuestion + vbDefaultButton2    ' Definisce pulsante.
        Title = "MsgBox Aggiorna DB"    ' Definisce Titolo.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            GoTo INIZIO
        Else
            Exit Sub
        End If
    End If
INIZIO:
    ' 201602
    ' Aggiorna tblFasiElaboraDati cob data inizio
    Call Tempofase("11", True)    'Scrive inizio
    Me.box11.Visible = True

    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    NumMesiC = DatiGen.Fields("Mesi_consumo")
    intYear = DatiGen.Fields("Anno_Calcolo")
    intMonth = DatiGen.Fields("Mese_Calcolo")
    intLastDay = Day(DateAdd("m", 1, DateSerial(intYear, intMonth, 1)) - 1)
    intEndDate = DateSerial(intYear, intMonth, intLastDay)
    intEndDate = Format(intEndDate, "dd/mmm/yyyy")
    intStartDate = DateAdd("m", -NumMesiC + 1, DateSerial(intYear, intMonth, 1))
    intStartDate = Format(intStartDate, "dd/mmm/yyyy")
    AnnoC = Year(intStartDate)
    MeseC = Month(intStartDate)
    AnnoF = Year(intEndDate)
    MeseF = Month(intEndDate)
    ' Data inizio 12 mesi:
    NumMesiAbc = 12
    IntStartDateAbc = DateAdd("m", -NumMesiAbc + 1, DateSerial(intYear, intMonth, 1))
    IntStartDateAbc = Format(IntStartDateAbc, "dd/mmm/yyyy")
    AnnoC12 = Year(IntStartDateAbc)
    MeseC12 = Month(IntStartDateAbc)
    Dim qr1 As QueryDef
    Dim qr2 As QueryDef

    DatiGen.Close

    'GoTo start
    ' ***************************************************************************************************
    '  INIZIO                          Salva i Codici gestiti manualmente
    ' ***************************************************************************************************
    ' 20150421 Visualizza lo Status Meter
    Call acbInitMeter("0-Codici Manuali", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    ' Cancella dati
    conn.Execute "Delete * From tblArticoliManuali"
    MyQuery = "SELECT cod_art, " & "DES_ART, " & _
              "InsManualmente " & _
              " FROM tblArticoli " & _
              "WHERE InsManualmente = 'S' "

    Set rst = db0.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
    ' Numero Transazioni
    If Not rst.BOF Then    'se ci sono record nel recordset
        rst.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = rst.RecordCount
        rst.MoveFirst
    End If
    Do While Not rst.EOF
        Set qm = db0.QueryDefs("qryArticoliManualiIsrt")
        qm!iCOD_ART = rst![Cod_art]
        qm!iDES_ART = rst![Des_art]
        qm!iInsManualmente = rst![InsManualmente]
        qm.Execute
        rst.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        '                Call acbUpdateMeter(Int(intI))
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    rst.Close
    Set rst = Nothing


    ' ***************************************************************************************************
    '                 1.3.0 Memorizza le giacenze nella tblGiacenze per avere storico
    ' ***************************************************************************************************


    WriteToLog ("'  1.3.0 Carica giacenze nella tblGiacenze'")
    MyQuery = "SELECT * from tblGiacenze where anno = """ & AnnoF & """ AND  mese = """ & MeseF & """"
    ' Verifica che non esistano già le giacenze se si le cancella
    If GetSQLRecordcount(MyQuery) > 0 Then
        conn.Execute "Delete * from tblGiacenze where anno = """ & AnnoF & """ AND  mese = """ & MeseF & """"
    End If

    MyQuery = ""

    MyQuery = "INSERT INTO tblGiacenze (Cod_Art, Anno, Mese, Giacenza) " & _
              "SELECT Cd_Art, '" & AnnoF & "', '" & MeseF & "', qt_giac FROM tblImportGiacenza"

    conn.Execute MyQuery

    ' ***************************************************************************************************
    '                 1.3 Carica anagrafiche articoli in giacenza nella tblArticoli
    ' ***************************************************************************************************

    WriteToLog ("'  1.3 Carica anagrafiche articoli in giacenza nella tblArticoli'")

    ' *** Inizio caricamento tblAnagrafica
    Dim qr As QueryDef
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    ' Inserisco Giacenza in tblArticoli
    ' Visualizza lo Status Meter
    Call acbInitMeter("3-ANAGRAFICHE ARTICOLO", True)
    'Reset lo Status Meter
    intI = 0
    conn.Execute "Delete * From tblArticoli"
    Set qr = db0.QueryDefs("qryArtAppend")
    qr.Execute
    Call acbCloseMeter
    ' ***************************************************************************************************
    '                 1.4 Crea tblArticoliSpeditiUnique
    ' ***************************************************************************************************

    WriteToLog ("'  1.4 Crea tblArticoliSpeditiUnique'")
    ' Inserisco Spedizioni in tblArticoli
    conn.Execute "Delete * From tblArticoliSpeditiUnique"
    Set qr = db0.QueryDefs("qryArtSpedUniqueIsrt")
    qr.Execute

    ' Creo una tbl con Anagrafica articoli unica

    'Visualizza lo Status Meter
    '    Call acbInitMeter("4-ANAGRAFICHE ARTICOLO SPEDIZIONI", True)
    '    'Reset lo Status Meter
    '    intI = 0
    'MyQuery = "SELECT DISTINCT tblImportSpedito.COD_ART, " & _
     '        "tblImportSpedito.DESCRIZIONE, tblImportSpedito.LEAD_TIME, tblImportSpedito.ROP, " & _
     '        "tblImportSpedito.MAX_MINMAX_QUANTITY, tblImportSpedito.ROQ, " & _
     '        "tblImportSpedito.MAXIMUM_ORDER_QUANTITY,tblImportSpedito.CS_CSC " & _
     '        "FROM tblImportSpedito;"
    'Set rst = db0.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
    '    Do While Not rst.EOF
    '       Set qm = db0.QueryDefs("qryArtSpedizioniUniqueIsrt")
    '       qm!iCOD_ART = rst![Cod_art]
    '       qm!iDESCRIZIONE = rst![Descrizione]
    '       qm!iLEAD_TIME = rst![Lead_time]
    '       qm!iROP = rst![ROP]
    '       qm!iMAX_MINMAX_QUANTITY = rst![MAX_MINMAX_QUANTITY]
    '       qm!iROQ = rst![ROQ]
    '       qm!iMAXIMUM_ORDER_QUANTITY = rst![MAXIMUM_ORDER_QUANTITY]
    '       qm!iCS_CSC = rst![Cs_Csc]
    '       qm.Execute
    '       rst.MoveNext
    '       ' Aggiorna lo Status Meter --
    '        intI = intI + (1 / 1000)
    '        Call acbUpdateMeter(Int(intI))
    '        'Si ferma per 1 sec
    '         DoEvents
    '    Loop
    '    rst.Close
    '    Call acbCloseMeter

    ' ***************************************************************************************************
    '                 1.5 Carica anagrafiche articoli spediti nella tblArticoli
    ' ***************************************************************************************************


    WriteToLog ("'  1.5 Carica anagrafiche articoli spediti nella tblArticoli'")
    Call acbInitMeter("5-ANAG ART SPED", True)
    ' Test 1 inizio
    Set qr = db0.QueryDefs("qryArticoliImportSpeditoIsrt")
    qr.Execute
    Call acbCloseMeter
    ' Test 1 fine da canc

    ''Visualizza lo Status Meter
    '    Call acbInitMeter("5-ISRT ANAGRAFICHE ARTICOLO SPEDIZIONI", True)
    '    'Reset lo Status Meter
    '    intI = 0
    'Set bqry = db0.QueryDefs("qryArticoliImportSpedito")
    'Set brs = bqry.OpenRecordset
    '  Do While Not brs.EOF
    '    Set qr1 = db0.QueryDefs("qryImportSpeditoIsrt")
    '    qr1!iCOD_ART = brs![Cod_art]
    '    qr1!iDES_ART = brs![Descrizione]
    '    qr1!iLEAD_TIME = brs![Lead_time]
    '    qr1!iROP = brs![ROP]
    '    qr1!iMAX_MINMAX_QUANTITY = brs![MAX_MINMAX_QUANTITY]
    '    qr1!iROQ = brs![ROQ]
    '    qr1!iMAXIMUM_ORDER_QUANTITY = brs![MAXIMUM_ORDER_QUANTITY]
    '    qr1!iCS_CSC = brs![Cs_Csc]
    '    qr1.Execute
    '    brs.MoveNext
    '' Aggiorna lo Status Meter
    '        intI = intI + (1 / 300)
    '        Call acbUpdateMeter(Int(intI))
    '    ' Attende un secondo
    '    DoEvents
    '  Loop
    ''Close lo Status Meter
    '    Call acbCloseMeter
    '  brs.Close


    ' ***************************************************************************************************
    '                 1.6 Carica anagrafiche articoli ordinato nella tblArticoliCOrdersUnique
    ' ***************************************************************************************************

    WriteToLog ("'  1.6 Carica anagrafiche articoli ordinato nella tblArticoliCOrdersUnique'")
    ' *** Inserisco tblCOrders in tblArticoli ***
    'Visualizza lo Status Meter
    Call acbInitMeter("6-ANAGRAFICHE ARTICOLO ORDINI", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    ' Cancella tblArticoliCOrdersUnique
    conn.Execute "Delete * From tblArticoliCOrdersUnique"
    ' Inserisce gli articoli in tblArticoliCOrdersUnique
    MyQuery = "SELECT DISTINCT tblCOrders.COD_ART, tblCOrders.DESCRIZIONE, " & _
              "tblCOrders.LEAD_TIME, tblCOrders.ROP, tblCOrders.MAX_MINMAX_QUANTITY, " & _
              "tblCOrders.ROQ, tblCOrders.MAXIMUM_ORDER_QUANTITY, tblCOrders.CS_CSC  " & _
              "FROM tblCOrders;"
    Set rst = db0.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
    ' Numero Transazioni
    If Not rst.BOF Then    'se ci sono record nel recordset
        rst.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = rst.RecordCount
        rst.MoveFirst
    End If

    Do While Not rst.EOF
        Set qm = db0.QueryDefs("qryArtCordersUniqueIsrt")
        qm!iCOD_ART = rst![Cod_art]
        qm!iDESCRIZIONE = rst![Descrizione]
        qm!iLEAD_TIME = rst![Lead_time]
        qm!iROP = rst![ROP]
        qm!iMAX_MINMAX_QUANTITY = rst![MAX_MINMAX_QUANTITY]
        qm!iROQ = rst![ROQ]
        qm!iMAXIMUM_ORDER_QUANTITY = rst![MAXIMUM_ORDER_QUANTITY]
        qm!iCS_CSC = rst![Cs_Csc]
        qm.Execute
        rst.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        '                Call acbUpdateMeter(Int(intI))
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    rst.Close
    Call acbCloseMeter
    'start:

    WriteToLog ("'  1.7 Carica anagrafiche articoli gestiti manualmente nella tblArticoli'")

    ' *********************************************************************************
    ' Carica in Anagrafica articoli gli articoli nuovi presenti in tblArticoliManuali
    ' *********************************************************************************
    'Visualizza lo Status Meter
    Call acbInitMeter("7 - ISRT ANAGRAFICHE MANUALI", True)
    ' 20150421 Staus Meter Modifica
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0

    Set bqry = db0.QueryDefs("qryArticoliManuali")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    Do While Not brs.EOF
        Set qr1 = db0.QueryDefs("qryArticoliIsrt")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iDES_ART = brs![Des_art]
        qr1!iInsManualmente = brs![InsManualmente]
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        '                Call acbUpdateMeter(Int(intI))
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close

    WriteToLog ("'  1.8 Carica anagrafiche articoli ordinato nella tblArticoli'")

    ' ************************************************************************
    ' Carica in Anagrafica articoli gli articoli nuovi presenti in tblCorders
    ' ************************************************************************
    'Visualizza lo Status Meter
    Call acbInitMeter("8-ISRT ANAGRAFICHE ARTICOLO ORDINI", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    ' Carica in Anagrafica articoli gli articoli nuovi presenti in tblCorders
    Set bqry = db0.QueryDefs("qryArticoliImportCOrders")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If

    Do While Not brs.EOF
        Set qr1 = db0.QueryDefs("qryImportSpeditoIsrt")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iDES_ART = brs![Descrizione]
        qr1!iLEAD_TIME = brs![Lead_time]
        qr1!iROP = brs![ROP]
        qr1!iMAX_MINMAX_QUANTITY = brs![MAX_MINMAX_QUANTITY]
        qr1!iROQ = brs![ROQ]
        qr1!iMAXIMUM_ORDER_QUANTITY = brs![MAXIMUM_ORDER_QUANTITY]
        qr1!iCS_CSC = brs![Cs_Csc]
        qr1.Execute
        Set qr2 = db0.QueryDefs("qryArticoliSegnalazioniIsrt")
        qr2!iCOD_ART = brs![Cod_art]
        qr2!iUser = "System"
        qr2!IComments = "Nuovo Codice"
        qr2!iDate = Format(Date, "mm/dd/yyyy")
        qr2.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close
    
    ' 201602 Chiude Fase 11
    Call FaseEseguita("11")
    
    ' 201602 Inizia Fase 12
    Call Tempofase("12", True)
    
    
    
    ' start:
    WriteToLog ("'  1.9 Calcola Consumi'")

    ' *** Inizio calcolo CONSUMI Magazzino

    Set db0 = CurrentDb
    Set bqry = db0.QueryDefs("qryCalcoloMag")
    conn.Execute "Delete * From tblConsumi"
    ' Imposta i parametri qry parametrica
    Dim Counter
    Counter = 0
    ' Per ogni periodo carica i dati nella tblConsumi
    While Counter < NumMesiC
        intLastDay = Day(DateAdd("m", 1, DateSerial(intYear, intMonth - Counter, _
                                                    1)) - 1)
        intEndDate = DateSerial(intYear, intMonth - Counter, intLastDay)
        intEndDate = Format(intEndDate, "dd/mmm/yyyy")
        intStartDate = DateAdd("m", -Counter, DateSerial(intYear, intMonth, 1))
        intStartDate = Format(intStartDate, "dd/mmm/yyyy")
        bqry.Parameters("DataInizio").Value = Format(intStartDate, "dd/mm/yyyy")
        bqry.Parameters("DataFine").Value = Format(intEndDate, "dd/mm/yyyy")
        Set brs = bqry.OpenRecordset

        ' 20150420 Staus Meter Modifica
        ' Numero Transazioni
        If Not brs.BOF Then    'se ci sono record nel recordset
            brs.MoveLast    ' necessario per determinare l'attuale numero di record
            intMassimo = brs.RecordCount
            brs.MoveFirst
        End If
        'Visualizza lo Status Meter
        Call acbInitMeter("9-CALC CONS " & Month(intEndDate) & "-" & Year(intEndDate), True)
        'Reset lo Status Meter
        intI = 0
        sngPercento = 0

        'Inserisco i dati in tblConsumi
        Do While Not brs.EOF
            Set qm = db0.QueryDefs("qryConsumiIsrt")
            ' Cod_Art, Anno, Mese, Consumo, N_Spedito_Mese,Stock_Finale ;
            qm!iCOD_ART = brs.Fields("Cod_Art")
            qm!iAnno = Year(intStartDate)
            qm!iMese = Month(intStartDate)
            qm!iConsumo = brs.Fields("Consumo")
            qm!iN_Spedito_Mese = brs.Fields("N_Spedito_Mese")
            qm.Execute
            brs.MoveNext
            ' Aggiorna lo Status Meter
            sngPercento = intI / (intMassimo / 100)
            If sngPercento >= 1 Then
                Call acbUpdateMeter(Int(sngPercento))
            End If
            intI = intI + 1
        Loop
        brs.Close
        Counter = Counter + 1
    Wend
    'Close lo Status Meter
    Call acbCloseMeter
    
    ' 201602 Chiude Fase 12
    Call FaseEseguita("12")
     '20160216 Puntini Avanzamento
    Me.box12.Visible = True
    
    ' 201602 Inizia Fase 13
    Call Tempofase("13", True)
      
    'start:
    ' ******************************************************************************************************
    ' ***   CALCOLO ARTICOLI CORRELATI 28/03/2014 **********************************************************
    ' ******************************************************************************************************
    ' Visualizza puntatore mouse a clessidra
    Screen.MousePointer = vbHourGlass
    '20160216
    ' Apre la formCorrelati per indicare che sta calcolando questi articoli
    DoCmd.OpenForm "frmCorrelati"

    DataInizio = DateSerial(AnnoC12, MeseC12 - 1, 1)
    'start:
    'Cancella i dati tblConsumiCorrelati
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * from  tblConsumiCorrelati"
    WriteToLog ("'  1.9.1 Articoli Correlati'")

    DoCmd.SetWarnings True

    ' Debug.Print DataInizio         ' **** DATA INIZIO
    NextKeyValue = ""
    TempKeyValue = ""
    j = 0
    VarWhereIn = Null

    Set rst = DBEngine(0)(0).OpenRecordset("SELECT cod_art, cod_art_correlato FROM tblArticoliStato " & _
                                           "WHERE len(cod_art_correlato) > 1 " & _
                                           "Order by cod_art_correlato;")
    Do While Not rst.EOF
        NextKeyValue = rst.Fields(1).Value
        If TempKeyValue <> NextKeyValue Then    ' Testa quando cambia codice
            If j > 0 Then
                ' Aggiunge alla stringa l'articolo correlato (loop precedente)
                VarWhereIn = VarWhereIn & "'" & TempKeyValue & "'"
                '           Debug.Print VarWhereIn
                ' INSERISCO NELLA TABELLA tblConsumiCorrelati i valori sommati
                Varwhere = "[Cod_Art] IN (" & VarWhereIn & ") " & _
                           "AND DateSerial([Anno],[Mese],1) >= DateSerial(" & AnnoC12 & "," & MeseC12 & ",1)"
                '              Debug.Print VarWhere
                '               WriteToLog (VarWhere)

                strSQL = "INSERT INTO tblConsumiCorrelati" _
                         & " ([Cod_Art], [Anno],[Mese], [Consumo], [N_Spedito_Mese]) " _
                         & "SELECT '" & TempKeyValue _
                         & "', Anno, Mese, Sum(Consumo), Sum(N_Spedito_Mese) " _
                         & "FROM tblConsumi WHERE " & Varwhere _
                         & " GROUP BY tblConsumi.anno, tblConsumi.mese;"
                '               Debug.Print strSql
                CurrentDb.Execute strSQL
                strSQL = ""
                VarWhereIn = ""
            End If
        End If
        TempKeyValue = Trim(rst.Fields(1).Value)
        VarWhereIn = VarWhereIn & "'" & rst.Fields(0).Value & "', "
        '       Debug.Print VarWhereIn
        j = j + 1
        rst.MoveNext
    Loop
    rst.Close
    ' CARICA ULTIMO RECORD
    ' Ho lasciato questa vecchia routine di calcolo
    VarWhereIn = VarWhereIn & "'" & TempKeyValue & "'"
    '           Debug.Print VarWhereIn
    For intI = 1 To 12
        strData = DateAdd("m", intI, DataInizio)
        '              Debug.Print strData
        Mese = Month(strData)
        Anno = Year(strData)
        Varwhere = "[Cod_Art] IN (" & VarWhereIn & ") " & _
                   "AND [Anno] = '" & Anno & "' AND [Mese] = " & "'" & Mese & "'"
        '              Debug.Print VarWhere

        ConsumoMese = DSum("[Consumo]", "tblConsumi", Varwhere)
        NSpeditoMese = DSum("[N_Spedito_Mese]", "tblConsumi", Varwhere)
        '              Debug.Print ConsumoMese
        If Not IsNothing(ConsumoMese) Then
            strSQL = "INSERT INTO tblConsumiCorrelati" _
                     & " ([Cod_Art], [Anno],[Mese], [Consumo], [N_Spedito_Mese]) " _
                     & "VALUES ('" & TempKeyValue _
                     & "', '" & Anno _
                     & "', '" & Mese _
                     & "', '" & ConsumoMese _
                     & "', '" & NSpeditoMese & "');"
            '             Debug.Print strSql
            CurrentDb.Execute strSQL
        End If
    Next intI
    ' Visualizza puntatore mouse std
    Screen.MousePointer = vbDefault
    ' Cancella i consumi dei nuovi articoli correlati
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * " & _
                 "FROM tblConsumi " & _
                 "WHERE EXISTS ( " & _
                 "SELECT NULL " & _
                 "FROM tblConsumiCorrelati " & _
                 "WHERE COD_ART = tblConsumi.COD_ART " & _
                 ");"
    ' Inserisce in tblConsumi gli articoli di tblConsumiCorrelati
    DoCmd.RunSQL "INSERT INTO tblConsumi" _
                 & " ([Cod_Art], [Anno],[Mese], [Consumo], [N_Spedito_Mese]) " _
                 & "SELECT [Cod_Art], [Anno],[Mese], [Consumo], [N_Spedito_Mese] " _
                 & "FROM tblConsumicorrelati;"
    DoCmd.SetWarnings True
    ' Chiude Form Correlati
    DoCmd.Close acForm, "frmCorrelati"
    ' *** FINE CALCOLO ARTICOLI CORRELATI
    WriteToLog ("'  1.10 Inserisce Consumi in tblArticoli'")

    '***************************************
    ' *** Memorizzo i consumi in tblArticoli
    '***************************************
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
start:
    'Visualizza lo Status Meter
    Call acbInitMeter("10-AGGIORNO CONSUMI", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryCalcoloConsumo")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If

    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art"), brs.Fields("SConsumo"), brs.Fields("SSpedito"), brs.Fields("Num_Eventi")
        'Memorizzo i dati nella tblarticoli

        Set qr1 = db0.QueryDefs("qryArticoliUpdate")
        qr1!iCOD_ART = brs![Cod_art]
        qr1!iSConsumo = brs![SConsumo]
        qr1!iSSpedito = brs![SSpedito]
        qr1!iNum_Eventi = brs![Num_Eventi]
        qr1!iAnno_Calcolo = AnnoC
        qr1!iMese_Calcolo = MeseC
        qr1!iMesi_Consumo = NumMesiC
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close


    'Start:
    WriteToLog ("'  1.11 Calcola Dev Std'")

    ' ******** CALCOLO DEVIAZIONE STANDARD
    'Visualizza lo Status Meter
    Call acbInitMeter("11 CALCOLO DEV STD", True)
    'Reset lo Status Meter
    intI = 0
    intStartDate = DateAdd("m", -11, DateSerial(intYear, _
                                                intMonth, 1))
    strSQL = "TRANSFORM Nz(Sum([tblConsumi].[Consumo]),0) AS Consumo "
    strSQL = strSQL & "SELECT tblConsumi.Cod_Art "
    strSQL = strSQL & "FROM tblConsumi "
    strSQL = strSQL & "where [Anno] + Format([Mese], '00') >= " & _
             Year(intStartDate) & Format(Month(intStartDate), "00")
    ' strSql = strSql & " AND tblConsumi.Cod_Art = '0000107048A' "
    strSQL = strSQL & " GROUP BY tblConsumi.Cod_Art "
    strSQL = strSQL & "PIVOT [Anno] + Format([Mese], '00');"

    'Set bqry = Db0.CreateQueryDef("My_Query", strSQL)
    Set brs = db0.OpenRecordset(strSQL, dbOpenDynaset)
    ' Numero Transazioni
    ' Calcolato approssimativamente come Tot_n_articoli da funzione CountOfActivePart() per i 12 mesi
    intMassimo = CountOfActivePart()
    brs.MoveFirst

    'Reset lo Status Meter
    intI = 0
    sngPercento = 0

    Do While Not brs.EOF
        'Debug.Print brs.Fields("Cod_Art")
        For intCounter = 1 To 12    '12 per i 12 mesi
            '20/03/2013 Morotti: Modificato stringa seguente
            ' intStartDate = DateAdd("m", -NumMesiC + intCounter, DateSerial(intYear, _
              intMonth, 1))
            intStartDate = DateAdd("m", -12 + intCounter, DateSerial(intYear, _
                                                                     intMonth, 1))

            On Error Resume Next
            Dim varDummy As Variant
            'Debug.Print brs.Fields(Year(intStartDate) & Format(Month(intStartDate), _
             "00"))

            varDummy = brs.Fields(Year(intStartDate) & Format(Month(intStartDate), "00"))
            If Not IsNothing(varDummy) Then

                ' Debug.Print brs.Fields(Year(intStartDate) & Format(Month(intStartDate), _
                  "00"))
                ConsMeseT(intCounter) = brs.Fields(Year(intStartDate) & _
                                                   Format(Month(intStartDate), "00"))
            Else
                Response = MsgBox("Mancano i dati del periodo " & Year(intStartDate) & Format(Month(intStartDate), "00") & vbCrLf & _
                                  "Vuoi Continuare??", vbYesNo, "Continuare")
                If Response = vbNo Then
                    GoTo FINE
                End If
            End If

        Next intCounter
        ' Debug.Print RStDev(ConsMeseT(1), ConsMeseT(2), ConsMeseT(3), ConsMeseT(4), _
          ConsMeseT(5), ConsMeseT(6), ConsMeseT(7), ConsMeseT(8), ConsMeseT(9), _
          ConsMeseT(10), ConsMeseT(11), ConsMeseT(12))
        Set qr1 = db0.QueryDefs("qryDevStdUpdate")
        qr1!iCOD_ART = brs.Fields("Cod_Art")
        qr1!iDevStdConsumoMese = RStDev(ConsMeseT(1), ConsMeseT(2), ConsMeseT(3), ConsMeseT(4), _
                                        ConsMeseT(5), ConsMeseT(6), ConsMeseT(7), ConsMeseT(8), ConsMeseT(9), _
                                        ConsMeseT(10), ConsMeseT(11), ConsMeseT(12))
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter --Divido per 10 perchè sono circa 4000 righe
        sngPercento = intI / (intMassimo / 100)
        '                Call acbUpdateMeter(Int(intI))
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If

        intI = intI + 1
    Loop
FINE:
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close

    ' ********************************************************************
    ' Marco Morotti 18-11-2015
    ' Update tbltblArticoliStato con Stato_Articolo STRATEGICO per le
    ' Classi Merceologiche nella SELECT
    ' ********************************************************************
    DoCmd.SetWarnings False
    strSQL = "UPDATE tblArticoliStato INNER JOIN tblArticoli ON tblArticoli.Cod_art = tblArticoliStato.Cod_Art " & _
             "   SET tblArticoliStato.ID_StatoArticolo = 9 " & _
             "WHERE (((tblArticoli.Categ_Merc) Like 'A4*' Or " & _
             "      (tblArticoli.Categ_Merc) Like 'B6*' Or " & _
             "      (tblArticoli.Categ_Merc) Like 'N8*' Or " & _
             "      (tblArticoli.Categ_Merc) Like 'D4*' Or " & _
             "      (tblArticoli.Categ_Merc) = 'S10101')) " & _
             "  AND tblArticoliStato.ID_StatoArticolo IS NULL; "
    DoCmd.RunSQL strSQL
    ' Insert  tblArticoliStato con Stato_Articolo STRATEGICO per le
    ' Classi Merceologiche nella SELECT
    ' ********************************************************************
    strSQL = " INSERT INTO tblArticoliStato"
    strSQL = strSQL & "   SELECT tblArticoli.COD_ART AS COD_ART,"
    strSQL = strSQL & "          9                   AS ID_StatoArticolo,"
    strSQL = strSQL & "          ""Inserito "" & now() AS [NOTE]"
    strSQL = strSQL & "     FROM tblArticoli"
    strSQL = strSQL & "    WHERE (((Exists"
    strSQL = strSQL & "           (select Cod_Art"
    strSQL = strSQL & "                from tblArticoliStato"
    strSQL = strSQL & "               where tblArticoli.Cod_art = tblArticoliStato.Cod_art)) ="
    strSQL = strSQL & "          False))"
    strSQL = strSQL & "      AND (((tblArticoli.Categ_Merc) Like 'A4*' Or"
    strSQL = strSQL & "          (tblArticoli.Categ_Merc) Like 'B6*' Or"
    strSQL = strSQL & "          (tblArticoli.Categ_Merc) Like 'N8*' Or"
    strSQL = strSQL & "          (tblArticoli.Categ_Merc) Like 'D4*' Or"
    strSQL = strSQL & "          (tblArticoli.Categ_Merc) = 'S10101'));"
    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
    WriteToLog ("'1 - Fine Importa Consumi'")
    ' 20160216 AggiornaAll
    If leggiChiave("aggiornaAll") = "NO" Then
        MsgBox "Finish - STEP 1", vbInformation, "ARTICOLI IMPORT DA ORACLE "
    End If
    '201602 Chiude fase 13
    FaseEseguita ("13")
    '20160216 Puntini Avanzamento
    Me.box13.Visible = True
Exit_cmdImportDataOracle_Click:
    Exit Sub

Err_cmdImportDataOracle_Click:
    MsgBox Err.Description

    Resume Exit_cmdImportDataOracle_Click
End Sub
Private Sub cmdOnHand_Click()
    Dim ws As Workspace
    Dim Db As Database
    Dim db0 As Database    ' ******* DATA BASE CORRENTE ***********
    Dim Lconnect As String
    Dim MyQuery As String
    Dim qm As QueryDef
    Dim rst As Object
    Dim rs As Recordset
    Dim Response As Integer
    Dim conn As ADODB.Connection
    Dim intI As Double

    ' *** CONNESSIONE a Oracle

    ' On Error GoTo Err_Execute

    'Use {Microsoft ODBC for Oracle} ODBC connection
    'Lconnect = "ODBC;DSN=sun3000.scmgroup.com;UID=VPERAZZINI;PWD=BIC;SERVER=sun3000.scmgroup.com"
    Lconnect = LeggiOdbcConnect
    'Point to the current workspace
    Set ws = DBEngine.Workspaces(0)

    'Connect to Oracle
    Set Db = ws.OpenDatabase("", False, True, Lconnect)
    ' Setto il tempo di QueryTimeOut a 240 min
    Db.QueryTimeout = 240


    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    ' GoTo Start
    ' *** Inizio caricamento GIACENZA.sql
    '
    'Visualizza lo Status Meter
    ' 201602
  ' Aggiorna tblFasiElaboraDati cob data inizio
  Call Tempofase("61", True) 'Scrive inizio
    
    
    Call acbInitMeter("1-Aggiorna DISPONIBILE da ORACLE", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    ' Cancella dati
    conn.Execute "Delete * From tblInventoryControl"

    MyQuery = "SELECT CD_ART, QTY_SALE, QTY_STOCK, QTY_ACQ " & _
              " FROM on_hand; "


    Set rst = Db.OpenRecordset(MyQuery, dbOpenSnapshot, dbSQLPassThrough)
    ' Numero Transazioni
    If Not rst.BOF Then    'se ci sono record nel recordset
        rst.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = rst.RecordCount
        rst.MoveFirst
    End If
    Do While Not rst.EOF
        Set qm = db0.QueryDefs("qryInventoryControlIsrt")
        qm!iCD_ART = rst![Cd_Art]
        If IsNull(rst![QTY_SALE]) Then
            QTY_SALE = 0
        Else
            QTY_SALE = rst![QTY_SALE]
        End If
        qm!iQTY_SALE = QTY_SALE
        If IsNull(rst![QTY_STOCK]) Then
            QTY_STOCK = 0
        Else
            QTY_STOCK = rst![QTY_STOCK]
        End If
        qm!iQTY_STOCK = QTY_STOCK
        If IsNull(rst![QTY_ACQ]) Then
            QTY_ACQ = 0
        Else
            QTY_ACQ = rst![QTY_ACQ]
        End If
        qm!iQTY_ACQ = QTY_ACQ
        qm!iUpdateDate = Format(Date, "mm/dd/yyyy")
        ' Debug.Print ("Costo: " & rst![Cs_Csc])
        qm.Execute
        rst.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    rst.Close

    'Close lo Status Meter
    Call acbCloseMeter
    FaseEseguita ("61")
    Me.box61.Visible = True
End Sub

Private Sub cmdRopRoq_Click()
'DoCmd.SetWarnings False
'DoCmd.Hourglass True
' CALCOLO ROP ROQ SECONDO TABELLA LS
'Definizioni
'*************************************************************************
    Dim db0 As Database
    Dim ClasseA1Perc, ClasseA2Perc, ClasseA3Perc, ClasseA4Perc As Variant
    Dim ClasseB1Perc, ClasseB2Perc, ClasseB3Perc As Variant
    Dim ClasseC1Perc, ClasseC2Perc As Variant
    Dim DatiGen As New ADODB.Recordset
    Dim strSQL As String
    Dim qr1 As QueryDef
    Dim conn As ADODB.Connection
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    Dim rs As New ADODB.Recordset
    Dim brs As DAO.Recordset
    Dim bqry As DAO.QueryDef
    Dim ClasseA1LsPerc As Variant
    Dim ClasseA2LsPerc As Variant
    Dim ClasseA3LsPerc As Variant
    Dim ClasseA4LsPerc As Variant
    Dim ClasseB1LsPerc As Variant
    Dim ClasseB2LsPerc As Variant
    Dim ClasseB3LsPerc As Variant
    Dim ClasseC1LsPerc As Variant
    Dim ClasseC2LsPerc As Variant
    Dim Msg As String
    Dim Style As String
    Dim Title As String
    Dim Response As String
    Dim intI As Variant

    WriteToLog ("'Inizio: 3 - CALCOLO ROP e ROQ'")

    ' 20160216 Gestione Aggiona ALL
    If leggiChiave("aggiornaAll") = "NO" Then
        Msg = "Vuoi Calcolare Rop e Roq ?"    ' Definisce titolo messaggio.
        Style = vbYesNo + vbQuestion + vbDefaultButton2    ' Definisce pulsante.
        Title = "MsgBox Rop Roq"    ' Definisce Titolo.
        Response = MsgBox(Msg, Style, Title)
        If Response = vbYes Then
            GoTo INIZIO
        Else
            Exit Sub
        End If
    End If
INIZIO:
    ' 201602
  ' Aggiorna tblFasiElaboraDati cob data inizio
  Call Tempofase("31", True) 'Scrive inizio
  

    '*************************************************************************
    'memorizzo le soglie di Classi A1- C2 da tblDatiGenerali
    DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
    ClasseA1LsPerc = DatiGen.Fields("ClasseA1LsPerc")
    ClasseA2LsPerc = DatiGen.Fields("ClasseA2LsPerc")
    ClasseA3LsPerc = DatiGen.Fields("ClasseA3LsPerc")
    ClasseA4LsPerc = DatiGen.Fields("ClasseA4LsPerc")
    ClasseB1LsPerc = DatiGen.Fields("ClasseB1LsPerc")
    ClasseB2LsPerc = DatiGen.Fields("ClasseB2LsPerc")
    ClasseB3LsPerc = DatiGen.Fields("ClasseB3LsPerc")
    ClasseC1LsPerc = DatiGen.Fields("ClasseC1LsPerc")
    ClasseC2LsPerc = DatiGen.Fields("ClasseC2LsPerc")
    DatiGen.Close


    ' **************************************************************
    '                  Calcolo Classe A1-C1 Consumi VALORE
    ' **************************************************************

    'Cancella i dati tblConsumiParetoValoreLs
    DoCmd.SetWarnings False
    DoCmd.RunSQL "DELETE * from  tblConsumiParetoValoreLs"


    strSQL = "INSERT INTO tblConsumiParetoValoreLs " & "SELECT Cod_Art, " & _
             "SConsumo_12 * Cs_Csc As SConsumoValore " & "FROM [tblArticoli] " & _
             "WHERE tblArticoli.SConsumo_12 > 0 " & "AND tblArticoli.Cs_Csc > 0 " & _
             "Order By (SConsumo_12 * Cs_Csc) desc"

    DoCmd.RunSQL strSQL
    DoCmd.SetWarnings True
    Set db0 = CurrentDb
    Set conn = CurrentProject.Connection
    'Visualizza lo Status Meter
    Call acbInitMeter("AGG_CL ABCDEF CONSUMI", True)
    'Reset lo Status Meter
    intI = 0
    sngPercento = 0
    'memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
    Set bqry = db0.QueryDefs("qryConsumoParetoValoreLsCalc")
    Set brs = bqry.OpenRecordset
    ' Numero Transazioni
    If Not brs.BOF Then    'se ci sono record nel recordset
        brs.MoveLast    ' necessario per determinare l'attuale numero di record
        intMassimo = brs.RecordCount
        brs.MoveFirst
    End If
    'Aggiorno i dati in tblArticoli
    Do While Not brs.EOF
        '
        ' Debug.Print brs.Fields("Cod_art"), brs.Fields("SConsumo"), brs.Fields("SSpedito"), brs.Fields("Num_Eventi")
        'Memorizzo i dati nella tblarticoli
        Set qr1 = db0.QueryDefs("qryArticoliParetoValoreLsUpdate")
        qr1!iCOD_ART = brs![Cod_art]
        Select Case brs![CumPct]
        Case Is <= ClasseA1LsPerc / 100
            qr1!iAbcConsumoValoreLs = "A1"
            qr1!iPctConsumoValoreLs = brs![CumPct]
        Case Is <= (ClasseA1LsPerc + ClasseA2LsPerc) / 100
            qr1!iAbcConsumoValoreLs = "A2"
            qr1!iPctConsumoValoreLs = brs![CumPct]
        Case Is <= (ClasseA1LsPerc + ClasseA2LsPerc + ClasseA3LsPerc) / 100
            qr1!iAbcConsumoValoreLs = "A3"
            qr1!iPctConsumoValoreLs = brs![CumPct]
        Case Is <= (ClasseA1LsPerc + ClasseA2LsPerc + ClasseA3LsPerc + _
                    ClasseA4LsPerc) / 100
            qr1!iAbcConsumoValoreLs = "A4"
            qr1!iPctConsumoValoreLs = brs![CumPct]
        Case Is <= (ClasseA1LsPerc + ClasseA2LsPerc + ClasseA3LsPerc + _
                    ClasseA4LsPerc + ClasseB1LsPerc) / 100
            qr1!iAbcConsumoValoreLs = "B1"
            qr1!iPctConsumoValoreLs = brs![CumPct]
        Case Is <= (ClasseA1LsPerc + ClasseA2LsPerc + ClasseA3LsPerc + _
                    ClasseA4LsPerc + ClasseB1LsPerc + ClasseB2LsPerc) / 100
            qr1!iAbcConsumoValoreLs = "B2"
            qr1!iPctConsumoValoreLs = brs![CumPct]
        Case Is <= (ClasseA1LsPerc + ClasseA2LsPerc + ClasseA3LsPerc + _
                    ClasseA4LsPerc + ClasseB1LsPerc + ClasseB2LsPerc + ClasseB3LsPerc) / 100
            qr1!iAbcConsumoValoreLs = "B3"
            qr1!iPctConsumoValoreLs = brs![CumPct]
        Case Is <= (ClasseA1LsPerc + ClasseA2LsPerc + ClasseA3LsPerc + _
                    ClasseA4LsPerc + ClasseB1LsPerc + ClasseB2LsPerc + ClasseB3LsPerc + ClasseC1LsPerc) / 100
            qr1!iAbcConsumoValoreLs = "C1"
            qr1!iPctConsumoValoreLs = brs![CumPct]
        Case Else
            qr1!iAbcConsumoValoreLs = "C2"
            qr1!iPctConsumoValoreLs = brs![CumPct]
        End Select
        qr1.Execute
        brs.MoveNext
        ' Aggiorna lo Status Meter
        sngPercento = intI / (intMassimo / 100)
        If sngPercento >= 1 Then
            Call acbUpdateMeter(Int(sngPercento))
        End If
        intI = intI + 1
    Loop
    'Close lo Status Meter
    Call acbCloseMeter
    brs.Close

    ' *** AGGIORNA CLASSE D per CONSUMO LS***
    db0.Execute _
            "Update tblArticoli set tblArticoli.AbcConsumoValoreLs = 'D' WHERE tblArticoli.AbcConsumoValoreLs Is Null; ", _
            dbFailOnError
' 201602 Chiude Fase 31
    Call FaseEseguita("31")
    
 '20160216 Puntini Avanzamento
    Me.box31.Visible = True
    
    ' 201602 Inizia Fase 32
    Call Tempofase("32", True)



    ' **************************************************************
    '                  Calcolo Rop e Roq
    ' **************************************************************
    ' richiama ModSetRopRoq
    ' Start:
    
    
    WriteToLog ("'3.1 - SetRopRoq'")
    ' ********** LANCIO MODULO CALCOLO ROP ROQ
    SetRopRoq
    
     '20160216 Puntini avanzamento
    FaseEseguita ("32")
    Me.box32.Visible = True
    
    FaseEseguita ("33")
    Me.box33.Visible = True
    
    WriteToLog ("'Fine: 3 - CALCOLO ROP e ROQ'")

    ' MsgBox "Calcolo Rop e Roq ultimato!", vbOKOnly, ""
    MsgBox "Calcolo Rop e Roq ultimato!"
End Sub

Private Sub Corpo_Click_click()
    Call cmdImportaConsumi_Click
    Call cmdEventi_Click
    Call cmdRopRoq_Click
End Sub

'Private Sub cmdArtCorrelati_Click()
'Dim db As DAO.Database
'Dim rst As DAO.Recordset
'Dim TempKeyValue As Variant, NextKeyValue As Variant
'Dim intMonth As Integer, intYear As Integer, intEndDate As Variant, DataTest As Date, DataInizio As Date
'Dim VarWhereIn As Variant
'Dim DatiGen As New ADODB.Recordset
'
'' Vuoi calcolare i dati?
'Msg = "Vuoi Calcolare i Consumi Articoli Sostituiti?" ' Define message.
'Style = vbYesNo + vbQuestion + vbDefaultButton2 ' Define buttons.
'Title = "MsgBox Import" ' Define title.
'Response = MsgBox(Msg, Style, Title)
'If Response = vbYes Then GoTo Inizio Else ' User chose Yes.
'Exit Sub
'
'Inizio:
'
'' Visualizza puntatore mouse a clessidra
'    Screen.MousePointer = vbHourGlass
'
''Cancella i dati tblConsumiPareto
'DoCmd.SetWarnings False
'DoCmd.RunSQL "DELETE * from  tblConsumiCorrelati"
'
'DoCmd.SetWarnings True
'
'' memorizzo il periodo di partenza ed il numero di mesi per calcolo media consumi
''   Data inizio finestra
'Set conn = CurrentProject.Connection
'DatiGen.Open "tblDatiGenerali", conn, adOpenKeyset, adLockOptimistic
'    NumMesiC = DatiGen.Fields("Mesi_consumo")
'    intYear = DatiGen.Fields("Anno_Calcolo")
'    intMonth = DatiGen.Fields("Mese_Calcolo")
'    NumMesiAbc = 12
'    IntStartDateAbc = DateAdd("m", -NumMesiAbc + 1, DateSerial(intYear, intMonth, 1))
'    IntStartDateAbc = Format(IntStartDateAbc, "dd/mmm/yyyy")
'    AnnoC = Year(IntStartDateAbc) 'anno inizio
'    MeseC = Month(IntStartDateAbc) 'mese inizio
'    DataInizio = DateSerial(AnnoC, MeseC - 1, 1)
'' Debug.Print DataInizio         ' **** DATA INIZIO
'NextKeyValue = ""
'TempKeyValue = ""
'j = 0
'VarWhereIn = Null
'
'Set rst = DBEngine(0)(0).OpenRecordset("SELECT cod_art, cod_art_correlato FROM tblArticoliStato " & _
'            "WHERE len(cod_art_correlato) > 1 " & _
'            "Order by cod_art_correlato;")
'Do While Not rst.EOF
'NextKeyValue = rst.Fields(1).Value
'    If TempKeyValue <> NextKeyValue Then ' Testa quando cambia codice
'        If j > 0 Then
'            ' Inserisco anagrafica fittizia
''            strSql = "INSERT INTO tblArticoli ([Cod_Art]," & _
''                     "[Des_art], [InsManualmente]) " & _
''                     "VALUES ('" & rst.Fields(1).Value & "_C'," & _
''                     "'" & rst.Fields(1).Value & "', " & _
''                     "'C'" & ")"
''            CurrentDb.Execute strSql
'            ' Aggiunge alla stringa il l'articolo correlato (loop precedente)
'            VarWhereIn = VarWhereIn & "'" & TempKeyValue & "'"
''           Debug.Print VarWhereIn
'            For intI = 1 To 12
'               strData = DateAdd("m", intI, DataInizio)
''              Debug.Print strData
'               Mese = Month(strData)
'               Anno = Year(strData)
'               ' INSERISCO NELLA TABELLA tblConsumiCorrelati i valori sommati
'               VarWhere = "[Cod_Art] IN (" & VarWhereIn & ") " & _
'                            "AND [Anno] = '" & Anno & "' AND [Mese] = " & "'" & Mese & "'"
''              Debug.Print VarWhere
'
'               ConsumoMese = DSum("[Consumo]", "tblConsumi", VarWhere)
'               NSpeditoMese = DSum("[N_Spedito_Mese]", "tblConsumi", VarWhere)
''              Debug.Print ConsumoMese
'               If Not IsNothing(ConsumoMese) Then
'               strSql = "INSERT INTO tblConsumiCorrelati" _
'                        & " ([Cod_Art], [Anno],[Mese], [Consumo], [N_Spedito_Mese]) " _
'                        & "VALUES ('" & TempKeyValue _
'                        & "', '" & Anno _
'                        & "', '" & Mese _
'                        & "', '" & ConsumoMese _
'                        & "', '" & NSpeditoMese & "');"
''               Debug.Print strSql
'                CurrentDb.Execute strSql
'               End If
'            Next intI
'            VarWhereIn = ""
'        End If
'    End If
'    TempKeyValue = rst.Fields(1).Value
'        VarWhereIn = VarWhereIn & " '" & rst.Fields(0).Value & "', "
''       Debug.Print VarWhereIn
'    j = j + 1
'    rst.MoveNext
'Loop
'' CARICA ULTIMO RECORD
'VarWhereIn = VarWhereIn & "'" & TempKeyValue & "'"
''           Debug.Print VarWhereIn
'            For intI = 1 To 12
'               strData = DateAdd("m", intI, DataInizio)
''              Debug.Print strData
'               Mese = Month(strData)
'               Anno = Year(strData)
'               VarWhere = "[Cod_Art] IN (" & VarWhereIn & ") " & _
'                            "AND [Anno] = '" & Anno & "' AND [Mese] = " & "'" & Mese & "'"
''              Debug.Print VarWhere
'
'               ConsumoMese = DSum("[Consumo]", "tblConsumi", VarWhere)
'               NSpeditoMese = DSum("[N_Spedito_Mese]", "tblConsumi", VarWhere)
''              Debug.Print ConsumoMese
'               If Not IsNothing(ConsumoMese) Then
'               strSql = "INSERT INTO tblConsumiCorrelati" _
'                        & " ([Cod_Art], [Anno],[Mese], [Consumo], [N_Spedito_Mese]) " _
'                        & "VALUES ('" & TempKeyValue _
'                        & "', '" & Anno _
'                        & "', '" & Mese _
'                        & "', '" & ConsumoMese _
'                        & "', '" & NSpeditoMese & "');"
' '             Debug.Print strSql
'               CurrentDb.Execute strSql
'               End If
'            Next intI
'' Visualizza puntatore mouse std
'Screen.MousePointer = vbDefault
'
''SELECT sum(tblConsumi.Consumo)
''FROM tblConsumi
''WHERE tblConsumi.Cod_Art IN ('0001301243B', '0001301019G')
''AND tblConsumi.Anno = '2012' and tblConsumi.Mese = '12';
'
'End Sub


Private Sub Form_Current()
    FillOptions
End Sub

Private Sub Form_Load()
' Setta parametro Lancio All a 0
Call scrivichiave("AggiornaAll", "NO")

'   Loop nella tbl Language per settare le etichette
    Dim dbs As Database, MySet As Recordset, strDefaultLanguage, strForm As String, hWndParent As Long, frmActive As Form

'   Look up la lingua di default
    strDefaultLanguage = DLookup("[Selected_Language]", "tblDatiGenerali")
    strForm = Me.name
    'strForm = Screen.ActiveForm.Name

'   Assegna Data Base corrente alla variabile dbs
    Set dbs = CurrentDb
    
    Set MySet = dbs.OpenRecordset("select * from tblControlNames where strForm = '" & strForm & "'")
    
'   Testa se la select restituisce 0 records
    If MySet.RecordCount > 0 Then
        MySet.MoveFirst
        Do While Not MySet.EOF
            If MySet.Fields("strLanguage").Value = strDefaultLanguage Then
                Me(MySet.Fields("strControlName")).Caption = MySet.Fields("strControlCaption").Value
                
            
            End If
            MySet.MoveNext
        Loop
        MySet.Close
    End If
    
FillOptions
    
    
'    ' 0 Importa Consumi
'    Me.box01.Visible = False
'    Me.box02.Visible = False
'    Me.box03.Visible = False
'
'    ' 1 Aggiorna Dati
'    Me.box11.Visible = False
'    Me.box12.Visible = False
'    Me.box13.Visible = False
'
'    ' 2 Calcolo Consumi
'    Me.box21.Visible = False
'    Me.box22.Visible = False
'    Me.box23.Visible = False
'
'     ' 3 Calcolo Rop Roq
'    Me.box31.Visible = False
'    Me.box32.Visible = False
'    Me.box33.Visible = False
'
'     ' 4 Aggiorna tutto
'    Me.box401.Visible = False
'    Me.box402.Visible = False
'    Me.box403.Visible = False
'
'    Me.box411.Visible = False
'    Me.box412.Visible = False
'    Me.box413.Visible = False
'
'    Me.box421.Visible = False
'    Me.box422.Visible = False
'    Me.box423.Visible = False
'
'    Me.box431.Visible = False
'    Me.box432.Visible = False
'    Me.box433.Visible = False
    
End Sub

Private Sub FillOptions()
' Fill in the options for this switchboard page.

' Numero dei bottoni della form.
    Const conNumButtons = 6

    Dim con As Object
    Dim rs As Object
    Dim stSql As String
    Dim intOption As Integer  ' 0x
    Dim strNum As String, ctl As control
    Dim intJ As Integer

    ' la label si chima boxXY
    ' Y = intJ
    ' X = intOption
    intJ = 0
    For intJ = 0 To conNumButtons
        For intOption = 1 To 3
            ' Usa Format per prendere 2 digits
            strNum = Format(intJ & intOption, "00")
            Me("box" & strNum).Visible = False
        Next intOption
    Next intJ

    ' Apre la tblFasiElaboraDati, and find
    ' the first item for this Switchboard Page.
    Set con = Application.CurrentProject.Connection
    stSql = "SELECT * FROM [tblFasiElaboraDati]"
    stSql = stSql & " WHERE [Eseguita] = True "
    stSql = stSql & " ORDER BY [Num_Fase];"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open stSql, con, 1   ' 1 = adOpenKeyset

    ' If there are no options for this Switchboard Page,
    ' display a message.  Otherwise, fill the page with the items.

    While (Not (rs.EOF))
        Me("box" & rs![NUM_FASE]).Visible = True
        '                att = rs![attivo]
        '                If att = True Then Me("OptionLabel" & rs![ItemNumber]).ForeColor = RGB(255, 0, 0)
        rs.MoveNext
    Wend


    ' Close the recordset and the database.
    rs.Close
    Set rs = Nothing
    Set con = Nothing

End Sub
