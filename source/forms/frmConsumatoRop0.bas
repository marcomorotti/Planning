Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureSizeMode =1
    DatasheetGridlinesBehavior =0
    GridY =10
    Width =16911
    DatasheetFontHeight =9
    ItemSuffix =22
    Left =-5385
    Top =1170
    Right =9360
    Bottom =6225
    Filter ="([tblArticoli Query].[Classe_Evento] Not In (\"Slow\",\"Very-Fast\",\"Very-Slow\""
        "))"
    RecSrcDt = Begin
        0x0c1477c76ee1e340
    End
    RecordSource ="qryConsumatoRop0"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    PictureSizeMode =4
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
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
            SizeMode =3
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
            FontSize =11
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Calibri"
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
        Begin TextBox
            FELineBreak = NotDefault
            Width =1701
            LabelX =-1701
            FontSize =11
            BorderColor =12632256
            FontName ="Calibri"
            AsianLineBreak =1
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
            ShowDatePicker =1
        End
        Begin FormHeader
            Height =1102
            BackColor =-2147483612
            Name ="IntestazioneMaschera"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Width =2715
                    Height =390
                    FontSize =12
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Auto_Title0"
                    Caption ="Consumi con Rop = 0"
                    FontName ="Segoe UI"
                    GridlineColor =-2147483616
                    LayoutCachedWidth =2715
                    LayoutCachedHeight =390
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =450
                    Width =1530
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta1"
                    Caption ="Codice"
                    FontName ="Segoe UI"
                    ColumnGroup =1
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =60
                    LayoutCachedTop =450
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =1064
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1665
                    Top =450
                    Width =3380
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta3"
                    Caption ="Descrizione"
                    FontName ="Segoe UI"
                    ColumnGroup =2
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =1665
                    LayoutCachedTop =450
                    LayoutCachedWidth =5045
                    LayoutCachedHeight =1064
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5120
                    Top =450
                    Width =851
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta5"
                    Caption ="Cons.\015\01212 mesi"
                    FontName ="Segoe UI"
                    ColumnGroup =3
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =5120
                    LayoutCachedTop =450
                    LayoutCachedWidth =5971
                    LayoutCachedHeight =1064
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6046
                    Top =450
                    Width =1305
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta7"
                    Caption ="N.Spedizioni 12 mesi"
                    FontName ="Segoe UI"
                    ColumnGroup =4
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =6046
                    LayoutCachedTop =450
                    LayoutCachedWidth =7351
                    LayoutCachedHeight =1064
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =10978
                    Top =450
                    Width =842
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta9"
                    Caption ="ROP \015\012Act"
                    FontName ="Segoe UI"
                    ColumnGroup =5
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =10978
                    LayoutCachedTop =450
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =1064
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =11895
                    Top =450
                    Width =851
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta11"
                    Caption ="Rop Prop"
                    FontName ="Segoe UI"
                    ColumnGroup =6
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =11895
                    LayoutCachedTop =450
                    LayoutCachedWidth =12746
                    LayoutCachedHeight =1064
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =12821
                    Top =450
                    Width =851
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta13"
                    Caption ="Lotto Act"
                    FontName ="Segoe UI"
                    ColumnGroup =7
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =12821
                    LayoutCachedTop =450
                    LayoutCachedWidth =13672
                    LayoutCachedHeight =1064
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =13747
                    Top =450
                    Width =846
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta15"
                    Caption ="Lotto Prop"
                    FontName ="Segoe UI"
                    ColumnGroup =8
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =13747
                    LayoutCachedTop =450
                    LayoutCachedWidth =14593
                    LayoutCachedHeight =1064
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =14668
                    Top =450
                    Width =2205
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta17"
                    Caption ="\015\012Classe_Evento"
                    FontName ="Segoe UI"
                    ColumnGroup =9
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =14668
                    LayoutCachedTop =450
                    LayoutCachedWidth =16873
                    LayoutCachedHeight =1064
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =16327
                    Width =576
                    Height =576
                    FontSize =8
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
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =16327
                    LayoutCachedWidth =16903
                    LayoutCachedHeight =576
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =7426
                    Top =450
                    Width =1701
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta20"
                    Caption ="Giacenza attuale:"
                    FontName ="Segoe UI"
                    ColumnGroup =10
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =7426
                    LayoutCachedTop =450
                    LayoutCachedWidth =9127
                    LayoutCachedHeight =1064
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =9202
                    Top =450
                    Width =1701
                    Height =614
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta21"
                    Caption ="Costo Csc:"
                    FontName ="Segoe UI"
                    ColumnGroup =11
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =9202
                    LayoutCachedTop =450
                    LayoutCachedWidth =10903
                    LayoutCachedHeight =1064
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =88
                    Left =15750
                    Width =577
                    Height =577
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
                    Name ="cmdExcel"
                    Caption ="E&xcel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Search ..."
                    UnicodeAccessKey =120
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x000000000000000000000000000000030000000b00000014000000190000001d ,
                        0x000000230000002a000000320000003800000040000000470000004c0000003b ,
                        0x0000002700000005000000000000001000000020000000400000005000000050 ,
                        0x4040407060686080a0a0a0a0b0b0b0b0c0c8c0c0d0e0e0e0e0f0e0f0d0d8d0c0 ,
                        0x0000004000000020106010c0105810d080a080f0c0d0c0f0d0e0d0ffffffffff ,
                        0xfffffffffffffffff0fffffff0fffffff0fffffff0fffffff0fff0ffc0d8c0f0 ,
                        0x0000004000000020106010f0a0b8a0ffc0d0c0ffffffffffffffffffffffffff ,
                        0xfffffffff0fffffff0ffffff90c0a0ff107030ff107030fff0f8f0ffe0e8e0ff ,
                        0x0000004000000020106810f0b0d0c0ff80a080ffffffffffc0e0c0ff80c080ff ,
                        0x70b870ff70b870ff60b070ff107030ff107030ff60a870ffe0f0e0fff0f8f0ff ,
                        0x0000004000000020106810f0c0d0c0ff407840ffffffffffffffffffc0e0c0ff ,
                        0x70b870ff70b870ff208040ff107030ff60a070ff60b070ffa0d0a0fff0ffffff ,
                        0x4040406000000020107010f0c0d0c0ff206020fff0f0f0ffffffffffffffffff ,
                        0xe0f0e0ff409860ff107030ff207840ffb0e0c0fffffffffff0fffffff0ffffff ,
                        0x7070707000000020207020f0c0d0c0ff206020ffb0c8b0ffffffffffffffffff ,
                        0xc0d8d0ff107030ff107030ff60a860ff70b870ffa0d0a0fff0fff0fff0ffffff ,
                        0x9098909000000020207820f0c0d0c0ff206820ff90a890fffffffffff0f8f0ff ,
                        0x107030ff107030ff90c0a0ffe0f0e0ff70b870ff70b870ff90c890fff0f8f0ff ,
                        0xb0b0b0a000000020207820f0d0e0d0ff70a870ff90b090ffffffffff409060ff ,
                        0x308050ff60a070ffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xd0d0d0a000000020208020f0d0e8d0ff80b880ff80b080ffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xf0f0f0b000000020208020e0d0e8d0ffa0c8a0ff80b880ffd0e8d0ffd0e0d0ff ,
                        0xd0e0d0ffd0e0d0ffc0d8c0ffb0d8b0ffa0d0a0ff90c8a0ffd0e8d0ff50a050e0 ,
                        0xffffff2000000000208020a0a0c890ffe0f0e0ff90c890ff90c890ff90c890ff ,
                        0x90c890ff90c890ff90c890ff90c890ff90c890ff90c890fff0f8f0ff208020d0 ,
                        0x000000000000000030882030308820ffc0e0c0fff0f8f0ffb0d8b0ffa0d0a0ff ,
                        0xa0d0a0ffa0d0a0ffa0d0a0ffa0d0a0ffa0d0a0ff90d0a0fff0f8f0ff308820d0 ,
                        0x00000000000000000000000030882060308820f090c890ffd0e8d0fff0f8f0ff ,
                        0xf0f8f0fff0f8f0fff0f8f0fff0f8f0fff0f8f0fff0f8f0fff0f8f0ff308820d0 ,
                        0x0000000000000000000000000000000030902020309020a0309020e0309020ff ,
                        0x309020ff309020ff309020ff309020ff309020ff309020ff309020ff309020d0 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =15750
                    LayoutCachedWidth =16327
                    LayoutCachedHeight =577
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =748
            Name ="Corpo"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    IMEHold = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =30
                    Width =1530
                    Height =680
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Cod_art"
                    ControlSource ="Cod_art"
                    StatusBarText ="Codice articolo"
                    ColumnGroup =1
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =60
                    LayoutCachedTop =30
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =710
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1665
                    Top =30
                    Width =3380
                    Height =680
                    ColumnWidth =5550
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Des_art"
                    ControlSource ="Des_art"
                    StatusBarText ="Descrizione articolo"
                    ColumnGroup =2
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =1665
                    LayoutCachedTop =30
                    LayoutCachedWidth =5045
                    LayoutCachedHeight =710
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5120
                    Top =30
                    Width =851
                    Height =680
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="SConsumo"
                    ControlSource ="SConsumo_12"
                    Format ="Standard"
                    StatusBarText ="Somma Consumi nel periodo di calcolo"
                    ColumnGroup =3
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =5120
                    LayoutCachedTop =30
                    LayoutCachedWidth =5971
                    LayoutCachedHeight =710
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6046
                    Top =30
                    Width =1305
                    Height =680
                    TabIndex =3
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="SSpedito"
                    ControlSource ="SSpedito_12"
                    Format ="Standard"
                    StatusBarText ="Somma Spedizioni nel periodo di calcolo"
                    ColumnGroup =4
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =6046
                    LayoutCachedTop =30
                    LayoutCachedWidth =7351
                    LayoutCachedHeight =710
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10978
                    Top =30
                    Width =842
                    Height =680
                    TabIndex =8
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =255
                    Name ="ROP"
                    ControlSource ="ROP"
                    StatusBarText ="Dato import"
                    ColumnGroup =5
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =10978
                    LayoutCachedTop =30
                    LayoutCachedWidth =11820
                    LayoutCachedHeight =710
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11895
                    Top =30
                    Width =851
                    Height =680
                    TabIndex =6
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Punto_riordino"
                    ControlSource ="Punto_riordino"
                    Format ="Standard"
                    StatusBarText ="Punto di riordino (dato calcolato)"
                    ColumnGroup =6
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =11895
                    LayoutCachedTop =30
                    LayoutCachedWidth =12746
                    LayoutCachedHeight =710
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =12821
                    Top =30
                    Width =851
                    Height =680
                    TabIndex =7
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =255
                    Name ="ROQ"
                    ControlSource ="ROQ"
                    StatusBarText ="Dato import"
                    ColumnGroup =7
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =12821
                    LayoutCachedTop =30
                    LayoutCachedWidth =13672
                    LayoutCachedHeight =710
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =13747
                    Top =30
                    Width =846
                    Height =680
                    TabIndex =9
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Lotto_ec_acq"
                    ControlSource ="Lotto_ec_acq"
                    Format ="Standard"
                    StatusBarText ="Lotto economico di riacquisto (dato calcolato)"
                    ColumnGroup =8
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =13747
                    LayoutCachedTop =30
                    LayoutCachedWidth =14593
                    LayoutCachedHeight =710
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =14668
                    Top =30
                    Width =2205
                    Height =680
                    ColumnWidth =1755
                    TabIndex =10
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Classe_Evento"
                    ControlSource ="Classe_Evento"
                    ColumnGroup =9
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =14668
                    LayoutCachedTop =30
                    LayoutCachedWidth =16873
                    LayoutCachedHeight =710
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7426
                    Top =30
                    Height =680
                    TabIndex =4
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Giac_Media"
                    ControlSource ="Giac_Media"
                    Format ="Standard"
                    StatusBarText ="Giacenza media (dato calcolato)"
                    ColumnGroup =10
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =7426
                    LayoutCachedTop =30
                    LayoutCachedWidth =9127
                    LayoutCachedHeight =710
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9202
                    Top =30
                    Height =680
                    TabIndex =5
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Cs_Csc"
                    ControlSource ="Cs_Csc"
                    Format ="Standard"
                    StatusBarText ="Costo unitario Standard"
                    ColumnGroup =11
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483609

                    LayoutCachedLeft =9202
                    LayoutCachedTop =30
                    LayoutCachedWidth =10903
                    LayoutCachedHeight =710
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="PièDiPaginaMaschera"
            AutoHeight =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdClose_Click()
  DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdExcel_Click()
' Esporta Articoli con Rop = 0
'/// Inizia a creare File Csv

' // Per usare questa funzione aggiungere librerie Excel e Word

' Dim conCurrent As ADODB.Connection
Dim Db As DAO.Database
Dim brs As DAO.Recordset
Dim rstOutput As New ADODB.Recordset
Dim objField As ADODB.Field
Dim intFile As Integer
Dim strSQL As String, strDataLine As String
Dim filenm As String
Dim i As Integer


' Costruzione File da aprire
intFile = FreeFile
i = 1

filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") & _
    "ConsumatoRop0_" & i & ".tsv"

' Cerca File da aprire
    If FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & Format(Date, _
        "yyyymmdd") & "ConsumatoRop0_" & i & ".tsv") = False Then
    Else
      Do While FileExists(TrailingSlash(CurrentProject.Path) & "\Export\" & _
          Format(Date, "yyyymmdd") & "ConsumatoRop0_" & i & ".tsv") = True
       i = i + 1
      ' MsgBox "Il ciclo è stato eseguito " & i & " volte."
       Loop
    filenm = CurrentProject.Path & "\Export\" & Format(Date, "yyyymmdd") & _
        "ConsumatoRop0_" & i & ".tsv"
    End If
' Chiede se si vuole esportare il file
If vbYes = MsgBox("Vuoi esportare i dati in  " & filenm, vbQuestion + vbYesNo + _
    vbDefaultButton2, gstrAppTitle) Then
    ' Visualizza puntatore mouse a clessidra
    Screen.MousePointer = vbHourGlass

    Open filenm For Output As #intFile
    ' Inserisce testata
    strDataLine = "Cod_art" & Chr(9) & "Des_art" & Chr(9) & "Qtà_Consumata" & Chr(9) & _
        "N_Spedizioni" & Chr(9) & "Giacenza" & Chr(9) & "Cs_Csc" & Chr(9) & _
        "ROP_Act" & Chr(9) & "ROP_Prop" & Chr(9) & "ROQ_Act" & Chr(9) & "ROQ_Prop" & Chr(9) & _
        "Classe_Evento"

    Print #intFile, strDataLine
    'Ora loop per il recordset e scrive un TSV file per ogni record
    Set Db = CurrentDb
       Set brs = DBEngine(0)(0).OpenRecordset("SELECT * FROM qryConsumatoRop0")
   
    If Not brs.EOF And Not brs.BOF Then
    brs.MoveFirst
    End If
    While Not brs.EOF
    strDataLine = brs.Fields("Cod_art").Value & Chr(9) & _
        brs.Fields("Des_art").Value & Chr(9) & _
        brs.Fields("SConsumo_12").Value & Chr(9) & brs.Fields("SSpedito_12").Value & Chr(9) & _
        brs.Fields("Giac_Media").Value & Chr(9) & brs.Fields("Cs_Csc").Value & Chr(9) & _
        brs.Fields("ROP").Value & Chr(9) & brs.Fields("Punto_riordino").Value & Chr(9) & _
        brs.Fields("ROQ").Value & Chr(9) & brs.Fields("Lotto_ec_acq").Value & Chr(9) & _
         brs.Fields("Classe_Evento").Value
   
    Print #intFile, strDataLine
    brs.MoveNext
    strDataLine = ""
    Wend
    MsgBox ("I Dati sono stati salvati come TSV file" & CurrentProject.Path & _
        "\Export\" & Format(Date, "yyyymmdd") & "ConsumatoRop0_" & i & ".xls")
    brs.Close
    Set brs = Nothing
    Set Db = Nothing
    Close #intFile
Else
Exit Sub
' Visualizza puntatore mouse std
Screen.MousePointer = vbDefault

End If

End Sub
