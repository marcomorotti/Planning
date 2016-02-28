Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    TabularCharSet =238
    PictureSizeMode =1
    DatasheetGridlinesBehavior =0
    GridY =10
    Width =19877
    DatasheetFontHeight =9
    ItemSuffix =18
    Left =345
    Top =2445
    Right =15105
    Bottom =7320
    Filter ="([tblArticoli Query].[Classe_Evento] Not In (\"Slow\",\"Very-Fast\",\"Very-Slow\""
        "))"
    RecSrcDt = Begin
        0x6fbfdf516fe1e340
    End
    RecordSource ="qryConsumatoNoRopGt0"
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
                    Left =30
                    Top =30
                    Width =4245
                    Height =330
                    FontSize =12
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Auto_Title0"
                    Caption ="Articoli NON Consumati con Rop > 0"
                    FontName ="Segoe UI"
                    GridlineColor =-2147483616
                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =4275
                    LayoutCachedHeight =360
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =450
                    Width =1530
                    Height =585
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
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1665
                    Top =450
                    Width =3380
                    Height =585
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
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5120
                    Top =450
                    Width =851
                    Height =585
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
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6046
                    Top =450
                    Width =1305
                    Height =585
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta7"
                    Caption ="N. Spedizioni"
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
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =7426
                    Top =450
                    Width =842
                    Height =585
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta9"
                    Caption ="ROP Act"
                    FontName ="Segoe UI"
                    ColumnGroup =5
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =7426
                    LayoutCachedTop =450
                    LayoutCachedWidth =8268
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =8343
                    Top =450
                    Width =851
                    Height =585
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
                    LayoutCachedLeft =8343
                    LayoutCachedTop =450
                    LayoutCachedWidth =9194
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =9269
                    Top =450
                    Width =851
                    Height =585
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
                    LayoutCachedLeft =9269
                    LayoutCachedTop =450
                    LayoutCachedWidth =10120
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =10195
                    Top =450
                    Width =846
                    Height =585
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
                    LayoutCachedLeft =10195
                    LayoutCachedTop =450
                    LayoutCachedWidth =11041
                    LayoutCachedHeight =1035
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =11116
                    Top =450
                    Width =2205
                    Height =585
                    FontSize =10
                    FontWeight =700
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    ForeColor =-2147483615
                    Name ="Etichetta17"
                    Caption ="Classe_Evento"
                    FontName ="Segoe UI"
                    ColumnGroup =9
                    GroupTable =1
                    GridlineStyleLeft =1
                    GridlineStyleTop =1
                    GridlineStyleRight =1
                    GridlineStyleBottom =1
                    GridlineColor =-2147483616
                    LayoutCachedLeft =11116
                    LayoutCachedTop =450
                    LayoutCachedWidth =13321
                    LayoutCachedHeight =1035
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =13425
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

                    LayoutCachedLeft =13425
                    LayoutCachedWidth =14001
                    LayoutCachedHeight =576
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =380
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
                    Top =56
                    Width =1530
                    Height =286
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
                    LayoutCachedTop =56
                    LayoutCachedWidth =1590
                    LayoutCachedHeight =342
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1665
                    Top =56
                    Width =3380
                    Height =286
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
                    LayoutCachedTop =56
                    LayoutCachedWidth =5045
                    LayoutCachedHeight =342
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5120
                    Top =56
                    Width =851
                    Height =286
                    TabIndex =2
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="SConsumo"
                    ControlSource ="SConsumo"
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
                    LayoutCachedTop =56
                    LayoutCachedWidth =5971
                    LayoutCachedHeight =342
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6046
                    Top =56
                    Width =1305
                    Height =286
                    TabIndex =3
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="SSpedito"
                    ControlSource ="SSpedito"
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
                    LayoutCachedTop =56
                    LayoutCachedWidth =7351
                    LayoutCachedHeight =342
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7426
                    Top =56
                    Width =842
                    Height =286
                    TabIndex =4
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

                    LayoutCachedLeft =7426
                    LayoutCachedTop =56
                    LayoutCachedWidth =8268
                    LayoutCachedHeight =342
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8343
                    Top =56
                    Width =851
                    Height =286
                    TabIndex =5
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

                    LayoutCachedLeft =8343
                    LayoutCachedTop =56
                    LayoutCachedWidth =9194
                    LayoutCachedHeight =342
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9269
                    Top =56
                    Width =851
                    Height =286
                    TabIndex =6
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

                    LayoutCachedLeft =9269
                    LayoutCachedTop =56
                    LayoutCachedWidth =10120
                    LayoutCachedHeight =342
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10195
                    Top =56
                    Width =846
                    Height =286
                    TabIndex =7
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

                    LayoutCachedLeft =10195
                    LayoutCachedTop =56
                    LayoutCachedWidth =11041
                    LayoutCachedHeight =342
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =11116
                    Top =56
                    Width =2205
                    Height =286
                    ColumnWidth =1755
                    TabIndex =8
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

                    LayoutCachedLeft =11116
                    LayoutCachedTop =56
                    LayoutCachedWidth =13321
                    LayoutCachedHeight =342
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
