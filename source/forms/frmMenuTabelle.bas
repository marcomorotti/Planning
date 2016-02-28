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
    Width =5346
    DatasheetFontHeight =10
    ItemSuffix =9
    Left =2130
    Top =3660
    Right =7710
    Bottom =8145
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x3955d11af5b9e340
    End
    DatasheetFontName ="Arial"
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
            Height =4260
            BackColor =12311007
            Name ="Corpo"
            Begin
                Begin Label
                    SpecialEffect =4
                    BackStyle =1
                    OldBorderStyle =1
                    BorderWidth =2
                    OverlapFlags =85
                    TextAlign =2
                    Left =-45
                    Width =4785
                    Height =555
                    FontSize =20
                    FontWeight =700
                    BackColor =12311007
                    ForeColor =13209
                    Name ="Etichetta7"
                    Caption ="TABELLE"
                    FontName ="Verdana"
                    LayoutCachedLeft =-45
                    LayoutCachedWidth =4740
                    LayoutCachedHeight =555
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1700
                    Top =1474
                    Width =2268
                    Height =435
                    FontSize =10
                    FontWeight =700
                    ForeColor =13209
                    Name ="cmdImportDataOracle"
                    Caption ="Parametri Eventi"
                    OnClick ="=ApriMaschere(\"frmNewCategoriaEventi\")"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =4770
                    Width =576
                    Height =576
                    TabIndex =1
                    ForeColor =0
                    Name ="cmdClose"
                    Caption ="(demo)"
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
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
