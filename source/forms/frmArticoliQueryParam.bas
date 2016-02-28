Version =20
VersionRequired =20
Begin Form
    CloseButton = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridX =24
    GridY =24
    Width =8676
    ItemSuffix =44
    Left =1185
    Top =1965
    Right =10140
    Bottom =8505
    Filter ="[Classe_Evento] = 'Very-Fast' AND [AbcConsumoValoreLs] = 'A1' AND [Giac_Media] <"
        "> 0"
    RecSrcDt = Begin
        0x87f757c13202e440
    End
    RecordSource ="qryArticoliParam"
    Caption ="Articoli Matrice ClasseEvento-ValoreConsumo"
    PrtMip = Begin
        0x550300006e040000550300006e04000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    Begin
        Begin Label
            FontWeight =700
            BackColor =12632256
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
            SpecialEffect =2
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
            TextFontFamily =2
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
        Begin OptionButton
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =187
            Height =187
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
        Begin CheckBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Width =187
            Height =187
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
            AutoLabel = NotDefault
            SpecialEffect =2
            Height =255
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
        Begin ListBox
            AutoLabel = NotDefault
            SpecialEffect =2
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
        Begin ComboBox
            AutoLabel = NotDefault
            SpecialEffect =2
            Height =255
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
        Begin FormHeader
            Height =690
            BackColor =-2147483633
            Name ="FormHeader1"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =60
                    Top =435
                    Width =1005
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text14"
                    Caption ="Cod articolo"
                    LayoutCachedLeft =60
                    LayoutCachedTop =435
                    LayoutCachedWidth =1065
                    LayoutCachedHeight =675
                End
                Begin Label
                    OverlapFlags =85
                    Left =1605
                    Top =435
                    Width =900
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="Text18"
                    Caption ="Descrizione"
                    LayoutCachedLeft =1605
                    LayoutCachedTop =435
                    LayoutCachedWidth =2505
                    LayoutCachedHeight =675
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =6855
                    Top =435
                    Width =1080
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label44"
                    Caption ="Classe evento"
                    LayoutCachedLeft =6855
                    LayoutCachedTop =435
                    LayoutCachedWidth =7935
                    LayoutCachedHeight =675
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    Left =5115
                    Top =435
                    Width =825
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    Name ="lblCost"
                    Caption ="Giac"
                    LayoutCachedLeft =5115
                    LayoutCachedTop =435
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =675
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =6075
                    Top =435
                    Width =495
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Etichetta34"
                    Caption ="Csc"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    TextAlign =2
                    Left =7950
                    Top =435
                    Width =705
                    Height =240
                    FontWeight =400
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Etichetta41"
                    Caption ="ClasseC"
                    LayoutCachedLeft =7950
                    LayoutCachedTop =435
                    LayoutCachedWidth =8655
                    LayoutCachedHeight =675
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    TextFontFamily =34
                    Left =8295
                    Width =381
                    Height =456
                    FontSize =8
                    FontWeight =400
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
                    ControlTipText ="Close the form"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8295
                    LayoutCachedWidth =8676
                    LayoutCachedHeight =456
                End
            End
        End
        Begin Section
            Height =257
            BackColor =8421504
            Name ="txtFuelLog"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7956
                    Width =681
                    Height =256
                    TabIndex =5
                    Name ="txtflHrs"
                    ControlSource ="AbcConsumoValoreLs"
                    Format ="Fixed"

                    LayoutCachedLeft =7956
                    LayoutCachedWidth =8637
                    LayoutCachedHeight =256
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6735
                    Width =1221
                    Height =256
                    TabIndex =4
                    Name ="txtflKms"
                    ControlSource ="Classe_Evento"
                    Format ="Standard"

                    LayoutCachedLeft =6735
                    LayoutCachedWidth =7956
                    LayoutCachedHeight =256
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =5190
                    Width =735
                    Height =256
                    TabIndex =2
                    Name ="txtflliters"
                    ControlSource ="giac_media"
                    Format ="Standard"

                    LayoutCachedLeft =5190
                    LayoutCachedWidth =5925
                    LayoutCachedHeight =256
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =5925
                    Width =810
                    Height =256
                    TabIndex =3
                    Name ="FlCost"
                    ControlSource ="Cs_Csc"
                    Format ="#,###;-#,##0"

                    LayoutCachedLeft =5925
                    LayoutCachedWidth =6735
                    LayoutCachedHeight =256
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    TextAlign =1
                    Height =256
                    Name ="txtCod_art"
                    ControlSource ="Cod_art"
                    Format ="Short Date"

                    LayoutCachedWidth =1440
                    LayoutCachedHeight =256
                End
                Begin TextBox
                    Locked = NotDefault
                    SpecialEffect =0
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1440
                    Width =3750
                    Height =256
                    TabIndex =1
                    Name ="txtRifornimento"
                    ControlSource ="Des_art"

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =5190
                    LayoutCachedHeight =256
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="FormFooter2"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.name
    ' DoCmd.Maximize
End Sub
