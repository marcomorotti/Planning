Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularCharSet =161
    BorderStyle =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =3288
    DatasheetFontHeight =11
    ItemSuffix =10
    Left =570
    Top =1290
    Right =3855
    Bottom =4680
    Filter ="Indice='Giacenza_Valo'"
    RecSrcDt = Begin
        0xb7c8da6d1447e440
    End
    RecordSource ="qryDashBoard"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =255
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
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
        Begin BoundObjectFrame
            SizeMode =3
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
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
        Begin Section
            Height =3118
            Name ="Corpo"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1395
                    Top =2152
                    Width =1417
                    Height =315
                    TabIndex =1
                    Name ="Testo3"
                    ControlSource ="Actual"
                    Format ="€ #,##0.00;-€ #,##0.00"

                    LayoutCachedLeft =1395
                    LayoutCachedTop =2152
                    LayoutCachedWidth =2812
                    LayoutCachedHeight =2467
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =390
                            Top =2152
                            Width =1005
                            Height =315
                            Name ="Etichetta4"
                            Caption ="Attuale:"
                            LayoutCachedLeft =390
                            LayoutCachedTop =2152
                            LayoutCachedWidth =1395
                            LayoutCachedHeight =2467
                        End
                    End
                End
                Begin BoundObjectFrame
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =93
                    Left =690
                    Top =60
                    Width =1875
                    Height =1635
                    BorderColor =12835293
                    Name ="ProgressBarHiGood"
                    ControlSource ="GaugesHiGood"
                    GridlineStyleTop =4
                    GridlineColor =10921638
                    HorizontalAnchor =2

                    LayoutCachedLeft =690
                    LayoutCachedTop =60
                    LayoutCachedWidth =2565
                    LayoutCachedHeight =1695
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1395
                    Top =2587
                    Width =1417
                    Height =315
                    TabIndex =2
                    Name ="Testo5"
                    ControlSource ="Target"
                    Format ="€ #,##0.00;-€ #,##0.00"

                    LayoutCachedLeft =1395
                    LayoutCachedTop =2587
                    LayoutCachedWidth =2812
                    LayoutCachedHeight =2902
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextFontCharSet =161
                            Left =390
                            Top =2587
                            Width =1005
                            Height =315
                            Name ="Etichetta6"
                            Caption ="Target:"
                            LayoutCachedLeft =390
                            LayoutCachedTop =2587
                            LayoutCachedWidth =1395
                            LayoutCachedHeight =2902
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =119
                    TextFontCharSet =161
                    IMESentenceMode =3
                    Left =2550
                    Top =60
                    Width =727
                    Height =315
                    FontWeight =700
                    TabIndex =3
                    ForeColor =255
                    Name ="Testo7"
                    ControlSource ="PcntIndice"
                    Format ="Fixed"

                    LayoutCachedLeft =2550
                    LayoutCachedTop =60
                    LayoutCachedWidth =3277
                    LayoutCachedHeight =375
                End
                Begin Label
                    OverlapFlags =85
                    TextFontCharSet =161
                    TextAlign =2
                    Left =135
                    Top =1695
                    Width =2685
                    Height =405
                    FontSize =14
                    FontWeight =700
                    Name ="Etichetta9"
                    Caption ="VAL. CODICI GIACENTI"
                    LayoutCachedLeft =135
                    LayoutCachedTop =1695
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =2100
                End
            End
        End
    End
End
