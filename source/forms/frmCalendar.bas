Version =20
VersionRequired =20
Begin Form
    AllowEditing = NotDefault
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    CloseButton = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =0
    BorderStyle =3
    Width =4320
    ItemSuffix =68
    Left =2265
    Top =1725
    Right =8370
    Bottom =9255
    HelpContextId =500
    RecSrcDt = Begin
        0xed312f5baf8be140
    End
    Caption ="Calendar"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    Begin
        Begin Label
            TextAlign =3
            FontWeight =700
            BackColor =12632256
            FontName ="Arial"
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
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            Width =187
            Height =187
            LabelX =-146
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
            AddColon = NotDefault
            SpecialEffect =2
            LabelAlign =3
            Width =187
            Height =187
            LabelX =-146
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
        Begin OptionGroup
            SpecialEffect =3
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
            SpecialEffect =2
            LabelAlign =3
            TextFontFamily =0
            LabelX =-146
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
            SpecialEffect =2
            LabelAlign =3
            TextFontFamily =0
            LabelX =-146
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
            SpecialEffect =2
            LabelAlign =3
            TextFontFamily =0
            LabelX =-146
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
        Begin CustomControl
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
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
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
            Height =374
            BackColor =12632256
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    TextFontFamily =34
                    Left =2880
                    Top =14
                    FontSize =9
                    Name ="cmdSave"
                    Caption ="&Save"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Save changes."

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    TextFontFamily =34
                    Left =14
                    Top =14
                    FontSize =9
                    TabIndex =1
                    Name ="cmdCancel"
                    Caption ="Ca&ncel"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Cancel changes and close the window."

                End
            End
        End
        Begin Section
            Height =4590
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =150
                    Top =3330
                    TabIndex =7
                    Name ="ctlCalendar"
                    ShortcutMenuBar ="Form Control Shortcut Bar"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Left =1920
                    Top =3585
                    Width =120
                    Height =225
                    Name ="lblColon"
                    Caption =":"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =119
                    Left =1275
                    Top =3585
                    Width =645
                    Name ="txtHour"
                    Format ="00"
                    DefaultValue ="12"
                    InputMask ="09"
                    OnKeyPress ="[Event Procedure]"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Press + or - keys to scroll value."

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =255
                            TextAlign =1
                            Left =1279
                            Top =3330
                            Width =570
                            Height =255
                            Name ="lblHour"
                            Caption ="Hour:"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =87
                    Left =2055
                    Top =3585
                    Width =645
                    TabIndex =1
                    Name ="txtMinute"
                    Format ="00"
                    DefaultValue ="0"
                    InputMask ="09"
                    OnKeyPress ="[Event Procedure]"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Press + or - keys to scroll value."

                    Begin
                        Begin Label
                            BackStyle =0
                            OverlapFlags =93
                            TextAlign =1
                            Left =2059
                            Top =3330
                            Width =465
                            Height =255
                            Name ="lblMinute"
                            Caption ="Min:"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =600
                    Top =90
                    Width =321
                    Height =321
                    TabIndex =2
                    Name ="cmdPrevious"
                    Caption ="<-"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadadadad1dadadaadadadad11adadaddadadad111dadada ,
                        0xadadad1111adadaddadad11111dadadaadadad1111adadaddadadad111dadada ,
                        0xadadadad11adadaddadadadad1dadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ShortcutMenuBar ="Form Control Shortcut Bar"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =3405
                    Top =90
                    Width =321
                    Height =321
                    TabIndex =3
                    Name ="cmdNext"
                    Caption ="->"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddadada1adadadadaadadad11adadadaddadada111adadada ,
                        0xadadad1111adadaddadada11111adadaadadad1111adadaddadada111adadada ,
                        0xadadad11adadadaddadada1adadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ShortcutMenuBar ="Form Control Shortcut Bar"

                End
                Begin ComboBox
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    ColumnCount =2
                    Left =975
                    Top =120
                    Width =1290
                    TabIndex =4
                    Name ="cmbMonth"
                    RowSourceType ="Value List"
                    RowSource ="1;\"January\";2;\"February\";3;\"March\";4;\"April\";5;\"May\";6;\"June\";7;\"Ju"
                        "ly\";8;\"August\";9;\"September\";10;\"October\";11;\"November\";12;\"December\""
                    ColumnWidths ="0;1440"
                    AfterUpdate ="[Event Procedure]"
                    ShortcutMenuBar ="Form Control Shortcut Bar"

                End
                Begin ComboBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =2430
                    Top =120
                    Width =915
                    TabIndex =5
                    ColumnInfo ="\"\";\"\";\"3\";\"2\""
                    Name ="cmbYear"
                    RowSourceType ="Table/Query"
                    RowSource ="ztblYears"
                    ColumnWidths ="720"
                    AfterUpdate ="[Event Procedure]"
                    ShortcutMenuBar ="Form Control Shortcut Bar"

                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =150
                    Top =3855
                    Width =4035
                    Height =675
                    Name ="lblTimeInstruct"
                    Caption ="Press Tab to move from Calendar to Hour / Min boxes.  Type in Hour (24 hour cloc"
                        "k) and Minute values or use + and - keys to change the values."
                End
                Begin Label
                    SpecialEffect =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =150
                    Top =585
                    Width =554
                    Height =270
                    BackColor =-2147483633
                    Name ="Label14"
                    Caption ="Sun"
                End
                Begin Label
                    SpecialEffect =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =726
                    Top =585
                    Width =554
                    Height =270
                    BackColor =-2147483633
                    Name ="Label15"
                    Caption ="Mon"
                End
                Begin Label
                    SpecialEffect =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =1302
                    Top =585
                    Width =554
                    Height =270
                    BackColor =-2147483633
                    Name ="Label16"
                    Caption ="Tue"
                End
                Begin Label
                    SpecialEffect =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =1878
                    Top =585
                    Width =554
                    Height =270
                    BackColor =-2147483633
                    Name ="Label17"
                    Caption ="Wed"
                End
                Begin Label
                    SpecialEffect =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =2454
                    Top =585
                    Width =554
                    Height =270
                    BackColor =-2147483633
                    Name ="Label18"
                    Caption ="Thu"
                End
                Begin Label
                    SpecialEffect =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =3030
                    Top =585
                    Width =554
                    Height =270
                    BackColor =-2147483633
                    Name ="Label19"
                    Caption ="Fri"
                End
                Begin Label
                    SpecialEffect =1
                    OverlapFlags =87
                    TextAlign =2
                    Left =3606
                    Top =585
                    Width =554
                    Height =270
                    BackColor =-2147483633
                    Name ="Label20"
                    Caption ="Sat"
                End
                Begin OptionGroup
                    OverlapFlags =85
                    Left =137
                    Top =885
                    Width =4032
                    Height =2394
                    TabIndex =6
                    Name ="optCalendar"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="0"
                    ShortcutMenuBar ="Form Control Shortcut Bar"

                    Begin
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =137
                            Top =885
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl01"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =713
                            Top =885
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl02"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1289
                            Top =885
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl03"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1865
                            Top =885
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl04"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =2441
                            Top =885
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl05"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3017
                            Top =885
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl06"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3593
                            Top =885
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl07"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =137
                            Top =1274
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl08"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =713
                            Top =1274
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl09"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1289
                            Top =1274
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl10"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1865
                            Top =1274
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl11"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =2441
                            Top =1274
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl12"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3017
                            Top =1274
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl13"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3593
                            Top =1274
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl14"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =137
                            Top =1663
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl15"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =713
                            Top =1663
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl16"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1289
                            Top =1663
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl17"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1865
                            Top =1663
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl18"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =2441
                            Top =1663
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl19"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3017
                            Top =1663
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl20"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3593
                            Top =1663
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl21"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =137
                            Top =2052
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl22"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =713
                            Top =2052
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl23"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1289
                            Top =2052
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl24"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1865
                            Top =2052
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl25"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =2441
                            Top =2052
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl26"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3017
                            Top =2052
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl27"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3593
                            Top =2052
                            Width =576
                            Height =389
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl28"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =137
                            Top =2441
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl29"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =713
                            Top =2441
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl30"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1289
                            Top =2441
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl31"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1865
                            Top =2441
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl32"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =2441
                            Top =2441
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl33"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3017
                            Top =2441
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl34"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3593
                            Top =2441
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl35"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =137
                            Top =2860
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl36"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =713
                            Top =2860
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl37"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1289
                            Top =2860
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl38"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =1865
                            Top =2860
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl39"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =2441
                            Top =2860
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl40"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3017
                            Top =2860
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl41"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                        Begin ToggleButton
                            OverlapFlags =87
                            Left =3593
                            Top =2860
                            Width =576
                            Height =419
                            FontWeight =700
                            OptionValue =0
                            ForeColor =0
                            Name ="tgl42"
                            ShortcutMenuBar ="Form Control Shortcut Bar"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =12632256
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' Non-OCX "calendar" control to pick dates and times
' Form is opened by Public GetDate module as a dialog.
' If the user clicks Cancel, form closes, and GetDate does not
' update the control because this form isn't here anymore.
' When user clicks Save, form hides itself so that GetDate
' can "pluck" the date from the hidden ctlCalendar text box and
' add the time from the txtHour and txtMinute text boxes.
' Copyright 2002, Viescas Consulting, Inc.
Dim intFirst As Integer, intLast As Integer, intLastDay As Integer
Dim intMonth As Integer, intYear As Integer, intDay As Integer
Dim varDate As Variant

Private Sub cmbMonth_AfterUpdate()

    ' Call the common routine to re-draw the calendar
    '   for the newly selected month
    SetDays
    ' Get the previous year
    intYear = Year(varDate)
    ' .. and day value
    intDay = Day(varDate)
    ' Adjust if day was > 28 (not all months have 29, 30, 31 days)
    If intDay > 28 Then
        ' Make adjustment based on new month value
        Select Case Me.cmbMonth
            ' Set February to 28 (doesn't account for leap years)
            Case 2
                intDay = 28
            ' Set April, June, September, November to 30 if it was 31
            Case 4, 6, 9, 11
                If intDay = 31 Then intDay = 30
        End Select
    End If
    ' Calculate the new date based on the new month
    varDate = DateSerial(intYear, Me.cmbMonth, intDay)
    ' Move the highlighted day, if necessary
    Me.optCalendar.Value = intDay
    ' Save the new date in hidden form control for later
    Me.ctlCalendar = varDate
    
End Sub

Private Sub cmbYear_AfterUpdate()
Dim intMonth As Integer, intDay As Integer

    ' Call the common routine to re-draw the calendar
    '   for the newly selected year
    SetDays
    ' Get the previous month
    intMonth = Month(varDate)
    ' .. and day
    intDay = Day(varDate)
    ' Make adjustment for February - might be moving off a leap year
    If intMonth = 2 And intDay = 29 Then intDay = 28
    ' Calculate the new date based on the new year
    varDate = DateSerial(Me.cmbYear, intMonth, intDay)
    ' Move to the highlighted day
    Me.optCalendar.Value = Day(varDate)
    ' Save the new date in hidden form control for later
    Me.ctlCalendar = varDate
    
End Sub

Public Sub cmdCancel_Click()
    ' Closing doesn't pass the value back
    DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdNext_Click()
    ' Move to the next month -
    ' This also fixes the day number if the previous day
    '   isn't available in the new month.
    varDate = DateAdd("m", 1, varDate)
    ' Change the month value
    Me.cmbMonth = Month(varDate)
    ' Change the year value
    Me.cmbYear = Year(varDate)
    ' Re-draw the calendar
    SetDays
    ' Make sure the correct box is highlighted
    Me.optCalendar.Value = Day(varDate)
    ' Save the new date in hidden form control for later
    Me.ctlCalendar = varDate
    
End Sub

Private Sub cmdPrevious_Click()
    ' Move to the previous month
    ' This also fixes the day number if the previous day
    '   isn't available in the new month.
    varDate = DateAdd("m", -1, varDate)
    ' Change the month value
    Me.cmbMonth = Month(varDate)
    ' Change the year value
    Me.cmbYear = Year(varDate)
    ' Re-draw the calendar
    SetDays
    ' Make sure the correct box is highlighted
    Me.optCalendar.Value = Day(varDate)
    ' Save the new date in hidden form control for later
    Me.ctlCalendar = varDate
End Sub

Private Sub cmdSave_Click()
    ' Hiding this dialog lets the calling code in GetDate continue
    Me.Visible = False
End Sub

Private Sub optCalendar_AfterUpdate()
    ' Every time the user picks a new date box
    '   update the saved date value
    varDate = DateSerial(Me.cmbYear, Me.cmbMonth, optCalendar.Value)
    ' .. and update the hidden form control for later
    Me.ctlCalendar = varDate
End Sub

Private Sub Form_Load()
    ' Establish an initial value for the date
    If IsNothing(Me.OpenArgs) Then
        varDate = Date
    Else
        ' Should have date, time, and "DateOnly" indicator in OpenArgs:
        '   mm/dd/yyyy hh:mm,-1
        varDate = left(Me.OpenArgs, 10)
        Me.txtHour = Mid(Me.OpenArgs, 12, 2)
        Me.txtMinute = Mid(Me.OpenArgs, 15, 2)
        ' If "date only"
        If right(Me.OpenArgs, 2) = "-1" Then
            ' Hide some stuff
            Me.txtHour.Visible = False
            Me.txtMinute.Visible = False
            Me.lblColon.Visible = False
            Me.lblTimeInstruct.Visible = False
            Me.SetFocus
            '  .. and resize my window
            DoCmd.MoveSize , , , 4295
        End If
    End If
    ' Initialize the month selector
    Me.cmbMonth = Month(varDate)
    ' Initialize the year selector
    Me.cmbYear = Year(varDate)
    ' Call the common calendar draw routine
    SetDays
    ' Place the date/time value in a hidden control -
    '  The calling routine fetches it from here
    Me.ctlCalendar = varDate
    ' Highlight the correct day box in the calendar
    Me.optCalendar = Day(varDate)
End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)
Dim intHour As Integer
    ' Trapping key presses in the Hour box
    If KeyAscii = 43 Or KeyAscii = 61 Then  ' Plus sign key - add one to hour
        KeyAscii = 0  ' Swallow the Plus key
        ' Should have a value, but if not, set to 1
        If IsNothing(Me.txtHour) Then
            intHour = 1
        Else
            intHour = Me.txtHour + 1
        End If
        ' If we've wrapped to 24, then reset to zero
        If intHour = 24 Then intHour = 0
        ' Update the text box
        Me.txtHour = intHour
        ' Done
        Exit Sub
    End If
    
    If KeyAscii = 45 Or KeyAscii = 95 Then  ' Minus sign key - subtract one
        KeyAscii = 0  ' Swallow the Minus key
        ' Should have a value, but if not, set to zero
        If IsNothing(Me.txtHour) Then
            intHour = 0
        Else
            intHour = Me.txtHour
        End If
        intHour = intHour - 1
        ' If we've gone below zero, the wrap to 23
        If intHour = -1 Then intHour = 23
        ' Update the text box
        Me.txtHour = intHour
        ' Done
        Exit Sub
    End If
    ' All other key codes pass inspection
    ' The Input Mask allows only numbers and +/-
End Sub

Private Sub txtMinute_KeyPress(KeyAscii As Integer)
Dim intMinute As Integer
    ' Trapping key presses in the Minute box
    If KeyAscii = 43 Or KeyAscii = 61 Then  ' Plus sign key - add one to minute
        KeyAscii = 0  ' Swallow the Plus key
        ' Should have a value, but if not, set to 1
        If IsNothing(Me.txtMinute) Then
            intMinute = 1
        Else
            intMinute = Me.txtMinute + 1
        End If
        ' If we've wrapped to 60, then reset to zero
        If intMinute = 60 Then intMinute = 0
        ' Update the text box
        Me.txtMinute = intMinute
        ' Done
        Exit Sub
    End If
    
    If KeyAscii = 45 Or KeyAscii = 95 Then  ' Minus sign key - subtract one
        KeyAscii = 0  ' Swallow the Minus key
        ' Should have a value, but if not, set to zero
        If IsNothing(Me.txtMinute) Then
            intMinute = 0
        Else
            intMinute = Me.txtMinute
        End If
        intMinute = intMinute - 1
        ' If we've gone below zero, the wrap to 59
        If intMinute = -1 Then intMinute = 59
        ' Update the text box
        Me.txtMinute = intMinute
        ' Done
        Exit Sub
    End If
    ' All other key codes pass inspection
    ' The Input Mask allows only numbers and +/-
End Sub

Private Sub SetDays()
Dim intI As Integer, intJ As Integer, strNum As String, ctl As control

    ' Move the focus so we can dink with the calendar
    Me.cmbMonth.SetFocus
    
    ' First, clear all the boxes
    For intI = 1 To 42
        ' Controls are named "tglnn"
        ' Where nn = 01 to 42
        ' Using Format to get 2 digits
        strNum = Format(intI, "00")
        ' Establish a pointer to the control
        Set ctl = Me("tgl" & strNum)
        ' Clear the day number
        ctl.Caption = ""
        ' Reset the Option Value
        ctl.OptionValue = 0
        ' Reset the ForeColor to black
        ctl.ForeColor = 0
        ' And disable it
        ctl.Enabled = False
    Next intI

    intMonth = Me!cmbMonth
    intYear = Me!cmbYear
    ' The first box to set it the weekday of the first day of the month
    intFirst = Weekday(DateSerial(intYear, intMonth, 1), vbSunday)
    ' Calculate the last day number
    '   by adding 1 month to Day 1 and subtracting one
    intLastDay = Day(DateAdd("m", 1, DateSerial(intYear, intMonth, 1)) - 1)
    ' .. and the last box to set
    intLast = intFirst + intLastDay - 1
    
    ' Now set up all the boxes for the current month
    intJ = 1
    For intI = intFirst To intLast
        strNum = Format(intI, "00")
        ' Establish a pointer to the control
        Set ctl = Me("tgl" & strNum)
        ' Put the day number in the associated label caption
        ctl.Caption = intJ
        ' Set the value of the Toggle
        ctl.OptionValue = intJ
        ' Set the Fore Color to Blue
        ctl.ForeColor = 16711680
        ' and Enable it
        ctl.Enabled = True
        intJ = intJ + 1
    Next intI
    
    Set ctl = Nothing
    
    ' Fill in the remaining days for the next month
    If intLast <> 42 Then
        intJ = 1
        For intI = intLast + 1 To 42
            strNum = Format(intI, "00")
            ' Put the day number in the associated label caption
            Me("tgl" & strNum).Caption = intJ
            intJ = intJ + 1
        Next intI
    End If
    
    ' .. and the days from the previous month
    If intFirst <> 1 Then
        intJ = Day(DateSerial(intYear, intMonth, 1) - 1)
        For intI = intFirst - 1 To 1 Step -1
            strNum = Format(intI, "00")
            ' Put the day number in the associated label caption
            Me("tgl" & strNum).Caption = intJ
            intJ = intJ - 1
        Next intI
    End If

    
    ' Put the focus back
    Me.optCalendar.SetFocus
        
End Sub
