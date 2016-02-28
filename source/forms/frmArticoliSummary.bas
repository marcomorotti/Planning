Version =20
VersionRequired =20
Begin Form
    AllowEditing = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    ViewsAllowed =1
    Width =8445
    ItemSuffix =39
    Left =3540
    Top =2700
    Right =12360
    Bottom =10275
    Filter ="[Cod_Art] LIKE '0001344*'"
    RecSrcDt = Begin
        0x307835372ee2e340
    End
    RecordSource ="tblArticoli"
    Caption ="Ricerca articoli"
    PrtMip = Begin
        0xd0020000a0050000d0020000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =255
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    Begin
        Begin Label
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
            DisplayWhen =2
            Height =900
            BackColor =12632256
            Name ="FormHeader1"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =645
                    Top =45
                    Width =3750
                    Height =375
                    FontSize =14
                    BackColor =8421376
                    ForeColor =16777215
                    Name ="Text8"
                    Caption ="Articoli  trovati"
                    LayoutCachedLeft =645
                    LayoutCachedTop =45
                    LayoutCachedWidth =4395
                    LayoutCachedHeight =420
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =18
                    Left =-15
                    Top =540
                    Width =945
                    Height =285
                    FontSize =9
                    Name ="Text10"
                    Caption ="Codice"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =-15
                    LayoutCachedTop =540
                    LayoutCachedWidth =930
                    LayoutCachedHeight =825
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextFontFamily =18
                    Left =1800
                    Top =540
                    Width =1020
                    Height =285
                    FontSize =9
                    Name ="Text12"
                    Caption ="Descrizione"
                    FontName ="Times New Roman"
                    LayoutCachedLeft =1800
                    LayoutCachedTop =540
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =825
                End
            End
        End
        Begin Section
            Height =377
            BackColor =12632256
            Name ="Detail0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    Left =180
                    Top =75
                    Width =1185
                    BorderColor =1
                    Name ="txtCod_art"
                    ControlSource ="Cod_art"
                    StatusBarText ="Club ID"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"

                    LayoutCachedLeft =180
                    LayoutCachedTop =75
                    LayoutCachedWidth =1365
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =1425
                    Top =75
                    Width =7020
                    TabIndex =1
                    BorderColor =1
                    Name ="txtDes_art"
                    ControlSource ="Des_art"
                    StatusBarText ="Club Name"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"

                    LayoutCachedLeft =1425
                    LayoutCachedTop =75
                    LayoutCachedWidth =8445
                    LayoutCachedHeight =330
                End
            End
        End
        Begin FormFooter
            DisplayWhen =2
            Height =465
            BackColor =12632256
            Name ="FormFooter2"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =86
                    TextFontFamily =34
                    Left =45
                    Top =45
                    Width =1455
                    Name ="Details"
                    Caption ="&Vedi dettagli"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="View details for current contact."
                    UnicodeAccessKey =86

                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =67
                    TextFontFamily =34
                    Left =6975
                    Top =45
                    TabIndex =1
                    Name ="Close"
                    Caption ="&Chiudi"
                    OnClick ="[Event Procedure]"
                    FontName ="Arial"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Close this form"
                    UnicodeAccessKey =67

                    LayoutCachedLeft =6975
                    LayoutCachedTop =45
                    LayoutCachedWidth =8415
                    LayoutCachedHeight =405
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
Option Compare Database   'Use database order for string comparisons
Option Explicit

Private Sub Close_Click()
    DoCmd.Close acForm, Me.name
End Sub

Private Sub AName_DblClick(Cancel As Integer)
    Details_Click
End Sub

Private Sub MezziID_DblClick(Cancel As Integer)
    Details_Click
End Sub

Private Sub Details_Click()
Dim strFilter As String
    ' They asked for details (or double-clicked one of the controls)
    ' Set up the filter
    strFilter = "(Cod_art = '" & Me.txtCod_Art & "')"
    ' Open Mezzi filtered on the current row
    DoCmd.OpenForm FormName:="frmParts1", WhereCondition:=strFilter
    ' Close me
    DoCmd.Close acForm, Me.name
    ' Put focus on Mezzi
    Forms!frmParts1.SetFocus
End Sub
