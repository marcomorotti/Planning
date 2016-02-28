Version =20
VersionRequired =20
Begin Form
    MaxButton = NotDefault
    NavigationButtons = NotDefault
    PictureTiling = NotDefault
    AllowAdditions = NotDefault
    ScrollBars =2
    PictureAlignment =5
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4677
    DatasheetFontHeight =9
    ItemSuffix =18
    Left =1440
    Top =2790
    Right =6510
    Bottom =6225
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0430295f1523e440
    End
    RecordSource ="SELECT tblConsumi.Cod_Art, tblConsumi.Anno, tblConsumi.Mese, tblConsumi.Consumo,"
        " tblConsumi.N_Spedito_Mese FROM tblConsumi ORDER BY tblConsumi.Cod_Art, DateSeri"
        "al([Anno],[Mese],1) DESC; "
    Caption ="Consumi"
    DatasheetFontName ="Tahoma"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextAlign =3
            FontSize =9
            BackColor =12632256
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
        Begin Line
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
        Begin Image
            SpecialEffect =3
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
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
            FontSize =9
            FontWeight =400
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
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
            SpecialEffect =2
            BorderWidth =3
            LabelX =230
            LabelY =-30
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
            BackStyle =1
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
        Begin BoundObjectFrame
            SpecialEffect =3
            BorderColor =12632256
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
        Begin TextBox
            TextAlign =1
            FontSize =9
            BorderColor =12632256
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
            ShowDatePicker =1
        End
        Begin ListBox
            SpecialEffect =3
            BackColor =12632256
            BorderColor =12632256
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
        Begin ComboBox
            TextAlign =1
            FontSize =9
            BorderColor =12632256
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
        Begin Subform
            BorderColor =12632256
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
        Begin UnboundObjectFrame
            SpecialEffect =3
            BackStyle =0
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
            ForeColor =128
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
        Begin Tab
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
        Begin FormHeader
            Height =283
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontCharSet =204
                    TextAlign =1
                    Width =1134
                    Height =255
                    FontSize =10
                    Name ="TransactionDate Label"
                    Caption ="Anno"
                    FontName ="Verdana"
                    EventProcPrefix ="TransactionDate_Label"
                    LayoutCachedWidth =1134
                    LayoutCachedHeight =255
                End
                Begin Label
                    OverlapFlags =95
                    TextFontCharSet =204
                    TextAlign =2
                    Left =1134
                    Width =1134
                    Height =255
                    FontSize =10
                    Name ="PurchaseOrderID Label"
                    Caption ="Mese"
                    FontName ="Verdana"
                    EventProcPrefix ="PurchaseOrderID_Label"
                    LayoutCachedLeft =1134
                    LayoutCachedWidth =2268
                    LayoutCachedHeight =255
                End
                Begin Label
                    OverlapFlags =95
                    TextFontCharSet =204
                    TextAlign =2
                    Left =2268
                    Width =1134
                    Height =255
                    FontSize =10
                    Name ="TransactionDescription Label"
                    Caption ="Consumo"
                    FontName ="Verdana"
                    EventProcPrefix ="TransactionDescription_Label"
                    LayoutCachedLeft =2268
                    LayoutCachedWidth =3402
                    LayoutCachedHeight =255
                End
                Begin Label
                    OverlapFlags =87
                    TextFontCharSet =204
                    Left =3402
                    Width =1134
                    Height =255
                    FontSize =10
                    Name ="UnitsOrdered Label"
                    Caption ="N. Spedizioni"
                    FontName ="Verdana"
                    EventProcPrefix ="UnitsOrdered_Label"
                    LayoutCachedLeft =3402
                    LayoutCachedWidth =4536
                    LayoutCachedHeight =255
                End
            End
        End
        Begin Section
            Height =354
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextFontCharSet =204
                    Left =30
                    Top =30
                    Width =1134
                    Height =284
                    ColumnWidth =1365
                    FontSize =10
                    Name ="TransactionDate"
                    ControlSource ="Anno"
                    FontName ="Verdana"

                    LayoutCachedLeft =30
                    LayoutCachedTop =30
                    LayoutCachedWidth =1164
                    LayoutCachedHeight =314
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =95
                    TextFontCharSet =204
                    TextAlign =2
                    Left =1164
                    Top =30
                    Width =1134
                    Height =284
                    ColumnWidth =975
                    FontSize =10
                    TabIndex =1
                    Name ="PurchaseOrderID"
                    ControlSource ="=MonthName([Mese],True)"
                    FontName ="Verdana"

                    LayoutCachedLeft =1164
                    LayoutCachedTop =30
                    LayoutCachedWidth =2298
                    LayoutCachedHeight =314
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =95
                    TextFontCharSet =204
                    TextAlign =3
                    Left =2298
                    Top =30
                    Width =1134
                    Height =284
                    ColumnWidth =1635
                    FontSize =10
                    TabIndex =2
                    Name ="TransactionDescription"
                    ControlSource ="Consumo"
                    FontName ="Verdana"

                    LayoutCachedLeft =2298
                    LayoutCachedTop =30
                    LayoutCachedWidth =3432
                    LayoutCachedHeight =314
                End
                Begin TextBox
                    OverlapFlags =87
                    TextFontCharSet =204
                    TextAlign =3
                    Left =3432
                    Top =30
                    Width =1134
                    Height =284
                    ColumnWidth =1470
                    FontSize =10
                    TabIndex =3
                    Name ="UnitsOrdered"
                    ControlSource ="N_Spedito_Mese"
                    FontName ="Verdana"

                    LayoutCachedLeft =3432
                    LayoutCachedTop =30
                    LayoutCachedWidth =4566
                    LayoutCachedHeight =314
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
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


Private Sub ViewPurchaseOrders()
On Error GoTo Err_ViewPurchaseOrders
    If IsNull(Forms![Products]![ProductID]) Then
        MsgBox "Enter product information before entering purchase order."
    Else
        DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
        DoCmd.OpenForm "Purchase Orders"
        If Me![PurchaseOrderID] > 0 Then
            DoCmd.GoToControl "PurchaseOrderID"
            DoCmd.FindRecord Me![PurchaseOrderID]
        Else
            If Not IsNull(Forms![Purchase Orders]![PurchaseOrderID]) Then
                DoCmd.DoMenuItem acFormBar, 3, 0, , acMenuVer70
            End If
        End If
    End If

Exit_ViewPurchaseOrders:
    Exit Sub

Err_ViewPurchaseOrders:
    MsgBox Err.Description
    Resume Exit_ViewPurchaseOrders
End Sub
