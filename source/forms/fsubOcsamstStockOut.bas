Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    ScrollBars =2
    ViewsAllowed =1
    GridX =24
    GridY =24
    Width =16700
    ItemSuffix =66
    Left =1020
    Top =1815
    Right =18135
    Bottom =6540
    RecSrcDt = Begin
        0xc12cee6416e6e340
    End
    RecordSource ="qryCOrdersStockOut"
    OnCurrent ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    Begin
        Begin Label
            BackColor =-2147483633
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
        Begin Subform
            SpecialEffect =2
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
        Begin Tab
            Width =5103
            Height =3402
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
            Height =637
            BackColor =16777088
            Name ="FormHeader1"
            OnClick ="[Event Procedure]"
            Begin
                Begin Rectangle
                    OverlapFlags =93
                    Left =15900
                    Top =45
                    Width =765
                    Height =570
                    BackColor =16777164
                    Name ="Casella49"
                    LayoutCachedLeft =15900
                    LayoutCachedTop =45
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =615
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =1110
                    Top =240
                    Width =1065
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Label20"
                    Caption ="N. DOC"
                    FontName ="Tahoma"
                    LayoutCachedLeft =1110
                    LayoutCachedTop =240
                    LayoutCachedWidth =2175
                    LayoutCachedHeight =480
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =3360
                    Top =240
                    Width =990
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Label21"
                    Caption ="N. CLIENTE"
                    FontName ="Tahoma"
                    LayoutCachedLeft =3360
                    LayoutCachedTop =240
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =495
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =11460
                    Top =45
                    Width =600
                    Height =510
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Label23"
                    Caption ="Qtà \015\012Ord"
                    FontName ="Tahoma"
                    LayoutCachedLeft =11460
                    LayoutCachedTop =45
                    LayoutCachedWidth =12060
                    LayoutCachedHeight =555
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =12165
                    Top =45
                    Width =1050
                    Height =510
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta30"
                    Caption ="Data\015\012Prev."
                    FontName ="Tahoma"
                    LayoutCachedLeft =12165
                    LayoutCachedTop =45
                    LayoutCachedWidth =13215
                    LayoutCachedHeight =555
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =6765
                    Top =244
                    Width =1335
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta33"
                    Caption ="CODICE ART"
                    FontName ="Tahoma"
                    LayoutCachedLeft =6765
                    LayoutCachedTop =244
                    LayoutCachedWidth =8100
                    LayoutCachedHeight =499
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =4605
                    Top =255
                    Width =1605
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta36"
                    Caption ="CLIENTE"
                    FontName ="Tahoma"
                    LayoutCachedLeft =4605
                    LayoutCachedTop =255
                    LayoutCachedWidth =6210
                    LayoutCachedHeight =495
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =8239
                    Top =240
                    Width =2205
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta39"
                    Caption ="DESCRIZIONE"
                    FontName ="Tahoma"
                    LayoutCachedLeft =8239
                    LayoutCachedTop =240
                    LayoutCachedWidth =10444
                    LayoutCachedHeight =480
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =13230
                    Top =45
                    Width =812
                    Height =510
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta44"
                    Caption ="Qtà\015\012Sped."
                    FontName ="Tahoma"
                    LayoutCachedLeft =13230
                    LayoutCachedTop =45
                    LayoutCachedWidth =14042
                    LayoutCachedHeight =555
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    Left =15870
                    Top =45
                    Width =795
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta46"
                    Caption ="DISP."
                    FontName ="Tahoma"
                    LayoutCachedLeft =15870
                    LayoutCachedTop =45
                    LayoutCachedWidth =16665
                    LayoutCachedHeight =300
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Top =240
                    Width =330
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta55"
                    Caption ="P"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =14070
                    Top =45
                    Width =855
                    Height =510
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta56"
                    Caption ="Qtà\015\012Giacent."
                    FontName ="Tahoma"
                    LayoutCachedLeft =14070
                    LayoutCachedTop =45
                    LayoutCachedWidth =14925
                    LayoutCachedHeight =555
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =223
                    TextAlign =2
                    Left =14970
                    Top =45
                    Width =1035
                    Height =510
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta59"
                    Caption ="Qtà Tot\015\012Impegn."
                    FontName ="Tahoma"
                    LayoutCachedLeft =14970
                    LayoutCachedTop =45
                    LayoutCachedWidth =16005
                    LayoutCachedHeight =555
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =215
                    TextAlign =2
                    Left =15885
                    Top =300
                    Width =765
                    Height =255
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta60"
                    Caption ="Stefani"
                    FontName ="Tahoma"
                    LayoutCachedLeft =15885
                    LayoutCachedTop =300
                    LayoutCachedWidth =16650
                    LayoutCachedHeight =555
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =405
                    Top =240
                    Width =555
                    Height =240
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta62"
                    Caption ="UT."
                    FontName ="Tahoma"
                    LayoutCachedLeft =405
                    LayoutCachedTop =240
                    LayoutCachedWidth =960
                    LayoutCachedHeight =480
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =2265
                    Width =1050
                    Height =510
                    FontSize =10
                    FontWeight =700
                    BackColor =12632256
                    ForeColor =8404992
                    Name ="Etichetta65"
                    Caption ="Data\015\012Ordine"
                    FontName ="Tahoma"
                    LayoutCachedLeft =2265
                    LayoutCachedWidth =3315
                    LayoutCachedHeight =510
                End
            End
        End
        Begin Section
            Height =270
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    Width =285
                    Height =240
                    FontSize =10
                    TabIndex =1
                    Name ="txtliv_urgenza"
                    ControlSource ="liv_urgenza"

                    LayoutCachedWidth =285
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =93
                    Left =11385
                    Width =737
                    Height =240
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="txtqta_ord_umv"
                    ControlSource ="qta_ord_umv"

                    LayoutCachedLeft =11385
                    LayoutCachedWidth =12122
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1084
                    Width =1140
                    Height =240
                    ColumnWidth =2775
                    FontSize =10
                    TabIndex =3
                    Name ="ptPartTypeDescription"
                    ControlSource ="NUMERO_DOC"
                    StatusBarText ="Only the standard part types"
                    FontName ="Tahoma"

                    LayoutCachedLeft =1084
                    LayoutCachedWidth =2224
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =3360
                    Width =735
                    Height =240
                    FontSize =10
                    Name ="txtOcsaOhNumeratore"
                    ControlSource ="COD_CLI"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =3360
                    LayoutCachedWidth =4095
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    IMEHold = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    IMESentenceMode =3
                    Left =6668
                    Height =240
                    FontSize =10
                    TabIndex =4
                    Name ="txtCOD_ART"
                    ControlSource ="COD_ART"

                    LayoutCachedLeft =6668
                    LayoutCachedWidth =8108
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =8108
                    Width =3225
                    Height =240
                    FontSize =10
                    TabIndex =5
                    Name ="Testo34"
                    ControlSource ="Descrizione"

                    LayoutCachedLeft =8108
                    LayoutCachedWidth =11333
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4088
                    Width =2580
                    Height =240
                    FontSize =10
                    TabIndex =6
                    Name ="txtOcsaLnLin"
                    ControlSource ="DS_RAG_SOC"

                    LayoutCachedLeft =4088
                    LayoutCachedWidth =6668
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =13219
                    Width =737
                    Height =240
                    FontSize =10
                    TabIndex =7
                    Name ="txtQta_cons_umv"
                    ControlSource ="qta_cons_umv"

                    LayoutCachedLeft =13219
                    LayoutCachedWidth =13956
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =13999
                    Width =737
                    Height =240
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    ForeColor =8388608
                    Name ="txtESISTENZA"
                    ControlSource ="=IIf(IsNull([DISP]),0,[DISP])"
                    Format ="General Number"

                    LayoutCachedLeft =13999
                    LayoutCachedWidth =14736
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =87
                    TextAlign =1
                    Left =12122
                    Width =1054
                    Height =240
                    FontSize =10
                    TabIndex =9
                    Name ="txtData_Prev_Cons"
                    ControlSource ="Data_Prev_Cons"

                    LayoutCachedLeft =12122
                    LayoutCachedWidth =13176
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =14809
                    Width =737
                    Height =240
                    FontSize =10
                    TabIndex =10
                    Name ="txtImpegnato"
                    ControlSource ="Impegnato"
                    Format ="General Number"

                    LayoutCachedLeft =14809
                    LayoutCachedWidth =15546
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextFontCharSet =2
                    TextAlign =3
                    TextFontFamily =2
                    Left =15619
                    Width =240
                    Height =257
                    FontSize =11
                    TabIndex =11
                    ForeColor =8388863
                    Name ="txtDispRed"
                    ControlSource ="=IIf([txtEsistenza]-([txtqta_Ord_Umv]-[txtQta_Cons_Umv])<0,\"t\",IIf([txtEsisten"
                        "za]-([txtImpegnato])<0,\"2\",\"\"))"
                    Format ="General Number"
                    FontName ="Wingdings"

                    LayoutCachedLeft =15619
                    LayoutCachedWidth =15859
                    LayoutCachedHeight =257
                End
                Begin TextBox
                    OverlapFlags =85
                    Left =15904
                    Width =767
                    Height =240
                    FontSize =10
                    TabIndex =12
                    Name ="Testo60"
                    ControlSource ="DispAh"

                    LayoutCachedLeft =15904
                    LayoutCachedWidth =16671
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =300
                    Width =735
                    Height =240
                    FontSize =10
                    TabIndex =13
                    Name ="Testo63"
                    ControlSource ="UTENTE"

                    LayoutCachedLeft =300
                    LayoutCachedWidth =1035
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =2265
                    Width =1054
                    Height =240
                    FontSize =10
                    TabIndex =14
                    Name ="Testo64"
                    ControlSource ="DATA_ORDINE"

                    LayoutCachedLeft =2265
                    LayoutCachedWidth =3319
                    LayoutCachedHeight =240
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
Option Compare Database
Option Explicit


Private Sub Detail0_Click()
On Error GoTo Err_Detail0_Click:

Form.fsubOcsaMstP0.Requery
Exit_Err_Detail0_Click:
    Exit Sub

Err_Detail0_Click:
    Resume Exit_Err_Detail0_Click:
End Sub
Private Sub Form_Current()
[Forms]![frmOcsamstStockOut]![fsubOcsaMstPOStockOut].Requery
End Sub


' =IIf([txtEsistenza]-([txtqta_Ord_Umv]-[txtQta_Cons_Umv])<0;"2";IIf([txtImpegnato]-([txtEsistenza])<0;"t";""))

'Private Sub Form_Load()
'Select Case txtESISTENZA - ([txtqta_ord_umv] - [txtQta_cons_umv]) < 0
'        Case "2"
'            txtDispRed.ForeColor = vbRed
'            'txtDispRed = "Unacceptable"
''        Case 31 To 42
''            Your_txtBox_Name.ForeColor = vbBrown
''            Your_txtBox_Name = "Marginal"
'        Case "t"
'            txtDispRed.ForeColor = vbYellow
''            Your_txtBox_Name = "Effective"
''        Case 57 To 71
''            Your_txtBox_Name.ForeColor = vbGreen
''            Your_txtBox_Name = "Very Good"
''        Case 72 To 80
''            Your_txtBox_Name.ForeColor = vbGreen
''            Your_txtBox_Name = "Outstanding"
'        Case Else
'            txtDispRed = ""
'    End Select
'End Sub

Private Sub FormHeader1_Click()
Form.fsubOcsaMstP0.Requery
End Sub
