Version =20
VersionRequired =20
Begin Form
    NavigationButtons = NotDefault
    AllowUpdating =1
    ScrollBars =2
    ViewsAllowed =1
    GridX =24
    GridY =24
    Width =16215
    ItemSuffix =60
    Left =705
    Top =6645
    Right =17370
    Bottom =8625
    RecSrcDt = Begin
        0x89f9a47656e5e340
    End
    RecordSource ="SELECT tblPOrders.NUMERO_DOC, tblPOrders.COD_ART, tblPOrders.COD_FORN, tblPOrder"
        "s.RAG_SOC_FORN, tblPOrders.DATA_ORDINE, tblPOrders.QTA_ORDINE, tblPOrders.QTA_RE"
        "SIDUA, tblPOrders.DATA_RIC FROM tblPOrders WHERE (((tblPOrders.COD_ART) Like [fo"
        "rms]![frmOcsaMst]![fsubOcsaMst]![txtCod_Art])); "
    Caption ="frmPartsReceipt"
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
        Begin Page
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
        Begin FormHeader
            Height =840
            BackColor =8454016
            Name ="FormHeader1"
            Begin
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Top =570
                    Width =1140
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label20"
                    Caption ="N DOC"
                    FontName ="Tahoma"
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =93
                    Left =1290
                    Top =570
                    Width =1560
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Label21"
                    Caption ="COD FORNITORE"
                    FontName ="Tahoma"
                    LayoutCachedLeft =1290
                    LayoutCachedTop =570
                    LayoutCachedWidth =2850
                    LayoutCachedHeight =810
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Width =3975
                    Height =420
                    FontSize =14
                    FontWeight =700
                    BackColor =12632256
                    Name ="lblServiceOrder"
                    Caption ="ORDINI ACQUISTO"
                    LayoutCachedWidth =3975
                    LayoutCachedHeight =420
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =87
                    Left =2858
                    Top =566
                    Width =4545
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Etichetta53"
                    Caption ="RAGIONE SOCIALE"
                    FontName ="Tahoma"
                    LayoutCachedLeft =2858
                    LayoutCachedTop =566
                    LayoutCachedWidth =7403
                    LayoutCachedHeight =806
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =7410
                    Top =566
                    Width =1065
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Etichetta52"
                    Caption ="DATA ORD."
                    FontName ="Tahoma"
                    LayoutCachedLeft =7410
                    LayoutCachedTop =566
                    LayoutCachedWidth =8475
                    LayoutCachedHeight =806
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =9570
                    Top =311
                    Width =1065
                    Height =495
                    FontSize =9
                    FontWeight =700
                    Name ="Etichetta54"
                    Caption ="QTA' \015\012RESIDUA"
                    FontName ="Tahoma"
                    LayoutCachedLeft =9570
                    LayoutCachedTop =311
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =806
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    Left =10648
                    Top =566
                    Width =1620
                    Height =240
                    FontSize =9
                    FontWeight =700
                    Name ="Etichetta55"
                    Caption ="DATA CONSEGNA"
                    FontName ="Tahoma"
                    LayoutCachedLeft =10648
                    LayoutCachedTop =566
                    LayoutCachedWidth =12268
                    LayoutCachedHeight =806
                End
                Begin Label
                    BackStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    Left =8478
                    Top =311
                    Width =1065
                    Height =495
                    FontSize =9
                    FontWeight =700
                    Name ="Etichetta60"
                    Caption ="QTA' \015\012ORDINE"
                    FontName ="Tahoma"
                    LayoutCachedLeft =8478
                    LayoutCachedTop =311
                    LayoutCachedWidth =9543
                    LayoutCachedHeight =806
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =259
            BackColor =-2147483633
            Name ="Detail0"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    Left =30
                    Width =1020
                    Height =240
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    ForeColor =255
                    Name ="txtOrderN"
                    ControlSource ="COD_ART"

                    LayoutCachedLeft =30
                    LayoutCachedWidth =1050
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1270
                    Width =1437
                    Height =256
                    FontSize =10
                    TabIndex =1
                    Name ="txtPartsDescriptionFR"
                    ControlSource ="COD_FORN"

                    LayoutCachedLeft =1270
                    LayoutCachedWidth =2707
                    LayoutCachedHeight =256
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =9608
                    Width =1062
                    Height =256
                    FontSize =10
                    TabIndex =2
                    Name ="txtUm"
                    ControlSource ="QTA_RESIDUA"
                    Format ="General Number"

                    LayoutCachedLeft =9608
                    LayoutCachedWidth =10670
                    LayoutCachedHeight =256
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =85
                    Left =10693
                    Width =1190
                    Height =256
                    FontSize =10
                    TabIndex =3
                    RightMargin =113
                    Name ="txtDATA_RIC"
                    ControlSource ="DATA_RIC"

                    LayoutCachedLeft =10693
                    LayoutCachedWidth =11883
                    LayoutCachedHeight =256
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =247
                    IMESentenceMode =3
                    Width =1247
                    Height =256
                    ColumnWidth =2775
                    FontSize =10
                    Name ="txtRab"
                    ControlSource ="NUMERO_DOC"
                    StatusBarText ="Only the standard part types"

                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2730
                    Width =4692
                    Height =256
                    FontSize =10
                    TabIndex =5
                    Name ="Testo50"
                    ControlSource ="RAG_SOC_FORN"

                    LayoutCachedLeft =2730
                    LayoutCachedWidth =7422
                    LayoutCachedHeight =256
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =7445
                    Width =1055
                    Height =256
                    FontSize =10
                    TabIndex =6
                    Name ="txtDATA_ORDINE"
                    ControlSource ="DATA_ORDINE"

                    LayoutCachedLeft =7445
                    LayoutCachedWidth =8500
                    LayoutCachedHeight =256
                End
                Begin TextBox
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =8523
                    Width =1062
                    Height =256
                    FontSize =10
                    TabIndex =7
                    Name ="Testo59"
                    ControlSource ="QTA_ORDINE"
                    Format ="General Number"

                    LayoutCachedLeft =8523
                    LayoutCachedWidth =9585
                    LayoutCachedHeight =256
                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
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

Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.name
End Sub

Private Sub txtBuild_AfterUpdate()
    Requery
End Sub

Private Sub txtQtyAccettata_AfterUpdate()
' Private Sub txt01_AfterUpdate()
Dim strSQL As String
Dim strWhere As String
Dim strUM As String
Dim intUM As Variant
Dim LastID_Parts As Integer
Dim Causale As Integer

Causale = 1 'Entrata acquisto

' Se Parts Type, Inventory, Bin location, Qualità  sono vuoti esce Sub
If IsNothing(Me!txtID_PartsType) Then
    MsgBox "You must insert PARTS TYPE", vbExclamation, gstrAppTitle
    Me!txtQtyAccettata = Null
    Me!txtID_PartsType.SetFocus
    Exit Sub
ElseIf IsNothing(Me!txtID_Inventory) Then
MsgBox "You must insert INVENTORY", vbExclamation, gstrAppTitle
    Me!txtQtyAccettata = Null
    Me!txtID_Inventory.SetFocus
    Exit Sub
ElseIf IsNothing(Me!txtBinLocation) Then
MsgBox "You must insert BIN LOCATION", vbExclamation, gstrAppTitle
    Me!txtQtyAccettata = Null
    Me!txtBinLocation.SetFocus
    Exit Sub
ElseIf IsNothing(Me!ID_QualitaMateriale) Then
MsgBox "You must insert QUALITY MATERIAL", vbExclamation, gstrAppTitle
    Me!txtQtyAccettata = Null
    Me!txtID_QualitaMateriale.SetFocus
    Exit Sub
End If
    
' *** CERCA Unità di Misura
' -1- Setta PZ =NR
If StrConv(Me!txtUm, vbUpperCase) = "PZ" Then 'Setta l'unità di misura PZ = NR
    intUM = DLookup("ID_UmType", "tblUmType", "UM = 'NR'")
Else
    intUM = DLookup("ID_UmType", "tblUmType", "UM = '" & StrConv(Me!txtUm, vbUpperCase) & "'")
End If

' -2- Se non esiste UM nel Db la inserisce in tblUmType
If IsNothing(intUM) Then
    strUM = StrConv(Me!txtUm, vbUpperCase)
            
    strSQL = "INSERT INTO tblUMType ([UM])" & _
    "VALUES ('" & strUM & "')"
    
    CurrentDb.Execute strSQL, dbFailOnError
    
' -3- Carica anagrafica Parts in tblParts
    'Cerca Unità di Misura
    intUM = DLookup("ID_UmType", "tblUmType", "UM = '" & StrConv(Me!txtUm, vbUpperCase) & "'")
End If
    On Error GoTo Err_txtQtyAccettata_AfterUpdate
    strSQL = "INSERT INTO tblParts ([ID_PartsType]," & _
    "[PartsNumber]," & _
    "[ID_Um]," & _
    "[PartsDescription], [DateIns])" & _
    "VALUES (" & Me!txtID_PartsType & ", " & _
    "'" & Me!txtRab & "', " & _
    intUM & ", " & _
    "'" & SQLize(Me!txtGoodsDescription) & "', " & _
    "(#" & Format(Date, "mm/dd/yyyy") & "#)" & ")"
    
    CurrentDb.Execute strSQL, dbFailOnError
    
' -4- Carica Movimento magazzino
LastID_Parts = DMax("ID_Parts", "tblParts")  'Cerca l'ultimo recors inserito
If IsNull(LastID_Parts) Then LastID_Parts = 1 'Se non esistono record setta contatore = 1

strSQL = "INSERT INTO tblInventoryMovement" & _
    "([ID_Inventory]," & _
    "[ID_Item]," & _
    "[Qty]," & _
    "[Causale], " & _
    "[DbDocName], " & _
    "[DocName], " & _
    "[ID_QualitaMateriale]" & ", " & _
    "[DateIns])" & _
    "VALUES (" & Me!txtID_Inventory & ", " & _
    LastID_Parts & ", " & _
    "'" & Me!txtQtyAccettata & "', " & _
    "'" & Causale & "', " & _
    "'" & Me!txtID_PackingList & "', " & _
    "'" & Me!txtOrderN & "', " & _
    Me!txtID_QualitaMateriale & ", " & _
    "(#" & Format(Date, "mm/dd/yyyy") & "#)" & ")"

CurrentDb.Execute strSQL, dbFailOnError

' -5- Inserisce Mappatura di magazzino
strSQL = "INSERT INTO tblInventoryMap ([ID_Inventory]," & _
    "[ID_Part]," & _
    "[BinLocation] )" & _
    "VALUES (" & Me!txtID_Inventory & ", " & _
    LastID_Parts & ", " & _
    "'" & Me!txtBinLocation & "')"
    
    CurrentDb.Execute strSQL, dbFailOnError
    

' -6- Aggiorna flag Movimentato su Packing List
Me!txtMovement = True
MsgBox ("Item" & Me!txtGoodsDescription & " is been insert into Inventory" _
                    & " ")
'strSQL = "UPDATE tblPackingList SET [Movement] = True" & _
         " WHERE [ID_PackingList] = " & Me!txtID_PackingList
         
'DoCmd.SetWarnings False 'Silenzioso

'CurrentDb.Execute strSQL, dbFailOnError
' DoCmd.RunCommand acCmdRecordsGoToNext

' -7- Fine Sql
Requery

Exit_txtQtyAccettata_AfterUpdate:
    Exit Sub

Err_txtQtyAccettata_AfterUpdate:
    Select Case Err.Number
        Case 3022 'ignore duplicate keys
            'strSQL = "UPDATE tblMezziAttendance SET [TipoAttendance] = " & "'" & Me!txt01.Value & "'" & _
            '" WHERE [ID_MezziAttendance] = " & Me!txtID_Mezzi & "AND [AttendanceDate] = " & "(#" & _
            'Format(DateSerial(intYear, intMonth, 1), "mm/dd/yyyy") & "#)"
            'CurrentDb.Execute strSQL, dbFailOnError
            'DoCmd.RunCommand acCmdRecordsGoToNext
           
        Case Else
            MsgBox Err.Number & "-" & Err.Description
    End Select
    Resume Exit_txtQtyAccettata_AfterUpdate
End Sub ' ---Fine
Private Sub txtID_2PartsType_AfterUpdate()
    Me!fsubPartsReceipt.Form!txtID_3PartsType.Enabled = False
    Me!fsubPartsReceipt.Form!txtID_3PartsType = Null
    Me!fsubPartsReceipt.Form!txtID_4PartsType = Null
    Me!fsubPartsReceipt.Form!txtID_5PartsType = Null
    Me!fsubPartsReceipt.Form!txtId_6PartsType = Null
    Me!fsubPartsReceipt.Form!txtId_7PartsType = Null
    Me!fsubPartsReceipt.Form!txtID_8PartsType = Null
If Not IsNull([Forms]![frmPartsReceipt]![fsubPartsReceipt].[Form]![txtID_2PartsType]) Then
    Me!fsubPartsReceipt.Form!txtID_3PartsType.Enabled = True
    Me!fsubPartsReceipt.Form!txtID_4PartsType.Enabled = False
    Me!fsubPartsReceipt.Form!txtID_5PartsType.Enabled = False
    Me!fsubPartsReceipt.Form!txtId_6PartsType.Enabled = False
    Me!fsubPartsReceipt.Form!txtId_7PartsType.Enabled = False
    Me!fsubPartsReceipt.Form!txtID_8PartsType.Enabled = False
End If
End Sub
Private Sub txtID_2PartsType_GotFocus()
If ([Forms]![frmPartsReceipt]![fsubPartsReceipt].[Form]![txtID_PartsType]) = 0 _
    Or IsNull([Forms]![frmPartsReceipt]![fsubPartsReceipt].[Form]![txtID_PartsType]) Then
    MsgBox "Please Specify PARTS TYPE first"
    'txtID_PartsType.SetFocus
    '[Forms]![frmPartsReceipt]![fsubPartsReceipt].[Form]![txtID_PartsType].SetFocus
    Me!fsubPartsReceipt.Form!txtID_PartsType.SetFocus
    Me!fsubPartsReceipt.Form!txtID_2PartsType.Enabled = False
Else
    Me!txtID_2PartsType.Requery
End If
End Sub
Private Sub txtID_3PartsType_AfterUpdate()
    Me!txtID_4PartsType.Enabled = False
    Me!txtID_4PartsType = Null
    Me!txtID_5PartsType = Null
    Me!txtId_6PartsType = Null
    Me!txtId_7PartsType = Null
    Me!txtID_8PartsType = Null
If Not IsNull([Forms]![frmPartsReceipt]![fsubPartsReceipt].[Form]![txtID_3PartsType]) Then
    Me!txtID_4PartsType.Enabled = True
    Me!txtID_5PartsType.Enabled = False
    Me!txtId_6PartsType.Enabled = False
    Me!txtId_7PartsType.Enabled = False
    Me!txtID_8PartsType.Enabled = False
End If
End Sub
Private Sub txtID_3PartsType_GotFocus()
If Me!txtID_2PartsType = 0 Or IsNull(Me!txtID_2PartsType) Then
    MsgBox "Please Specify 2 PARTS TYPE first"
    Me!txtID_2PartsType.SetFocus
    Me!txtID_3PartsType.Enabled = False
Else
    Me!txtID_3PartsType.Requery
End If
End Sub
Private Sub txtID_4PartsType_AfterUpdate()
    Me!txtID_5PartsType.Enabled = False
    Me!txtID_5PartsType = Null
    Me!txtId_6PartsType = Null
    Me!txtId_7PartsType = Null
    Me!txtID_8PartsType = Null
If Not IsNull([Forms]![frmPartsReceipt]![fsubPartsReceipt].[Form]![txtID_4PartsType]) Then
    Me!txtID_5PartsType.Enabled = True
    Me!txtId_6PartsType.Enabled = False
    Me!txtId_7PartsType.Enabled = False
    Me!txtID_8PartsType.Enabled = False
End If
End Sub
Private Sub txtID_4PartsType_GotFocus()
If Me!txtID_3PartsType = 0 Or IsNull(Me!txtID_3PartsType) Then
    MsgBox "Please Specify 3 PARTS TYPE first"
    Me!txtID_3PartsType.SetFocus
    Me!txtID_4PartsType.Enabled = False
Else
    Me!txtID_4PartsType.Requery
End If
End Sub
Private Sub txtID_5PartsType_AfterUpdate()
    Me!txtId_6PartsType.Enabled = False
    Me!txtId_6PartsType = Null
    Me!txtId_7PartsType = Null
    Me!txtID_8PartsType = Null
If Not IsNull([Forms]![frmPartsReceipt]![fsubPartsReceipt].[Form]![txtID_5PartsType]) Then
    Me!txtId_6PartsType.Enabled = True
    Me!txtId_7PartsType.Enabled = False
    Me!txtID_8PartsType.Enabled = False
End If
End Sub
Private Sub txtID_5PartsType_GotFocus()
If Me!txtID_4PartsType = 0 Or IsNull(Me!txtID_4PartsType) Then
    MsgBox "Please Specify 4 PARTS TYPE first"
    Me!txtID_4PartsType.SetFocus
    Me!txtID_5PartsType.Enabled = False
Else
    Me!txtID_5PartsType.Requery
End If
End Sub
Private Sub txtId_6PartsType_AfterUpdate()
    Me!txtId_7PartsType.Enabled = False
    Me!txtId_7PartsType = Null
    Me!txtID_8PartsType = Null
If Not IsNull([Forms]![frmPartsReceipt]![fsubPartsReceipt].[Form]![txtId_6PartsType]) Then
    Me!txtId_7PartsType.Enabled = True
    Me!txtID_8PartsType.Enabled = False
End If
End Sub
Private Sub txtId_6PartsType_GotFocus()
If Me!txtID_5PartsType = 0 Or IsNull(Me!txtID_5PartsType) Then
    MsgBox "Please Specify 5 PARTS TYPE first"
    Me!txtID_5PartsType.SetFocus
    Me!txtId_6PartsType.Enabled = False
Else
    Me!txtId_6PartsType.Requery
End If
End Sub
Private Sub txtId_7PartsType_AfterUpdate()
    Me!txtID_8PartsType.Enabled = False
    Me!txtID_8PartsType = Null
If Not IsNull([Forms]![frmPartsReceipt]![fsubPartsReceipt].[Form]![txtId_7PartsType]) Then
    Me!txtID_8PartsType.Enabled = True
End If
End Sub
Private Sub txtId_7PartsType_GotFocus()
If Me!txtId_6PartsType = 0 Or IsNull(Me!txtId_6PartsType) Then
    MsgBox "Please Specify 6 PARTS TYPE first"
    Me!txtId_6PartsType.SetFocus
    Me!txtId_7PartsType.Enabled = False
Else
    Me!txtId_7PartsType.Requery
End If
End Sub
Private Sub txtID_8PartsType_GotFocus()
If Me!txtId_7PartsType = 0 Or IsNull(Me!txtId_7PartsType) Then
    MsgBox "Please Specify 7 PARTS TYPE first"
    Me!txtId_7PartsType.SetFocus
    Me!txtID_8PartsType.Enabled = False
Else
    Me!txtID_8PartsType.Requery
End If
End Sub
Private Sub txtBuild2_AfterUpdate()
  Requery
End Sub
