Version =20
VersionRequired =20
Begin Form
    Modal = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =4320
    DatasheetFontHeight =10
    ItemSuffix =8
    Left =180
    Top =90
    Right =10545
    Bottom =5670
    DatasheetGridlinesColor =12632256
    ShortcutMenuBar ="Form Shortcut Bar"
    Toolbar ="Custom Form Toolbar"
    RecSrcDt = Begin
        0x8eca8895beaee140
    End
    Caption ="Calendar"
    MenuBar ="Custom Form Menu Bar"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0xa0050000a0050000a0050000a005000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
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
        Begin CommandButton
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
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
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
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
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
        Begin FormHeader
            Height =374
            Name ="FormHeader"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =83
                    Left =2880
                    Top =14
                    FontSize =9
                    FontWeight =700
                    ForeColor =0
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
                    Left =14
                    Top =14
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =0
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
            Height =4260
            Name ="Detail"
            Begin
                Begin CustomControl
                    Enabled = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =85
                    AutoActivate =1
                    Name ="Calendar1"
                    OleData = Begin
                        0x000e0000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff020000002bc9278e64121c108a2f0402 ,
                        0x24009c02000000000000000000000000200bcb7fcd83be0107000000c0010000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000038000000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff01000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000010000007a010000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fefffffffdfffffffffffffffffffffffffffffffffffffffeffffff ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff020000002bc9278e64121c108a2f0402 ,
                        0x24009c02000000000000000000000000a089c548568fc20105000000c0010000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000038000000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff01000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000010000007a010000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffefffffffeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffeffffff0200000003000000040000000500000006000000feffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff38000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000800c41d0000d8130000d2070c00010005000080000000000000 ,
                        0xa000100000800000a00001000100020000000100000001000000010000000100 ,
                        0x0000010000000100000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000000000000000000000001000000bc024442010005417269616c0100 ,
                        0x000090014442010005417269616c01000000bc02c0d4010005417269616c0000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OLEClass ="Calendar"
                    ControlTipText ="Double Click to Select"
                    Class ="MSCAL.Calendar.7"

                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =3
                    Left =1875
                    Top =3180
                    Width =120
                    Height =225
                    FontWeight =700
                    BackColor =12632256
                    Name ="lblColon"
                    Caption =":"
                    FontName ="Arial"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextFontFamily =0
                    Left =1230
                    Top =3180
                    Width =645
                    TabIndex =1
                    Name ="txtHour"
                    Format ="00"
                    DefaultValue ="12"
                    InputMask ="09"
                    OnKeyPress ="[Event Procedure]"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Press + or - keys to scroll value."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =1234
                            Top =2925
                            Width =570
                            Height =255
                            FontWeight =700
                            BackColor =12632256
                            Name ="lblHour"
                            Caption ="Hour:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =1
                    OverlapFlags =87
                    TextFontFamily =0
                    Left =2010
                    Top =3180
                    Width =645
                    TabIndex =2
                    Name ="txtMinute"
                    Format ="00"
                    DefaultValue ="0"
                    InputMask ="09"
                    OnKeyPress ="[Event Procedure]"
                    ShortcutMenuBar ="Form Control Shortcut Bar"
                    ControlTipText ="Press + or - keys to scroll value."

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =2014
                            Top =2925
                            Width =465
                            Height =255
                            FontWeight =700
                            BackColor =12632256
                            Name ="lblMinute"
                            Caption ="Min:"
                            FontName ="Arial"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =105
                    Top =3450
                    Width =4035
                    Height =675
                    FontWeight =700
                    BackColor =12632256
                    Name ="lblTimeInstruct"
                    Caption ="Press Tab to move from Calendar to Hour / Min boxes.  Type in Hour (24 hour cloc"
                        "k) and Minute values or use + and - keys to change the values."
                    FontName ="Arial"
                End
            End
        End
        Begin FormFooter
            Height =0
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

' This form demonstrates both using a custom control (MSCal.OCX)
' and manipulating a Class via Property Set
' See also the GetDateOCX function that activates this form/module.

' Place to save the "date only" indicator
Dim intDateOnly As Integer
' Variable for the Property Set
Dim ctlThisControl As control
' Optional variable for the Property Set
Dim intSet As Integer
' Place to save the date value
Dim varDate As Variant

Private Sub cmdCancel_Click()
    ' Close without saving
    DoCmd.Close acForm, Me.name
End Sub

Private Sub cmdSave_Click()
    ' Saves the changed value back in the calling control
    
    ' Do some error trapping here in case the calling control can't
    ' accept a date/time value.
    On Error GoTo Save_Error
    
    ' Make sure we got a valid control to point to
    If intSet Then
        ' OK - save the value
        If (intDateOnly = -1) Then
            ' Passing back date only
            ctlThisControl.Value = Me.Calendar1.Value
        Else
            ' Do date and time
            ctlThisControl.Value = Me.Calendar1.Value + TimeValue(Me.txtHour & ":" & Me.txtMinute)
        End If
    End If
    
Save_Exit:
    DoCmd.Close acForm, Me.name
    Exit Sub
    
Save_Error:
    MsgBox "An error occured attempting to save the date value.", vbCritical, "Scm Portafoglio Ordini"
    ErrorLog "frmCalendarOCX_Save", Err, Error
    Resume Save_Exit
    
End Sub

Private Sub Form_Load()
    ' Hide myself until properties are set
    Me.Visible = False
End Sub

Public Property Set ctlToUpdate(Optional intD As Integer = 0, ctl As control)
' This procedure is called as a property of the Class Module
' GetDateOCX opens this form by creating a new instance of the class
'  and then sets the required properties via a SET statement.

    ' First, validate the kind of control passed
    Select Case ctl.ControlType
        ' Text box, combo box, and list box are OK
        Case acTextBox, acListBox, acComboBox
        Case Else
            MsgBox "Invalid control passed to the Calendar."
            DoCmd.Close acForm, Me.name
    End Select
    
    ' Save the pointer to the control to update
    Set ctlThisControl = ctl
    
    ' Save the date only value
    intDateOnly = intD
    ' If "date only"
    If (intDateOnly = -1) Then
        ' Resize my window
        DoCmd.MoveSize , , , 3935
        ' Hide some stuff just to be sure
        Me.txtHour.Visible = False
        Me.txtMinute.Visible = False
        Me.lblColon.Visible = False
        Me.lblTimeInstruct.Visible = False
        Me.SetFocus
    End If
    
    ' Set the flag to indicate we got the pointer
    intSet = True
    ' Save the "current" value of the control
    varDate = ctlThisControl.Value
    ' Make sure we got a valid date value
    If Not IsDate(varDate) Then
        ' If not, set the default to today
        varDate = Now
        Me.Calendar1.Value = Date
        Me.txtHour = Format(Hour(varDate), "00")
        Me.txtMinute = Format(Minute(varDate), "00")
    Else
        ' Otherwise, set the date/time to the one in the control
        ' Make sure we have a Date data type, not just text
        varDate = CDate(varDate)
        Me.Calendar1.Value = varDate
        Me.txtHour = Format(Hour(varDate), "00")
        Me.txtMinute = Format(Minute(varDate), "00")
    End If

End Property

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

Private Sub Calendar1_DblClick()

    ' If they double-click the calendar, act as though they clicked Save
    cmdSave_Click

End Sub
