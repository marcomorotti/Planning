Option Compare Database

Function fEnumControls(ByVal strfrmToEnum As String)
' 20160225 Marco
' Stampa i Controlli di una Form
' non visualizza i Controlli di una SubForm
' Es. ?fEnumControls("frmParts1")
Dim astrCtlName() As String
Dim i As Integer, intCnt As Integer
Dim frm As Form
'Un-Comment the next two lines for Access 95
'Const acPage = 124
'Const acTabCtl = 123
    'if form is closed, exit function
    If Not IsLoaded(strfrmToEnum) Then
        MsgBox "Form " & strfrmToEnum & " is probably closed!! " & _
            vbCrLf & "Please open it & try again.", vbCritical
        Exit Function
    End If

    Set frm = Forms(strfrmToEnum)
    'Count the number of controls
    intCnt = frm.Count
    
    'Initialize the array to hold control names
    ReDim astrCtlName(0 To intCnt - 1, 0 To 1)

    For i = 0 To intCnt - 1
        astrCtlName(i, 0) = frm(i).name
        'Use ControlType to determine the Type of Control
        Select Case frm(i).ControlType
            Case acLabel: astrCtlName(i, 1) = "Label"
            Case acRectangle: astrCtlName(i, 1) = "Rectangle"
            Case acLine: astrCtlName(i, 1) = "Line"
            Case acImage: astrCtlName(i, 1) = "Image"
            Case acCommandButton: astrCtlName(i, 1) = "Command Button"
            Case acOptionButton: astrCtlName(i, 1) = "Option button"
            Case acCheckBox: astrCtlName(i, 1) = "Check box"
            Case acOptionGroup: astrCtlName(i, 1) = "Option group"
            Case acBoundObjectFrame: astrCtlName(i, 1) = "Bound object frame"
            Case acTextBox: astrCtlName(i, 1) = "Text Box"
            Case acListBox: astrCtlName(i, 1) = "List box"
            Case acComboBox: astrCtlName(i, 1) = "Combo box"
            Case acSubform: astrCtlName(i, 1) = "SubForm"
            Case acObjectFrame: astrCtlName(i, 1) = "Unbound object frame or chart"
            Case acPageBreak: astrCtlName(i, 1) = "Page break"
            Case acPage: astrCtlName(i, 1) = "Page"
            Case acCustomControl: astrCtlName(i, 1) = "ActiveX (custom) control"
            Case acToggleButton: astrCtlName(i, 1) = "Toggle Button"
            Case acTabCtl: astrCtlName(i, 1) = "Tab Control"
        End Select
    Next i
    
    'Print out the array in an orderly fashion
    Debug.Print "Control Name", "Control Type"
    Debug.Print "------------", "------------"
    For i = 0 To intCnt - 1
        Debug.Print astrCtlName(i, 0), astrCtlName(i, 1)
    Next i
    Erase astrCtlName
End Function