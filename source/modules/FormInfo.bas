Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' From Access 2000 Developer's Handbook, Volume I
' by Getz, Litwin, and Gilbert (Sybex)
' Copyright 1999.  All rights reserved.

' FormInfo Class module

' Generally, leave this False. If you want to
' debug odd behaviors, change it to True.
#Const DEBUGGING = False

' ==================================
' Windows API declarations.
' ==================================

Private Declare Function SendMessage Lib "user32" _
 Alias "SendMessageA" _
 (ByVal hWnd As Long, ByVal wMsg As Long, _
 ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
 Alias "SetWindowLongA" (ByVal hWnd As Long, _
 ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowLong Lib "user32" _
 Alias "GetWindowLongA" (ByVal hWnd As Long, _
 ByVal nIndex As Long) As Long

Private Declare Function ShowWindow _
 Lib "user32" _
 (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function MoveWindow _
 Lib "user32" _
 (ByVal hWnd As Long, _
 ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, _
 ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private Declare Function GetWindowRect _
 Lib "user32" _
 (ByVal hWnd As Long, _
 lpRect As RECT) As Long

Private Declare Function GetParent _
 Lib "user32" _
 (ByVal hWnd As Long) As Long

Private Declare Function IsIconic _
 Lib "user32" _
 (ByVal hWnd As Long) As Long
 
 Private Declare Function IsZoomed _
  Lib "user32" _
  (ByVal hWnd As Long) As Long

Private Declare Function GetClientRect _
 Lib "user32" _
 (ByVal hWnd As Long, _
 lpRect As RECT) As Long

Private Declare Function GetDeviceCaps _
 Lib "gdi32" _
 (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function GetSystemMetrics _
 Lib "user32" _
 (ByVal nIndex As Long) As Long
 
Private Declare Function GetDC _
 Lib "user32" _
 (ByVal hWnd As Long) As Long

Private Declare Function ReleaseDC _
 Lib "user32" _
 (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Declare Function SystemParametersInfoRect _
 Lib "user32" Alias "SystemParametersInfoA" _
 (ByVal uAction As Long, ByVal uParam As Long, _
 ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
 
Private Declare Function ClientToScreen _
 Lib "user32" _
 (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
    x As Long
    Y As Long
End Type

Private Const SPI_GETWORKAREA = 48

Private Const SM_CYCAPTION = 4
Private Const SM_CXFULLSCREEN = 16
Private Const SM_CYFULLSCREEN = 17

Private Const LOGPIXELSX = 88        '  Logical pixels/inch in X
Private Const LOGPIXELSY = 90        '  Logical pixels/inch in Y

Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Const SW_NORMAL = 1

' GetWindowLong Constant
Private Const GWL_STYLE = -16
Private Const WS_SYSMENU = &H80000

' Windows Style constant
Private Const WS_CAPTION = &HC00000

' Windows message constant.
Private Const WM_NCPAINT = &H85

' Store rectangle coordinates.
Private Type RECT
    left As Long
    Top As Long
    right As Long
    bottom As Long
End Type

' ==================================
' Constants
' ==================================

' Windows 95/98/NT4/2000 puts a 2-pixel
' border around the MDI client area, which
' doesn't get taken into account automatically.
' If you're using NT 3.51, you're on your own.
Private Const adhcBorderWidthX = 2
Private Const adhcBorderWidthY = 2

Private Const adhcTop = "Top"
Private Const adhcLeft = "Left"
Private Const adhcRight = "Right"
Private Const adhcBottom = "Bottom"
Private Const adhcWidth = "Width"
Private Const adhcHeight = "Height"
Private Const adhcState = "State"

Private Const adhcTwipsPerInch = 1440

' ==================================
' Enums
' ==================================

Private Enum ReferenceType
    rtPopup = 0
    rtNormal = 1
End Enum

Public Enum WindowState
    wsNormal = 0
    wsMinimized = 1
    wsMaximized = 2
End Enum

' ==================================
' Private variables
' ==================================

Private frm As Form

Private mptCurrentScreen As POINTAPI
Private mptTwipsPerPixel As POINTAPI
Private mptClientOffset As POINTAPI

' Store form coordinates.
Private Type COORDS
    left As Long
    Top As Long
    Width As Long
    Height As Long
    State As WindowState
End Type

Private mCoords As COORDS
Private mfIsSubform As Boolean
Private mfIsPopup As Boolean


' ==================================
' Public Methods and Properties
' ==================================

Public Property Get ScreenX() As Long
    ' Return the current screen size, in pixels.
    ScreenX = mptCurrentScreen.x
End Property

Public Property Get ScreenY() As Long
    ' Return the current screen size, in pixels.
    ScreenY = mptCurrentScreen.Y
End Property

Public Property Get ScreenXinTwips() As Long
    ' Return the current screen size, in pixels.
    ScreenXinTwips = mptCurrentScreen.x * mptTwipsPerPixel.x
End Property

Public Property Get ScreenYInTwips() As Long
    ' Return the current screen size, in pixels.
    ScreenYInTwips = mptCurrentScreen.Y * mptTwipsPerPixel.Y
End Property

Public Property Get WindowState() As WindowState
    ' Retrieve the current window state
    ' as a single value.
    On Error GoTo HandleErrors
    If IsMaximized Then
        WindowState = wsMaximized
    ElseIf IsMinimized Then
        WindowState = wsMinimized
    Else
        WindowState = wsNormal
    End If

ExitHere:
    Exit Property
    
HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.WindowState", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let WindowState(Value As WindowState)
    On Error GoTo HandleErrors
    Select Case Value
        Case wsMaximized
            IsMaximized = True
        Case wsMinimized
            IsMinimized = True
        Case wsNormal
            ' Just set one of minimized
            ' or maximized to False. That
            ' will restore the window.
            IsMaximized = False
    End Select
    
ExitHere:
    Exit Property
    
HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.WindowState", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Sub RetrieveCoords(AppName As String)

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    Dim strName As String
    Dim c As COORDS

    On Error GoTo HandleErrors

    ' Use the name of the application as the highest
    ' level, and the form's name as the next level.
    ' This way, you could have multiple forms in the same
    ' app use this code.
    strName = frm.name
    With c
        .State = Val(GetSetting(AppName, _
         strName, adhcState, Me.WindowState))
        Select Case .State
            Case wsMinimized
                IsMinimized = True
            Case wsNormal
                .Top = _
                 GetSetting(AppName, strName, adhcTop, .Top)
                .left = _
                 GetSetting(AppName, strName, adhcLeft, .left)
                .Width = _
                 GetSetting(AppName, strName, adhcWidth, .Width)
                .Height = _
                 GetSetting(AppName, strName, adhcHeight, .Height)
        
                ' Only muck with the form's size if
                ' you get values for width and height
                ' that make sense.
                If .Width > 0 And .Height > 0 Then
                    Call SetSize(.left, .Top, .Width, .Height)
                End If
            Case wsMaximized
                ' Don't set this form to be maximized
                ' or not. If the MDI Client window is maximized
                ' you want to bypass these settings and just
                ' maximize the window.
        End Select
    End With
   
ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.RetrieveCoords", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Public Sub SaveCoords(AppName As String)

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    Dim strName As String

    On Error GoTo HandleErrors

    strName = frm.name
    
    ' Use the name of the application as the highest
    ' level, and the form's name as the next level.
    ' This way, you could have multiple forms in the same
    ' app use this code.
    
    Call SaveSetting( _
     AppName, strName, adhcState, WindowState)
    Call SaveSetting( _
     AppName, strName, adhcTop, Top)
    Call SaveSetting( _
     AppName, strName, adhcLeft, left)
    Call SaveSetting( _
     AppName, strName, adhcWidth, Width)
    Call SaveSetting( _
     AppName, strName, adhcHeight, Height)
    
ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.SaveCoords", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Public Property Let PercentWidth(Value As Double)
    ' Size the window to the specified percentage
    ' of the available screen width.
    On Error GoTo HandleErrors
    Dim cParent As COORDS
    
    Call GetCoords(frm.hWnd, mCoords)
    Call GetParentCoords(frm.hWnd, cParent)
    mCoords.Width = cParent.Width * Value / 100
    Call SetSize(Width:=mCoords.Width)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.PercentWidth", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get PercentWidth() As Double
    On Error GoTo HandleErrors
    
    Dim cParent As COORDS
    Call GetCoords(frm.hWnd, mCoords)
    Call GetParentCoords(frm.hWnd, cParent)
    PercentWidth = mCoords.Width / cParent.Width * 100

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.PercentWidth", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let PercentHeight(Value As Double)
    ' Size the window to the specified percentage
    ' of the available screen Height.
    On Error GoTo HandleErrors
    Dim cParent As COORDS
    
    Call GetCoords(frm.hWnd, mCoords)
    Call GetParentCoords(frm.hWnd, cParent)
    mCoords.Height = cParent.Height * Value / 100
    Call SetSize(Height:=mCoords.Height)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.PercentHeight", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get PercentHeight() As Double
    On Error GoTo HandleErrors
    
    Dim cParent As COORDS
    Call GetCoords(frm.hWnd, mCoords)
    Call GetParentCoords(frm.hWnd, cParent)
    PercentHeight = mCoords.Height / cParent.Height * 100

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.PercentHeight", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get IsPopup() As Boolean
    ' You can't tell from properties if a
    ' form is a Popup form or not -- you
    ' can either set the popup property to
    ' True, or you can open it with the
    ' acDialog flag. This property
    ' checks the parent of the form, and
    ' if it's Access, this is a popup form.
    IsPopup = mfIsPopup
End Property


Public Property Get IsSubForm() As Boolean
    IsSubForm = mfIsSubform
End Property

Public Property Get IsMaximized() As Boolean
    On Error GoTo HandleErrors
    IsMaximized = IsZoomed(frm.hWnd)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.IsMaximized", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let IsMaximized(MaxIt As Boolean)
    On Error GoTo HandleErrors
    If MaxIt Then
        Call ShowWindow(frm.hWnd, SW_MAXIMIZE)
    Else
        Call ShowWindow(frm.hWnd, SW_NORMAL)
    End If

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.IsMaximized", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get IsMinimized() As Boolean
    On Error GoTo HandleErrors
    IsMinimized = IsIconic(frm.hWnd)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.IsMinimized", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let IsMinimized(MinIt As Boolean)
    On Error GoTo HandleErrors
    If MinIt Then
        Call ShowWindow(frm.hWnd, SW_MINIMIZE)
    Else
        Call ShowWindow(frm.hWnd, SW_NORMAL)
    End If

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.IsMinimized", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Set Form(Value As Form)
    On Error GoTo HandleErrors
    Set frm = Value
    Call GetScreenInfo
    Call GetClientOffsets
    mfIsPopup = GetIsPopup()
    mfIsSubform = GetIsSubform()
    
ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Form", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get Form() As Form
    On Error GoTo HandleErrors
    Set Form = frm

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Form", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get TwipsPerPixelX() As Long
    ' Return the screen's vertical Twips/Pixel ratio.
    On Error GoTo HandleErrors
    TwipsPerPixelX = mptTwipsPerPixel.x

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.TwipsPerPixelX", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get TwipsPerPixelY() As Long
    ' Return the screen's horizontal Twips/Pixel ratio.
    On Error GoTo HandleErrors
    TwipsPerPixelY = mptTwipsPerPixel.Y

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.TwipsPerPixelY", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Sub GetSize( _
 Optional ByRef left As Long, Optional ByRef Top As Long, _
 Optional ByRef Width As Long, Optional ByRef Height As Long, _
 Optional ByVal InTwips As Boolean = False)
   
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ' Get all the coords at once.
    ' If you want more than one coordinate,
    ' this will be slightly more efficient than
    ' requesting each one individually.
    On Error GoTo HandleErrors
    Dim intMultiplierX As Integer
    Dim intMultiplierY As Integer
    
    If InTwips Then
        intMultiplierX = mptTwipsPerPixel.x
        intMultiplierY = mptTwipsPerPixel.Y
    Else
        intMultiplierX = 1
        intMultiplierY = 1
    End If
    Call GetCoords(frm.hWnd, mCoords)
    
    With mCoords
        left = .left * intMultiplierX
        Top = .Top * intMultiplierY
        Width = .Width * intMultiplierX
        Height = .Height * intMultiplierY
    End With

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.GetSize", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Public Sub SetSize( _
 Optional ByVal left As Variant, _
 Optional ByVal Top As Variant, _
 Optional ByVal Width As Variant, _
 Optional ByVal Height As Variant, _
 Optional ByVal InTwips As Boolean = False)

    ' Set the form's location/size, either in
    ' pixels (the default) or twips. If you specify
    ' twips, it's the same as using DoCmd.MoveSize,
    ' except that this can work with any form, not just
    ' the current form. MoveSize is limited to
    ' the current form, only.

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    With mCoords
        If IsMissing(left) Then
            left = .left
        End If
        If Not IsNumeric(left) Then
            left = .left
        End If
        If InTwips Then
            left = left \ mptTwipsPerPixel.x
        End If
        
        If IsMissing(Top) Then
            Top = .Top
        End If
        If Not IsNumeric(Top) Then
            Top = .Top
        End If
        If InTwips Then
            Top = Top \ mptTwipsPerPixel.Y
        End If
        
        If IsMissing(Width) Then
            Width = .Width
        End If
        If Not IsNumeric(Width) Then
            Width = .Width
        End If
        If InTwips Then
            Width = Width \ mptTwipsPerPixel.x
        End If
        
        If IsMissing(Height) Then
            Height = .Height
        End If
        If Not IsNumeric(Height) Then
            Height = .Height
        End If
        If InTwips Then
            Height = Height \ mptTwipsPerPixel.Y
        End If
    End With
    Call MoveWindow(frm.hWnd, _
     left, Top, Width, Height, 1)

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.SetSize", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Public Property Get left() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    left = mCoords.left

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Left", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let left(Value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hWnd, _
         Value, .Top, .Width, .Height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Left", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get LeftInTwips() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    LeftInTwips = mCoords.left * mptTwipsPerPixel.x

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.LeftInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let LeftInTwips(Value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hWnd, _
         Value \ mptTwipsPerPixel.x, .Top, .Width, .Height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.LeftInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get Top() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    Top = mCoords.Top

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Top", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let Top(Value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hWnd, _
         .left, Value, .Width, .Height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Top", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get TopInTwips() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    TopInTwips = mCoords.Top * mptTwipsPerPixel.Y

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.TopInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let TopInTwips(Value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hWnd, _
         .left, Value \ mptTwipsPerPixel.Y, .Width, .Height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.TopInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get Width() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    Width = mCoords.Width

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Width", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let Width(Value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hWnd, _
         .left, .Top, Value, .Height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Width", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get WidthInTwips() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    WidthInTwips = mCoords.Width * mptTwipsPerPixel.x

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.WidthInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let WidthInTwips(Value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hWnd, _
         .left, .Top, Value \ mptTwipsPerPixel.x, .Height, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.WidthInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property


Public Property Get Height() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    Height = mCoords.Height

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Height", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let Height(Value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hWnd, _
         .left, .Top, .Width, Value, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Height", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get HeightInTwips() As Long
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    HeightInTwips = mCoords.Height * mptTwipsPerPixel.Y

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.HeightInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let HeightInTwips(Value As Long)
    On Error GoTo HandleErrors
    Call GetCoords(frm.hWnd, mCoords)
    With mCoords
        Call MoveWindow(frm.hWnd, _
         .left, .Top, .Width, Value \ mptTwipsPerPixel.Y, 1)
    End With

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.HeightInTwips", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Sub FillClientArea()
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    ' Move a form so that it fills the entire
    ' MDI Client area in Access. This allows you to
    ' remove the Close button (or the entire
    ' caption bar, if you like) and "maximize" the
    ' form without Access adding those pesky buttons
    ' back.

    Dim hWndParent As Long
    Dim cParent As COORDS
    
    On Error GoTo HandleErrors
        
    If IsZoomed(frm.hWnd) Then
        Call ShowWindow(frm.hWnd, SW_NORMAL)
    End If
    Call GetParentCoords(frm.hWnd, cParent)
    With cParent
        Call MoveWindow(frm.hWnd, _
         .left, .Top, .Width, .Height, 1)
    End With
    
ExitHere:
    Exit Sub
    
HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.FillClientArea", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere

End Sub

Public Sub Center()

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ' Center a form within the confines of the
    ' MDI parent window. If the form is a popup
    ' form, then center on the screen.
    
    On Error GoTo HandleErrors
    Dim cParent As COORDS

    ' Get the coordinates of the current form.
    Call GetCoords(frm.hWnd, mCoords)
    Call GetParentCoords(frm.hWnd, cParent)
    
    ' Calculate the width of the child form,
    ' and calculate its new coordinates, relative
    ' to its parent.
    With mCoords
        .left = (cParent.Width - .Width) \ 2
        .Top = (cParent.Height - .Height) \ 2
        
        ' Move the child window to its new location.
        Call MoveWindow(frm.hWnd, _
         .left, .Top, .Width, .Height, 1)
    End With

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.Center", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Public Property Let ShowSystemMenu(ShowIt As Boolean)
    On Error GoTo HandleErrors
    Dim lngOldStyle As Long
    Dim lngNewStyle As Long
    
    ' Show or hide a form's system menu, depending
    ' on the value in ShowIt.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ' If there is nothing to do, get out.
    If Me.ShowSystemMenu = ShowIt Then
        Exit Property
    End If

    ' Get the current window style of the form.
    lngOldStyle = GetWindowLong(frm.hWnd, GWL_STYLE)
    
    If ShowIt Then
    ' Turn on the bit that enables system menu.
    lngNewStyle = lngOldStyle Or WS_SYSMENU
    Else
    ' Turn off the bit the shows the system menu.
    lngNewStyle = lngOldStyle And Not WS_SYSMENU
    End If
    
    ' Set the new window style.
    Call SetWindowLong(frm.hWnd, GWL_STYLE, lngNewStyle)
    
    ' The 1 as the third parameter tells
    ' the window to repaint its entire border.
    Call SendMessage(frm.hWnd, WM_NCPAINT, 1, 0)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.ShowSystemMenu", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get ShowSystemMenu() As Boolean
    
    ' Retrieve info about a form's system menu.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    On Error GoTo HandleErrors
    Dim lngOldStyle As Long
  
    lngOldStyle = GetWindowLong(frm.hWnd, GWL_STYLE)
    ShowSystemMenu = ((lngOldStyle And WS_SYSMENU) = WS_SYSMENU)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.ShowSystemMenu", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Let ShowCaptionBar(ShowIt As Boolean)

    ' Show or remove a form's caption bar, depending
    ' on the value in ShowIt.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    On Error GoTo HandleErrors
    Dim lngOldStyle As Long
    Dim lngNewStyle As Long
    Dim rct As RECT
    Dim intDiff As Integer
    
    ' If there is nothing to do, get out.
    If Me.ShowCaptionBar = ShowIt Then
        GoTo ExitHere
    End If

    Call GetCoords(frm.hWnd, mCoords)
    
    ' Get the current window style of the form.
    lngOldStyle = GetWindowLong(frm.hWnd, GWL_STYLE)
    
    If ShowIt Then
        ' Turn off the bit that enables the caption.
        lngNewStyle = lngOldStyle Or WS_CAPTION
    Else
        ' Turn off the bit that enables the caption.
        lngNewStyle = lngOldStyle And Not WS_CAPTION
    End If
    
    ' Set the new window style.
    lngOldStyle = SetWindowLong(frm.hWnd, _
     GWL_STYLE, lngNewStyle)
    
    ' How much room does that caption take up?
    intDiff = GetSystemMetrics(SM_CYCAPTION)
    
    ' Calculate the new height.
    If ShowIt Then
        mCoords.Height = mCoords.Height + intDiff
    Else
        mCoords.Height = mCoords.Height - intDiff
    End If

    ' Move the window to the same left and top,
    ' but with new width and height.
    ' This will make the new form appear
    ' a little shorter or a little taller.
    With mCoords
        Call MoveWindow(frm.hWnd, _
         .left, .Top, .Width, .Height, 1)
    End With
    Call GetClientOffsets
    
ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.ShowCaptionBar", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get ShowCaptionBar() As Boolean

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    On Error GoTo HandleErrors
    Dim lngOldStyle As Long
    Dim lngNewStyle As Long
    Dim rct As RECT
    Dim intDX As Integer
    Dim intDY As Integer

    ' Get the current window style of the form.
    lngOldStyle = GetWindowLong(frm.hWnd, GWL_STYLE)

    ' Turn off the bit that enables the caption.
    ShowCaptionBar = ((lngOldStyle And WS_CAPTION) = WS_CAPTION)

ExitHere:
    Exit Property

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.ShowCaptionBar", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Property

Public Property Get ClientOffsetXinTwips() As Long
    ' Retrieve the client offset, converted to twips.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ClientOffsetXinTwips = mptClientOffset.x * mptTwipsPerPixel.x
End Property

Public Property Get ClientOffsetYinTwips() As Long
    ' Retrieve the client offset, converted to twips.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ClientOffsetYinTwips = mptClientOffset.Y * mptTwipsPerPixel.Y
End Property

Public Property Get ClientOffsetX() As Long
    ' Retrieve the client offset, in pixels.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ClientOffsetX = mptClientOffset.x
End Property

Public Property Get ClientOffsetY() As Long
    ' Retrieve the client offset, in pixels.
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.

    ClientOffsetY = mptClientOffset.Y
End Property

' ==================================
' Private Methods
' ==================================

Private Sub GetCoords(hWnd As Long, cToFill As COORDS)

    ' Fill in rct with the coordinates of the client area.
    ' Fill in rctWindow with coordinates of the window.

    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    On Error GoTo HandleErrors
    Dim hWndParent As Long
    Dim rctParent As RECT
    Dim rct As RECT

    ' Find the position of the window in question, in
    ' relation to its parent window (the Access desktop,
    ' the MDIClient window).
    hWndParent = GetParent(hWnd)

   ' Get the coordinates of the current window and its parent.
    Call GetWindowRect(hWnd, rct)

    ' Catch the case where the form is Popup (that is,
    ' its parent is NOT the Access main window.)  In that
    ' case, don't subtract off the coordinates of the
    ' Access MDIClient window.
    Select Case ReferenceType
        Case rtPopup
            ' No special calculations necessary.
        Case rtNormal
            ' Get the MDI Client window's coordinates.
            Call GetWindowRect(hWndParent, rctParent)
    
            ' Subtract off the left and top parent coordinates,
            ' since you need coordinates relative to the parent
            ' for the MoveWindow function call.
            ' Also, subtract off the small border of the
            ' MDI Client window that no one will admit to.
            With rct
                .left = .left - rctParent.left - adhcBorderWidthX
                .Top = .Top - rctParent.Top - adhcBorderWidthY
                .right = .right - rctParent.left - adhcBorderWidthX
                .bottom = .bottom - rctParent.Top - adhcBorderWidthY
            End With
    End Select
    cToFill = RctToCoords(rct)

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.GetCoords", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Private Sub HandleError(strName As String, _
 lngNumber As Long, strDescription As String)
    
#If DEBUGGING Then
    MsgBox "Error: " & strDescription & _
     " (" & lngNumber & ")", vbExclamation, strName
    ' Trigger a breakpoint. Remove this
    ' if you don't want a breakpoint here.
    Debug.Assert False
#End If
End Sub

Private Sub GetScreenInfo()
    ' This procedure fills in the module variables
    ' mptCurrentScreen, mptTwipsPerPixel
    
    ' From Access 2000 Developer's Handbook, Volume I
    ' by Getz, Litwin, and Gilbert (Sybex)
    ' Copyright 1999.  All rights reserved.
    
    On Error GoTo HandleErrors
    Dim lngDC As Long
    Dim ptCurrentDPI As POINTAPI
    Const HWND_DESKTOP = 0
    
    lngDC = GetDC(HWND_DESKTOP)
    
    ' If the call to GetDC didn't fail (and it had
    ' better not, or things are really busted),
    ' then get the info.
    
    If lngDC <> 0 Then
        ' Find the number of pixels in both directions
        ' on the screen, (640x480, 800x600, 1024x768,
        ' 1280x1024?). This also takes into account
        ' the size of the task bar, where ever it is.
        mptCurrentScreen.x = _
         GetSystemMetrics(SM_CXFULLSCREEN)
        mptCurrentScreen.Y = _
         GetSystemMetrics(SM_CYFULLSCREEN)
        
        ' Get the pixels/inch ratio, as well.
        ptCurrentDPI.x = GetDeviceCaps(lngDC, LOGPIXELSX)
        ptCurrentDPI.Y = GetDeviceCaps(lngDC, LOGPIXELSY)
        
        mptTwipsPerPixel.x = _
         adhcTwipsPerInch / ptCurrentDPI.x
        mptTwipsPerPixel.Y = _
         adhcTwipsPerInch / ptCurrentDPI.Y
    
        ' Release the information context.
        Call ReleaseDC(HWND_DESKTOP, lngDC)
    End If

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.GetScreenInfo", _
             Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub

Private Function GetIsPopup() As Boolean
    ' Is this form a popup form? To find that
    ' out, check the parent of the form. If
    ' it's Access, you know it's popup (otherwise,
    ' the parent is the MDI Client window.)

    Dim hWndParent As Long
    
    hWndParent = GetParent(frm.hWnd)
    GetIsPopup = (hWndParent = Application.hWndAccessApp)
End Function

Private Function GetIsSubform() As Boolean
    ' Is our form currently loaded as a subform?
    ' Check its Parent property to find out.
    
    Dim strName As String
    
    On Error Resume Next
    strName = frm.Parent.name
    GetIsSubform = (Err.Number = 0)
    Err.Clear
End Function

Private Function ReferenceType() As ReferenceType
    Dim rt As ReferenceType
    
    ' Need some easy way to know what to
    ' use as a reference point in all
    ' position calculations. This function
    ' wraps that decision-making up.
    
    If IsPopup Then
        rt = rtPopup
    Else
        rt = rtNormal
    End If
    ReferenceType = rt
End Function

Private Function RctToCoords(rct As RECT) As COORDS
    ' Convert from a RECT struct to a COORDS struct.
    ' This seems to come up often.
    Dim c As COORDS
    With c
        .left = rct.left
        .Top = rct.Top
        .Width = rct.right - rct.left
        .Height = rct.bottom - rct.Top
    End With
    RctToCoords = c
End Function

Private Sub GetClientOffsets()
    Dim p As POINTAPI
    Dim rct As RECT
    
    ' Get coords of upper-left corner of client area.
    ' Because Access coordinates are from the upper-left
    ' corner of the client area, but Windows coordinates
    ' measure from the upper-left corner of the main Access
    ' window, you need some way of calculating the offset
    ' of the MDI Client window. This procedure
    ' does that.
    
    ' Convert 0, 0 within the form's client area to
    ' absolute screen coordinates.
    p.x = 0
    p.Y = 0
    Call ClientToScreen(frm.hWnd, p)
    Call GetWindowRect(frm.hWnd, rct)
    
    mptClientOffset.x = p.x - rct.left
    mptClientOffset.Y = p.Y - rct.Top
End Sub

Private Sub GetParentCoords(hWnd As Long, c As COORDS)
    On Error GoTo HandleErrors
    Dim hWndParent As Long
    Dim rctParent As RECT
    
    Select Case ReferenceType
        Case rtNormal
            hWndParent = GetParent(hWnd)
            Call GetClientRect(hWndParent, rctParent)
        
        Case rtPopup
            ' Get the desktop coordinates.
            ' If you want to fill the ENTIRE Windows
            ' desktop, use this code.
            
            ' Call GetWindowRect(GetDesktopWindow(), rctParent)

            ' If you want to respect the task bar and other docked
            ' toolbars, use the following code, instead of the commented
            ' out code above.
            Call SystemParametersInfoRect( _
             SPI_GETWORKAREA, 0, rctParent, 0)
    End Select
    c = RctToCoords(rctParent)

ExitHere:
    Exit Sub

HandleErrors:
    Select Case Err.Number
        Case Else
            Call HandleError("FormInfo.GetParentCoords", Err.Number, Err.Description)
    End Select
    Resume ExitHere
End Sub