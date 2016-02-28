Option Compare Database

'***************** Code Start ***************
'This code was originally written by Dev Ashish.
'It is not to be altered or distributed,
'except as part of an application.
'You are free to use it in any application,
'provided the copyright notice is left unchanged.
'
'Code Courtesy of
'Dev Ashish
' Per Testarlo ?fIsAppRunning("Word",True)
Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_NORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_MAXIMIZE = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5
Private Const SW_MINIMIZE = 6
Private Const SW_SHOWMINNOACTIVE = 7
Private Const SW_SHOWNA = 8
Private Const SW_RESTORE = 9
Private Const SW_SHOWDEFAULT = 10
Private Const SW_MAX = 10

Private Declare Function apiFindWindow Lib "user32" Alias _
    "FindWindowA" (ByVal strClass As String, _
    ByVal lpWindow As String) As Long

Private Declare Function apiSendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal _
    wParam As Long, lParam As Long) As Long
    
Private Declare Function apiSetForegroundWindow Lib "user32" Alias _
    "SetForegroundWindow" (ByVal hWnd As Long) As Long
    
Private Declare Function apiShowWindow Lib "user32" Alias _
    "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
    
Private Declare Function apiIsIconic Lib "user32" Alias _
    "IsIconic" (ByVal hWnd As Long) As Long
    
Function fIsAppRunning(ByVal strAppName As String, _
        Optional fActivate As Boolean) As Boolean
    Dim lngH As Long, strClassName As String
    Dim lngX As Long, lngTmp As Long
    Const WM_USER = 1024
    On Local Error GoTo fIsAppRunning_Err
    fIsAppRunning = False
    Select Case LCase$(strAppName)
        Case "excel":       strClassName = "XLMain"
        Case "word":        strClassName = "OpusApp"
        Case "access":      strClassName = "OMain"
        Case "powerpoint95": strClassName = "PP7FrameClass"
        Case "powerpoint97": strClassName = "PP97FrameClass"
        Case "notepad":     strClassName = "NOTEPAD"
        Case "paintbrush":  strClassName = "pbParent"
        Case "wordpad":     strClassName = "WordPadClass"
        Case Else:          strClassName = vbNullString
    End Select
    
    If strClassName = "" Then
        lngH = apiFindWindow(vbNullString, strAppName)
    Else
        lngH = apiFindWindow(strClassName, vbNullString)
    End If
    If lngH <> 0 Then
        apiSendMessage lngH, WM_USER + 18, 0, 0
        lngX = apiIsIconic(lngH)
        If lngX <> 0 Then
            lngTmp = apiShowWindow(lngH, SW_SHOWNORMAL)
        End If
        If fActivate Then
            lngTmp = apiSetForegroundWindow(lngH)
        End If
        fIsAppRunning = True
    End If
fIsAppRunning_Exit:
    Exit Function
fIsAppRunning_Err:
    fIsAppRunning = False
    Resume fIsAppRunning_Exit
End Function
'******************** Code End ****************