
Option Compare Database
Option Explicit

Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
        (ByVal hwndCaller As Long, ByVal pszFile As String, _
         ByVal uCommand As Long, ByVal dwData As Long) As Long

Const HH_DISPLAY_TOPIC = &H0
Const HH_SET_WIN_TYPE = &H4
Const HH_GET_WIN_TYPE = &H5
Const HH_GET_WIN_HANDLE = &H6
Const HH_DISPLAY_TEXT_POPUP = &HE
Const HH_HELP_CONTEXT = &HF
Const HH_TP_HELP_CONTEXTMENU = &H10
Const HH_TP_HELP_WM_HELP = &H11

Public Sub Show_Help(HelpFileName As String, MycontextID As Long)
    'A specific topic identified by the variable context-ID is started in
    'response to this button click.
    Dim hwndHelp As Long

    'The return value is the window handle of the created Help window.
    Select Case MycontextID
        Case Is = 0
            hwndHelp = HtmlHelp(Application.hWndAccessApp, HelpFileName, _
                       HH_DISPLAY_TOPIC, MycontextID)
        Case Else
            hwndHelp = HtmlHelp(Application.hWndAccessApp, HelpFileName, _
                       HH_HELP_CONTEXT, MycontextID)
    End Select
End Sub

Public Function HelpEntry()
    ' L' Help File è nel ModGlobals
    'Identify the name of the context-id.
    Dim FormHelpId As Long
    ' La Variabile FormHelpFile è Salvata nel modGlobals
    Dim curForm As Form
    ' HelpContextId = 47

    'Set the curForm variable to the currently active form.
    Set curForm = Screen.ActiveForm

    'Check the Help file property of the form. If a Help file exists,
    'assign the name and context-id to the respective variables.
    If FormHelpId = 0 Then
       FormHelpId = curForm.HelpContextId
    End If
    Show_Help FormHelpFile, FormHelpId
End Function