Option Compare Database
Option Explicit

Private Const mconMeterForm = "frmStatusMeter"

Private Function IsOpen(strForm As String)
    IsOpen = (SysCmd(acSysCmdGetObjectState, acForm, strForm) > 0)
End Function

Public Sub acbCloseMeter()

    On Error GoTo HandleErr

    DoCmd.Close acForm, mconMeterForm

ExitHere:
    Exit Sub
HandleErr:
    
        MsgBox "Error#" & Err.Number & ": " & Err.Description, _
         , "acbCloseMeter"
    
    Resume ExitHere
End Sub

Public Sub acbInitMeter(strTitle As String, fIncludeCancel As Boolean)

    ' Initializes the status meter to 0.
    '
    ' In:
    '     strTitle - Title of status meter form

    On Error GoTo HandleErr

    DoCmd.OpenForm mconMeterForm
    Forms(mconMeterForm).InitMeter(fIncludeCancel) = strTitle

ExitHere:
    Exit Sub
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "Error#" & Err.Number & ": " & Err.Description, _
         , "acbInitMeter"
    End Select
    If IsOpen(mconMeterForm) Then Call acbCloseMeter
    Resume ExitHere
    Resume
End Sub

Public Function acbUpdateMeter(intValue As Integer) As Boolean

    ' Updates the status meter and returns whether
    ' the Cancel button was pressed.
    '
    ' In:
    '     intValue - percentage value 0-100

    On Error GoTo HandleErr

    Forms(mconMeterForm).UpdateMeter = intValue

    ' Return value is False if cancelled.
    If Forms(mconMeterForm).Cancelled Then
        Call acbCloseMeter
        acbUpdateMeter = False
    Else
        acbUpdateMeter = True
    End If

ExitHere:
    Exit Function
HandleErr:
    Select Case Err.Number
    Case Else
        MsgBox "Error#" & Err.Number & ": " & Err.Description, _
         , "acbUpdateMeter"
    End Select
    If IsOpen(mconMeterForm) Then Call acbCloseMeter
    Resume ExitHere
End Function