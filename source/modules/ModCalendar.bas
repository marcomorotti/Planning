Option Compare Database
Option Explicit

' Place holder for the form class
Dim frmCalOCX As Form_frmCalendarOCX

Public Function GetDate(ctl As control, Optional intDateOnly As Integer = 0) As Integer
'-----------------------------------------------------------
' Inputs: A Control object containing a date/time value
'         Optional "date only" (no time value) flag
' Outputs: Sets the Control to the value returned by frmCalendar
' Created By: JLV 09/05/01
' Last Revised: JLV 09/05/01
'-----------------------------------------------------------
Dim varDateTime As Variant, strDateTime As String, frm As Form
    ' Error trap just in case
    On Error GoTo Error_Date
    
    ' First, validate the kind of control passed
    Select Case ctl.ControlType
        ' Text box, combo box, and list box are OK
        Case acTextBox, acListBox, acComboBox
        Case Else
            GetDate = False
            Exit Function
    End Select
    
    ' If the control has no value
    If IsNothing(ctl.Value) Then
        If intDateOnly Then
            ' Set default date
            varDateTime = Date
        Else
            ' .. or default date and time
            varDateTime = Now
        End If
    Else
        ' Otherwise, pick up the current value
        varDateTime = ctl.Value
        ' Make sure it's a date/time
        If vbDate <> VarType(varDateTime) Then
            GetDate = False
            Exit Function
        End If
    End If
    ' Turn the date and time into a string to pass to the form
    strDateTime = Format(varDateTime, "dd/mm/yyyy hh:nn")
    ' Make sure we don't have an old copy of frmCalendar hanging around
    If IsFormLoaded("frmCalendar") Then DoCmd.Close acForm, "frmCalendar"
    ' Open the calendar as a dialog so this code waits, and pass the date/time value
    DoCmd.OpenForm "frmCalendar", WindowMode:=acDialog, OpenArgs:=strDateTime & "," & intDateOnly
    ' If the form is gone, user canceled the update
    If Not IsFormLoaded("frmCalendar") Then Exit Function
    ' Get a pointer to the now-hidden form
    Set frm = Forms!frmCalendar
    ' Grab the date part off the hidden text box
    strDateTime = Format(frm.ctlCalendar.Value, "dd/mm/yyyy")
    If Not intDateOnly Then
        ' If looking for date and time, also grab the hour and minute
        strDateTime = strDateTime & " " & frm.txtHour & ":" & frm.txtMinute
    End If
    ' Stuff the returned value back in the caller's control
    ctl.Value = DateValue(strDateTime) + TimeValue(strDateTime)
    ' Close the calendar form to clean up
    DoCmd.Close acForm, "frmCalendar"
    GetDate = True

Exit_Date:
    Exit Function
    
Error_Date:
    ' This code is pretty simple and does check for a usable control type,
    '  .. so this should never happen.
    ' But if it does, log it...
    ErrorLog "GetDate", Err, Error
    GetDate = False
    Resume Exit_Date
    
End Function

Function GetDateOCX(ctlToUpdate As control, Optional intDateOnly As Integer = 0)
'-----------------------------------------------------------
' Inputs: A Control object containing a date/time value
'         Optional "date only" (no time value) flag
' Outputs: Sets the Control to the value returned by frmCalendar
' Created By: JLV 11/15/02
' Last Revised: JLV 11/15/02
'-----------------------------------------------------------

' Set an error trap
On Error GoTo ProcErr

    ' Open the OCX calendar form by setting a new object
    ' NOTE: Uses a module variable in the Declarations section
    '       so that the form doesn't go away when this code exits
    Set frmCalOCX = New Form_frmCalendarOCX
    ' Call the calendar form's public method to
    '  pass it the control to update and the "date only" flag
    Set frmCalOCX.ctlToUpdate(intDateOnly) = ctlToUpdate
    ' Put the focus on the OCX calendar form
    frmCalOCX.SetFocus

ProcExit:
    ' Done
    Exit Function

ProcErr:
    MsgBox "An error has occurred in GetDateOCX.  " _
        & "Error number " & Err.Number & ": " & Err.Description _
        & vbCrLf & vbCrLf & "If this problem persists, note the error message and " _
        & "call your programmer.", , "Ooops . . .       (unexpected error)"
    Resume ProcExit
End Function