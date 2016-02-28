Option Compare Database
'setup a new type to hold selection of control properties
Public Type CtrlProps
    name As String
    Wide As Integer
    High As Integer
    Topper As Integer
    Lefter As Integer
End Type

'possible variables for a "home" position for both form and controls on the form
'Public CtlHome() As CtrlProps
'Public FrmHomeX As Integer
'Public FrmHomeY As Integer
'
'Public Sub SetHomePosition(frm As Form)
''possible initial set to having a home position for shrinking forms and replacing controls while keeping them together
'Dim frmctl As Control
'ReDim CtlHome(0 To 0)
'i = 0
'For Each frmctl In frm.Controls
'    ReDim Preserve ctlSet(0 To i)
'    CtlHome(i).Name = frmctl.Name
'    CtlHome(i).High = frmctl.Height
'    CtlHome(i).Wide = frmctl.Width
'    CtlHome(i).Topper = frmctl.Top
'    CtlHome(i).Lefter = frmctl.Left
'    i = i + 1
'Next
'FrmHomeX = frm.Width
'FrmHomeY = frm.Detail.Height
'End Sub

Public Sub MoveCtls(frm As Form, ctl As control, Xinc As Integer, Yinc As Integer)
Dim obj As control

For Each obj In frm.Controls

    If obj.Top > (ctl.Top + ctl.Height + 1) Then
        obj.Top = obj.Top + Yinc
    End If
    If obj.left > (ctl.left + ctl.Width + 1) And obj.Top + obj.Height > ctl.Top And obj.Top < ctl.Top + ctl.Height Then
        obj.left = obj.left + Xinc
    End If
Next

End Sub

Public Sub ResizeCtl(frm As Form, ctl As control, iniX As Integer, iniY As Integer, speed As Integer)

Dim ctlX As Integer
Dim ctlY As Integer
Dim subfrmX As Integer
Dim subfrmY As Integer
Dim going As Boolean

ctlX = ctl.Width
ctlY = ctl.Height
subfrmX = ctl.Form.Width
' .detail se form americane
' If frm.Name = "frmCategoriaEventiCva" Then
If ctl.Form.name = "frmCategoriaEventiCva" Then
   subfrmY = 3000
Else
    subfrmY = ctl.Form.Corpo.Height
End If
going = True

If speed > 0 Then
    Do While going = True
        'first check if the control is still smaller than the subform size
        If ctlX < subfrmX Then
            'then check to see if there is space for it to move
            If WillFit(frm, ctl, speed, True, False) = "Yes" Then
                ctl.Width = ctl.Width + speed: ctlX = ctlX + speed
                GoTo NextX
            End If
            'if the space is not there, then we first need to check if we need to increase the form's size
            'and then we need to check if there is a control (or more) to the right that needs moved
            If InStr(1, WillFit(frm, ctl, speed, True, False), "FormX") > 0 Then Call ResizeFrm(frm, speed, 0)
            If InStr(1, WillFit(frm, ctl, speed, True, False), "ControlX") > 0 Then Call MoveCtls(frm, ctl, speed, 0)
            'after all of the moving is done, increase the control size by the speed
            ctl.Width = ctl.Width + speed: ctlX = ctlX + speed
        End If
NextX:
        'repeat from above for below the control
        If ctlY < subfrmY Then
            If WillFit(frm, ctl, speed, False, True) = "Yes" Then
                ctl.Height = ctl.Height + speed: ctlY = ctlY + speed
                GoTo NextY
            End If
            If InStr(1, WillFit(frm, ctl, speed, False, True), "FormY") > 0 Then Call ResizeFrm(frm, 0, speed)
            If InStr(1, WillFit(frm, ctl, speed, False, True), "ControlY") > 0 Then Call MoveCtls(frm, ctl, 0, speed)
            ctl.Height = ctl.Height + speed: ctlY = ctlY + speed
        End If
NextY:
        If ctlX >= subfrmX And ctlY >= subfrmY Then going = False
    Loop
Else
    Do While going = True
        'same shrinking is still in play.  Might be able to switch that to a "home" position for
        'form and controls, enabling them to stay together better as things can get kinda wonky
        'now depending on how the forms and buttons are arranged
        If ctlX > iniX Then Call ResizeFrm(frm, speed, 0): Call MoveCtls(frm, ctl, speed, 0): ctl.Width = ctl.Width + speed: ctlX = ctlX + speed
        If ctlY > iniY Then Call ResizeFrm(frm, 0, speed): Call MoveCtls(frm, ctl, 0, speed): ctl.Height = ctl.Height + speed: ctlY = ctlY + speed
        If ctlX <= iniX And ctlY <= iniY Then going = False
    Loop
End If
End Sub

Public Sub ResizeFrm(frm As Form, Xinc As Integer, Yinc As Integer)
frm.Width = frm.Width + Xinc  ' commented out 4/27 by Jeff N
frm.Detail.Height = frm.Detail.Height + Yinc
End Sub

Private Function WillFit(frm As Form, ctl As control, speed As Integer, Optional dirX As Boolean, Optional dirY As Boolean) As String
'function returns "ControlX","ControlY","FormY","FormX" or "Yes" multiples are separated with a ; ("ControlX;FormX")
'depending on what needs to be moved/resized
Dim continue As Boolean
Dim ctlSet() As CtrlProps
Dim Sizer As String
Dim frmctl As control

'if we are not changing either x or y directions, no need to see if it will fit, exit the function
If dirX = False And dirY = False Then Exit Function

'first check the form size to see if the change in size will fit
If dirX = True Then
    If ctl.Width + ctl.left + speed < frm.Width Then continue = True Else continue = False: If Nz(Sizer, "") = "" Then Sizer = "FormX" Else Sizer = Sizer & ";" & "FormX"
End If
If dirY = True Then
    If ctl.Height + ctl.Top + speed < frm.Detail.Height Then continue = True Else continue = False: If Nz(Sizer, "") = "" Then Sizer = "FormY" Else Sizer = Sizer & ";" & "FormY"
End If
'if it doesn't fit, send set it to Form and exit the function as there are no controls to the right
If continue = False Then WillFit = Sizer: Exit Function

'add all of the controls to the array that we created so that way we can
'more easily compare and address items we want
i = 0
ReDim ctlSet(0 To 0)
For Each frmctl In frm.Controls
    ReDim Preserve ctlSet(0 To i)
    ctlSet(i).name = frmctl.name
    ctlSet(i).High = frmctl.Height
    ctlSet(i).Wide = frmctl.Width
    ctlSet(i).Topper = frmctl.Top
    ctlSet(i).Lefter = frmctl.left
    i = i + 1
Next
'this section addresses the controls to the right and if they need to move
If dirX = True Then
    'if we are changing the x direction
    For i = 0 To UBound(ctlSet())
        'cycle thru all of the controls in our array
        'if the control being looked at is below the top of the control being resized then
        If ctlSet(i).Topper + ctlSet(i).High > ctl.Top Then
            'if the control being looked at is to the right of the control being resized AND the control being looked at is not the one being resized then
            If ctlSet(i).Lefter > ctl.left + ctl.Width And ctlSet(i).name <> ctl.name Then
                'if there is space for the control to be resized then do nothing, otherwise we need to go to the next step
                If ctlSet(i).Lefter >= ctl.left + ctl.Width + speed + 1 Then continue = True Else continue = False: GoTo FalserX
            End If
        End If
    Next i
FalserX:
    'if it will run into a control then we next need to see if moving the controls will need the form to be resized
    If continue = False Then
        'return that we do indeed need to move a control out of the way of the resizing
        If Nz(Sizer, "") = "" Then Sizer = "ControlX" Else Sizer = Sizer & ";" & "ControlX"
        'sort array so that way we can get the right most control
        Call OrderArray(ctlSet(), True)
        'check right most control for sizing
        'if the form needs to be resized because moving the control would be too much, then return that as a parameter also
        If ctlSet(UBound(ctlSet())).Lefter + ctlSet(UBound(ctlSet())).Wide + speed > frm.Width Then: If Nz(Sizer, "") = "" Then Sizer = "FormX" Else Sizer = Sizer & ";" & "FormX"
        WillFit = Sizer
        Exit Function
    End If

End If

'this addresses the controls to the bottom
'(same as above but for the Y direction)
If dirY = True Then
    For i = 0 To UBound(ctlSet())
        If ctlSet(i).Lefter + ctlSet(i).Wide > ctl.left Then
            If ctlSet(i).Topper > ctl.Top + ctl.Height And ctlSet(i).name <> ctl.name Then
                If ctlSet(i).Topper >= ctl.Top + ctl.Height + speed + 1 Then continue = True Else continue = False: GoTo FalserY
            End If
        End If
    Next i
FalserY:
    'if it will run into a control then we next need to see if moving the controls will need the form to be resized
    If continue = False Then
        If Nz(Sizer, "") = "" Then Sizer = "ControlY" Else Sizer = Sizer & ";" & "ControlY"
        'sort array for bottom most control
        Call OrderArray(ctlSet(), False)
        'check bottom most control for sizing
        If ctlSet(UBound(ctlSet())).Topper + ctlSet(UBound(ctlSet())).High + speed > frm.Detail.Height Then: If Nz(Sizer, "") = "" Then Sizer = "FormY" Else Sizer = Sizer & ";" & "FormY"
        WillFit = Sizer
        Exit Function
    
    End If
End If

Sizer = "Yes"
WillFit = Sizer
End Function

Private Function WillShrink(frm As Form, ctl As control, speed As Integer, Optional dirX As Boolean, Optional dirY As Boolean) As String

End Function
Private Function OrderArray(ByRef List() As CtrlProps, LeftOrTop As Boolean)
'Left is True, Top is False
Dim Temp As CtrlProps
If LeftOrTop = True Then
    For i = LBound(List()) To UBound(List()) - 1
        For j = i + 1 To UBound(List())
            If List(i).Lefter + List(i).Wide > List(j).Lefter + List(j).Wide Then
                Temp = List(j)
                List(j) = List(i)
                List(i) = Temp
            End If
        Next j
    Next i
Else
    For i = LBound(List()) To UBound(List()) - 1
        For j = i + 1 To UBound(List())
            If List(i).Topper + List(i).High > List(j).Topper + List(j).High Then
                Temp = List(j)
                List(j) = List(i)
                List(i) = Temp
            End If
        Next j
    Next i
End If

End Function