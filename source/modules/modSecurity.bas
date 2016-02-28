Option Compare Database
Option Explicit

Public Sub SetupUsers()
Dim Db As DAO.Database, rst As DAO.Recordset
Dim ws As DAO.Workspace, usr As DAO.User, grp As DAO.Group

    ' Use this sub to set up users and groups in the current workgroup
    ' Code reads through tblEmployees and creates users specified
    '  in the UserID field.
    ' Also creates the AppAdmin, DeptMgrs, and Employees groups.
    
    ' Open a recordset on qryEmployeeSignon that contains all the data we need
    Set Db = CurrentDb
    Set rst = Db.OpenRecordset("qryEmployeeSignon")
    
    ' Set a local error trap in case you've alread run this
    On Error Resume Next
    
    ' Call the undo procedure first to be sure we're starting clean
    Call UndoUsersAndGroups
    
    ' Now, create the new groups
    Set ws = DBEngine(0)
    Set grp = ws.CreateGroup("AppAdmin", "9999")
    ws.Groups.Append grp
    Set grp = ws.CreateGroup("DeptMgrs", "9999")
    ws.Groups.Append grp
    Set grp = ws.CreateGroup("Employees", "9999")
    ws.Groups.Append grp
    ws.Groups.Refresh
    
    ' Now read all employee records, build users, and set groups
    Do Until rst.EOF
        ' Create user
        Set usr = ws.CreateUser(rst!UserID, "9999", rst!Password)
        ws.Users.Append usr
        ' Add to Users and Employees groups
        Set grp = usr.CreateGroup("Users")
        usr.Groups.Append grp
        Set grp = usr.CreateGroup("Employees")
        usr.Groups.Append grp
        ' Check for IsAdmin
        If (rst!IsAdmin = True) Then
            ' Add to AppAdmin
            Set grp = usr.CreateGroup("AppAdmin")
            usr.Groups.Append grp
            ' And Admins
            Set grp = usr.CreateGroup("Admins")
            usr.Groups.Append grp
            ' If also dept manager,
            If rst!EmployeeNumber = rst!ManagerNumber Then
                ' Add to that group, too
                Set grp = usr.CreateGroup("DeptMgrs")
                usr.Groups.Append grp
            End If
        ' Not admin - is department manager?
        ElseIf rst!EmployeeNumber = rst!ManagerNumber Then
            ' Add to Admins and DeptMgrs
            Set grp = usr.CreateGroup("Admins")
            usr.Groups.Append grp
            Set grp = usr.CreateGroup("DeptMgrs")
            usr.Groups.Append grp
        End If
        rst.MoveNext
    Loop
    ws.Users.Refresh
    ws.Groups.Refresh
    
    rst.Close
    Set rst = Nothing
    Set grp = Nothing
    Set usr = Nothing
    Set ws = Nothing
    Set Db = Nothing
    
End Sub

Public Sub UndoUsersAndGroups()
' Run this procedure to remove the custom users and groups
Dim ws As DAO.Workspace, intI As Integer, strUser As String

    ' Set a local error trap just in case you've already run this
    On Error Resume Next
    ' First, delete all the special groups
    Set ws = DBEngine(0)
    ws.Groups.Delete "AppAdmin"
    ws.Groups.Delete "DeptMgrs"
    ws.Groups.Delete "Employees"
    
    ' Now loop and delete all users except the special ones
    ' Must step backwards through the users
    For intI = ws.Users.Count - 1 To 0 Step -1
        strUser = ws.Users(intI).name
        If strUser = "Admin" Or strUser = "Engine" Or _
            strUser = "Creator" Or strUser = "HousingAdmin" Then
        ' Don't delete these sytem IDs
        Else
            ' Got one - delete it
            ws.Users.Delete strUser
        End If
    Next intI
    ws.Users.Refresh
    ws.Groups.Refresh
    Set ws = Nothing
End Sub