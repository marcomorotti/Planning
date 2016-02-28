Function fExistTable(strTableName As String) As Integer
Dim Db As Database
Dim i As Integer
    Set Db = DBEngine.Workspaces(0).Databases(0)
    fExistTable = False
    Db.TableDefs.Refresh
    For i = 0 To Db.TableDefs.Count - 1
        If strTableName = Db.TableDefs(i).name Then
            'Table Exists
            fExistTable = True
            Exit For
        End If
    Next i
    Set Db = Nothing
End Function