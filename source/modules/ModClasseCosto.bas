Option Compare Database

Public Function ClasseCosto(Cs_Csc) As Variant
Dim MyQuery As String
Dim Db As Database, rst As Recordset
Dim intCounter As Integer
Dim ClasseCostoT(1 To 21) As String
Dim At(1 To 21) As Variant
Set Db = CurrentDb
MyQuery = "SELECT tblCategorieCosto.ClasseCosto, tblCategorieCosto.Da, tblCategorieCosto.A, " & _
            "tblCategorieCosto.[MesiCopertura] " & _
            "FROM tblCategorieCosto " & _
            "ORDER BY tblCategorieCosto.A"
Set rst = Db.OpenRecordset(MyQuery, dbOpenDynaset)
rst.MoveFirst
For intCounter = 1 To 21
        ClasseCostoT(intCounter) = rst!ClasseCosto
        At(intCounter) = rst!A
Select Case Cs_Csc
    Case Is <= At(intCounter)
          ClasseCosto = ClasseCostoT(intCounter)
          rst.Close
          Exit Function
    Case Else
         ClasseCosto = ""
    End Select
rst.MoveNext
Next intCounter
rst.Close

End Function