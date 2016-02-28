' REV  DATE                          DESCRIPTION
' 1.0  2013-08-26              Release iniziale

'
'==============================================================================
' NOME: GetSQLRecordcount
' Scopo: restituisce il recordcount di una non-action query
' RETURNS: Long, numero di record di una query, 0 se errore
' ARGUMENTS: strSQL: SQL Statement to evaluate
' USAGE: first word must be "SELECT" (case-insensitive)
'        current handling for CurrentDb only
'       GetSQLRecordcount("SELECT * FROM tblArticoli")
' DEPENDANCIES: DAO
'
' REVISIONS:
'  REV |    DATE    | REV TYPE | DESCRIPTION
'------------------------------------------------------------------------------
'  R01   8/16/2010    INITIAL
'
'==============================================================================
'ErrHandler V3.01
Public Function GetSQLRecordcount(strSQL As String) As Long
On Error GoTo Error_Proc
Dim Ret As Long
'=========================
 Const ERRN_SQL_NOTSELECT = 8000 + vbObjectError
 Const ERRM_SQL_NOTSELECT As String = _
   "The SQL is not a Select statement."

 Dim rs As DAO.Recordset
'=========================

 'si assicura che la frase inizi con "SELECT
 If StrComp(left(LTrim(strSQL), 6), "SELECT", vbTextCompare) <> 0 Then
   'raise error
   Err.Raise ERRN_SQL_NOTSELECT, , ERRM_SQL_NOTSELECT
 End If

 Set rs = CurrentDb.OpenRecordset("SELECT Count(*) FROM (" & strSQL & ") As vTbl")
 
 If rs.EOF Then
   Ret = 0
 Else
   Ret = rs.Fields(0)
 End If
 
 rs.Close
 
'=========================
Exit_Proc:
 Set rs = Nothing
 GetSQLRecordcount = Ret
 Exit Function
Error_Proc:
 Select Case Err.Number
   Case Else
     MsgBox "Error: " & Trim(Str(Err.Number)) & vbCrLf & _
       "Desc: " & Err.Description & vbCrLf & vbCrLf & _
       "Module: modSQLUtil, Procedure: GetSQLRecordcount" _
       , vbCritical, "Error!"
 End Select
 Resume Exit_Proc
 Resume
End Function