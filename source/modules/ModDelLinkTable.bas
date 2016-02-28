Option Compare Database

'---------------------------------------------------------------------------------------
' Procedura : DeleteAttachedTbls
' Author    : Marco Morotti
' Purpose   : Deletes all the linked tables from the active database.  It only removes
'             the links and does not actually delete the tables from the back-end
'             database.
' Copyright : Marco Morotti
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2011-Jun-10                 Initial Release
'---------------------------------------------------------------------------------------
Function DeleteAttachedTbls()
On Error GoTo Error_Handler
    Dim Db As DAO.Database
    Dim tdf As DAO.TableDef
 
    DoCmd.SetWarnings False
    Set Db = CurrentDb()
    For Each tdf In Db.TableDefs
        If (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
            DoCmd.DeleteObject acTable, tdf.name
        End If
    Next
 
Error_Handler_Exit:
    DoCmd.SetWarnings True
    Set tdf = Nothing
    Set Db = Nothing
    Exit Function
 
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Error Source: DeleteAttachedTbls" & vbCrLf & _
           "Error Description: " & Err.Description, _
           vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function