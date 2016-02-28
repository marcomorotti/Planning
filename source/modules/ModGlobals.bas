
Option Compare Database   'Use database order for string comparisons

' Common Form_Error codes
Public Const errCancel As Long = 2501
Public Const errCancel2 As Long = 2001
Public Const errDuplicate As Long = 3022
Public Const errInvalid As Long = 2113
Public Const errValidation As Long = 2116
Public Const errInputMask As Long = 2279
Public Const errRI As Long = 3200
Public Const errCustomValidate As Long = 3316
Public Const errTableValidate As Long = 3317
Public Const errSearchEnd As Long = 8504
Public Const errSpellCheck As Long = 9536
Public Const errGeneral As Long = 3316
Public Const errPropNotFound As Long = 3270

' Places to save current user info
Public gstrThisEmployee As String
Public glngThisEmployeeID As Long
Public glngThisDeptID As Long
Public gintIsManager As Integer
Public gintIsAdmin As Integer


' Il Direttorio del Help File
Public Const FormHelpFile As String = "C:\PortafoglioOrdini\HelpChm\portafoglioordini.chm"

' The release of this code
Public Const gTHISVERSION As Currency = 2.6

' String constant for application name
Public Const gstrAppTitle As String = "Portafoglio-Ordini"

' Places to save current user info
Public gstrThisUser As String


' ***********************************************************
' replace any single quotes ('), with 2 single quotes.
' replace any double quotes (") found, with 2 double quotes.

Function SQLize(strIn As String) As String
'   CHR(34) is a Double Quote("), CHR(39) is a Single Quote(')
    Dim Tmp As String
    Tmp = Replace(strIn, Chr(34), Chr(34) & Chr(34), , , vbBinaryCompare)
    Tmp = Replace(Tmp, "'", "''", , , vbBinaryCompare)
    SQLize = Tmp
End Function