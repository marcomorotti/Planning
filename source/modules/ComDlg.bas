Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This source code is copyrighted material which
'may not be published in any form without explicit prior permission
'from the author. However, you are free to use the source code
'in your private, non-commercial, projects without permission.
'You are allowed to use these functions and routines in commercial
'products, provided the documentation of these products makes a
'reference to the original source. The following reference is recommended:
'-----------------------------------------------------------
'(PART OF) THIS SOFTWARE IS BASED ON SOURCE CODE, ORIGINALLY CREATED
'BY ROMKE SOLDAAT (ROMKE@SOLDAAT.COM), AND PUBLISHED IN MICROSOFT
'OFFICE & VISUAL BASIC FOR APPLICATIONS DEVELOPER, BY INFORMANT
'COMMUNICATIONS GROUP (WWW.INFORMANT.COM)
'-----------------------------------------------------------
' Permission to use in Microsoft Office Access 2003 Inside Out
' granted by Romke Soldaat, January 15, 2003
'===========================================================

Option Compare Database
Option Explicit
 
DefStr S
DefLng N
DefBool B
DefVar V
 
' OFN constants.
Const OFN_ALLOWMULTISELECT   As Long = &H200
Const OFN_CREATEPROMPT       As Long = &H2000
Const OFN_EXPLORER           As Long = &H80000
Const OFN_EXTENSIONDIFFERENT As Long = &H400
Const OFN_FILEMUSTEXIST      As Long = &H1000
Const OFN_HIDEREADONLY       As Long = &H4
Const OFN_LONGNAMES          As Long = &H200000
Const OFN_NOCHANGEDIR        As Long = &H8
Const OFN_NODEREFERENCELINKS As Long = &H100000
Const OFN_OVERWRITEPROMPT    As Long = &H2
Const OFN_PATHMUSTEXIST      As Long = &H800
Const OFN_READONLY           As Long = &H1
 
' The maximum length of a single file path.
Const MAX_PATH As Long = 260
' This MAX_BUFFER value allows you to select approx.
' 500 files with an average length of 25 characters.
' Change this value as needed.
Const MAX_BUFFER As Long = 50 * MAX_PATH
' String constants:
Const sBackSlash As String = "\"
Const sPipe As String = "|"
 
' API functions to use the Windows common dialog boxes.
Private Declare Function GetOpenFileName _
  Lib "comdlg32.dll" Alias "GetOpenFileNameA" _
  (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName _
  Lib "comdlg32.dll" Alias "GetSaveFileNameA" _
  (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetActiveWindow _
  Lib "user32" () As Long
 
' Type declaration, used by GetOpenFileName and
' GetSaveFileName.
Private Type OPENFILENAME
  lStructSize       As Long
  hwndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  Flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String  ' Can also be a Long.
End Type
 
' Private variables.
Private OFN As OPENFILENAME
Private colFileTitles As New Collection
Private colFileNames As New Collection
Private sFullName
Private sFileTitle
Private sPath
Private sExtension

'public enumeration variable:
Public Enum XFlags
  PathMustExist = OFN_PATHMUSTEXIST
  FileMustExist = OFN_FILEMUSTEXIST
  PromptToCreateFile = OFN_CREATEPROMPT
End Enum
 
Property Let AllowMultiSelect(bFlag)
  SetFlag OFN_ALLOWMULTISELECT, bFlag
End Property
 
Property Let DialogTitle(sCaption)
  OFN.lpstrTitle = sCaption
End Property
 
Property Let Filter(vFilter)
  If IsArray(vFilter) Then _
    vFilter = Join(vFilter, vbNullChar)
    OFN.lpstrFilter = _
      Replace(vFilter, sPipe, vbNullChar) & vbNullChar
End Property
 
Property Get Filter()
  With OFN
    If .nFilterIndex Then
      Dim sTemp As Variant
      sTemp = Split(.lpstrFilter, vbNullChar)
      Filter = sTemp(.nFilterIndex * 2 - 2) & sPipe & _
        sTemp(.nFilterIndex * 2 - 1)
    End If
  End With
End Property
 
Property Let FilterIndex(nIndex)
  OFN.nFilterIndex = nIndex
End Property
 
Property Get FilterIndex() As Long
  FilterIndex = OFN.nFilterIndex
End Property
 
Property Let RestoreCurDir(bFlag)
  SetFlag OFN_NOCHANGEDIR, bFlag
End Property
 
Property Let ExistFlags(nFlags As XFlags)
  OFN.Flags = OFN.Flags Or nFlags
End Property
 
Property Let CheckBoxVisible(bFlag)
  SetFlag OFN_HIDEREADONLY, Not bFlag
End Property
 
Property Let CheckBoxSelected(bFlag)
  SetFlag OFN_READONLY, bFlag
End Property
 
Property Get CheckBoxSelected() As Boolean
  CheckBoxSelected = OFN.Flags And OFN_READONLY
End Property
 
Property Let fileName(sFileName)
  If Len(sFileName) <= MAX_PATH Then _
    OFN.lpstrFile = sFileName
End Property
 
Property Get fileName() As String
  fileName = sFullName
End Property
 
Property Get FileNames() As Collection
  Set FileNames = colFileNames
End Property
 
Property Get FileTitle() As String
  FileTitle = sFileTitle
End Property
 
Property Get FileTitles() As Collection
  Set FileTitles = colFileTitles
End Property
 
Property Let directory(sInitDir)
  OFN.lpstrInitialDir = sInitDir
End Property
 
Property Get directory() As String
  directory = sPath
End Property
 
Property Let Extension(sDefExt)
  OFN.lpstrDefExt = LCase$(left$( _
    Replace(sDefExt, ".", vbNullString), 3))
End Property
 
Property Get Extension() As String
  Extension = sExtension
End Property
 
Function ShowOpen() As Boolean
  ShowOpen = Show(True)
End Function
 
Function ShowSave() As Boolean
  ' Set or clear appropriate flags for Save As dialog.
  SetFlag OFN_ALLOWMULTISELECT, False
  SetFlag OFN_PATHMUSTEXIST, True
  SetFlag OFN_OVERWRITEPROMPT, True
  ShowSave = Show(False)
End Function
 
Private Function Show(bOpen)
  With OFN
    .lStructSize = Len(OFN)
    ' Could be zero if no owner is required.
    .hwndOwner = GetActiveWindow
    ' If the RO checkbox must be checked, we should also
    ' display it.
    If .Flags And OFN_READONLY Then _
      SetFlag OFN_HIDEREADONLY, False
    ' Create large buffer if multiple file selection
    ' is allowed.
    .nMaxFile = IIf(.Flags And OFN_ALLOWMULTISELECT, _
      MAX_BUFFER + 1, MAX_PATH + 1)
    .nMaxFileTitle = MAX_PATH + 1
    ' Initialize the buffers.
    .lpstrFile = .lpstrFile & String$( _
      .nMaxFile - 1 - Len(.lpstrFile), 0)
    .lpstrFileTitle = String$(.nMaxFileTitle - 1, 0)
 
    ' Display the appropriate dialog.
    If bOpen Then
      Show = GetOpenFileName(OFN)
    Else
      Show = GetSaveFileName(OFN)
    End If
 
    If Show Then
      ' Remove trailing null characters.
      Dim nDoubleNullPos
      nDoubleNullPos = InStr(.lpstrFile & vbNullChar, _
                              String$(2, 0))
      If nDoubleNullPos Then
        ' Get the file name including the path name.
        sFullName = left$(.lpstrFile, nDoubleNullPos - 1)
        ' Get the file name without the path name.
        sFileTitle = left$(.lpstrFileTitle, _
          InStr(.lpstrFileTitle, vbNullChar) - 1)
        ' Get the path name.
        sPath = left$(sFullName, .nFileOffset - 1)
        ' Get the extension.
        If .nFileExtension Then
          sExtension = Mid$(sFullName, .nFileExtension + 1)
        End If
        ' If sFileTitle is a string,
        ' we have a single selection.
        If Len(sFileTitle) Then
          ' Add to the collections.
          colFileTitles.Add _
            Mid$(sFullName, .nFileOffset + 1)
          colFileNames.Add sFullName
        Else  ' Tear multiple selection apart.
          Dim sTemp As Variant, nCount
          sTemp = Split(sFullName, vbNullChar)
          ' If array contains no elements,
          ' UBound returns -1.
          If UBound(sTemp) > LBound(sTemp) Then
            ' We have more than one array element!
            ' Remove backslash if sPath is the root folder.
            If Len(sPath) = 3 Then _
              sPath = left$(sPath, 2)
            ' Loop through the array, and create the
            ' collections; skip the first element
            ' (containing the path name), so start the
            ' counter at 1, not at 0.
            For nCount = 1 To UBound(sTemp)
              colFileTitles.Add sTemp(nCount)
              ' If the string already contains a backslash,
              ' the user must have selected a shortcut
              ' file, so we don't add the path.
              colFileNames.Add IIf(InStr(sTemp(nCount), _
                sBackSlash), sTemp(nCount), _
                sPath & sBackSlash & sTemp(nCount))
            Next
            ' Clear this variable.
            sFullName = vbNullString
          End If
        End If
        ' Add backslash if sPath is the root folder.
        If Len(sPath) = 2 Then _
          sPath = sPath & sBackSlash
      End If
    End If
  End With
End Function
 
Private Sub SetFlag(nValue, bTrue)
  ' Wrapper routine to set or clear bit flags.
  With OFN
    If bTrue Then
      .Flags = .Flags Or nValue
    Else
      .Flags = .Flags And Not nValue
    End If
  End With
End Sub
 
Private Sub Class_Initialize()
  ' This routine runs when the object is created.
  OFN.Flags = OFN.Flags Or OFN_EXPLORER Or _
              OFN_LONGNAMES Or OFN_HIDEREADONLY
End Sub