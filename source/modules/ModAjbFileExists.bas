Option Compare Database

Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long
'/// Use
'Examples
'
'   1. Look for a file named MyFile.mdb in the Data folder:
'          FileExists ("C:\Data\MyFile.mdb")
'   2. Look for a folder named System in the Windows folder on C: drive:
'          FolderExists ("C:\Windows\System")
'   3. Look for a file named MyFile.txt on a network server:
'          FileExists ("\\MyServer\MyPath\MyFile.txt")
'   4. Check for a file or folder name Wotsit on the server:
'          FileExists("\\MyServer\Wotsit", True)
'   5. Check the folder of the current database for a file named GetThis.xls:
'          FileExists (TrailingSlash(CurrentProject.Path) & "GetThis.xls")
'\\\ Fine Use

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While right$(strFile, 1) = "\"
            strFile = left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Function FolderExists(strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function
'Function TrailingSlash(varIn As Variant) As String
'    If Len(varIn) > 0 Then
'        If right(varIn, 1) = "\" Then
'            TrailingSlash = varIn
'        Else
'            TrailingSlash = varIn & "\"
'        End If
'    End If
'End Function