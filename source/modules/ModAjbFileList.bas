Option Compare Database
Public Function ListFiles(strPath As String, Optional strFileSpec As String, _
    Optional bIncludeSubfolders As Boolean, Optional lst As ListBox)
On Error GoTo Err_Handler
'Purpose:   List the files in the path.
    'Arguments: strPath = the path to search.
    '           strFileSpec = "*.*" unless you specify differently.
    '           bIncludeSubfolders: If True, returns results from subdirectories of strPath as well.
    '           lst: if you pass in a list box, items are added to it. If not, files are listed to immediate window.
    '               The list box must have its Row Source Type property set to Value List.
    'Method:    FilDir() adds items to a collection, calling itself recursively for subfolders.
    Dim colDirList As New Collection
    Dim varItem As Variant
    
    Call FillDir(colDirList, strPath, strFileSpec, bIncludeSubfolders)

' /// Use
'In the Immediate window
'
'To list the files in C:\Data, open the Immediate Window (Ctrl+G), and enter:
'    Call ListFiles("C:\Data")
'
'To limit the results to zip files:
'    Call ListFiles("C:\Data", "*.zip")
'
'To include files in subdirectories as well:
'    Call ListFiles("C:\Data", , True)
'In a list box
'
'To show the files in a list box:
'
'   1. Create a new form.
'   2. Add a list box, and set these properties:
'          Name              lstFileList
'          Row Source Type   Value List
'   3. Set the On Load property of the form to:
'          [Event Procedure]
'   4. Click the Build button (...) beside this. Access opens the code window. Set up the event procedure like this:
'
'          Private Sub Form_Load()
'              Call ListFiles("C:\Data", , , Me.lstFileList)
'          End Sub

' \\\ Fine Use

    'Add the files to a list box if one was passed in. Otherwise list to the Immediate Window.
    If lst Is Nothing Then
        For Each varItem In colDirList
            Debug.Print varItem
        Next
    Else
        For Each varItem In colDirList
        lst.AddItem varItem
        Next
    End If

Exit_Handler:
    Exit Function

Err_Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Resume Exit_Handler
End Function

Private Function FillDir(colDirList As Collection, ByVal strFolder As String, strFileSpec As String, _
    bIncludeSubfolders As Boolean)
    'Build up a list of files, and then add add to this list, any additional folders
    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant

    'Add the files to the folder.
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & strFileSpec)
    Do While strTemp <> vbNullString
        colDirList.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Build collection of additional subfolders.
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0& Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop
        'Call function recursively for each subfolder.
        For Each vFolderName In colFolders
            Call FillDir(colDirList, strFolder & TrailingSlash(vFolderName), strFileSpec, True)
        Next vFolderName
    End If
End Function

Public Function TrailingSlash(varIn As Variant) As String
    If Len(varIn) > 0& Then
        If right(varIn, 1&) = "\" Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & "\"
        End If
    End If
End Function