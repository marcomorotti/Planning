Option Compare Database
Option Explicit
Public formOpen As Variant

 Function IsLoaded(strFrmName As String) As Boolean
    
    '  Determines if a form is loaded.
    
    Const conFormDesign = 0
    Dim intX As Integer
    
    IsLoaded = False
    For intX = 0 To Forms.Count - 1
        If Forms(intX).FormName = strFrmName Then
            If Forms(intX).CurrentView <> conFormDesign Then
                IsLoaded = True
                Exit Function  ' Quit function once form has been found.
            End If
        End If
    Next

End Function