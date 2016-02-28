Option Compare Database
Option Explicit

Function ApriMaschere( _
  FormName As Variant, _
  Optional View As AcFormView = acNormal, _
  Optional FilterName As Variant, _
  Optional WhereCondition As Variant, _
  Optional DataMode As AcFormOpenDataMode = acFormPropertySettings, _
  Optional WindowMode As AcWindowMode = acWindowNormal, _
  Optional OpenArgs As Variant _
)
' Funzione utilizzata nell'evento Clic del pulsante di comando per
' aprire le maschere dal Pannello comandi principale. L'uso della funzione
' permette di evitare la ripetizione del codice nelle routine evento.
' Es. =ApriMaschere("frmNewTipoVeicolo")

  If CurrentProject.AllForms(FormName).IsLoaded Then
    DoCmd.Close acForm, FormName
  End If
' Apre la maschera specifica.
  DoCmd.OpenForm FormName, _
                 View, _
                 FilterName, _
                 WhereCondition, _
                 DataMode, _
                 WindowMode, _
                 OpenArgs
On Error GoTo Err_ApriMaschere

Esci_ApriMaschere:
    Exit Function

Err_ApriMaschere:
    MsgBox Err.Description
    Resume Esci_ApriMaschere

End Function