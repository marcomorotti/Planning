Option Compare Database
Option Explicit

Sub CompactDatabaseX2()

   Dim dbsMezzi As Database
   Dim prpLoop As Property

   Set dbsMezzi = OpenDatabase("PortafoglioOrdini.accdb")

   ' Show the properties of the original database nella finestra immediata
   With dbsMezzi
      Debug.Print .name & ", version " & .Version
      Debug.Print "  CollatingOrder = " & .CollatingOrder
      .Close
   End With

  

End Sub