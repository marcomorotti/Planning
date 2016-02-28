Option Compare Database

'*************************************************************
'Declarations section of the module.
'*************************************************************
' Per testare :
' ?Val(RStDev(80, 104, 46, 36, 54, 28, 35, 72, 32, 62, 52, 28))
' ?RStDev(80, 104, 46, 36, 54, 28, 35, 72, 32, 62, 52, 28)
Option Explicit


Function RSum(ParamArray FieldValues()) As Variant
   '--------------------------------------------------
   ' Function RSum() adds all the arguments passed to it.
   ' If all arguments do not contain any data, RSum will return a
   ' null value.
   '--------------------------------------------------
   Dim dblTotal As Double, blnValid As Boolean
   Dim varArg As Variant
   For Each varArg In FieldValues
      If IsNumeric(varArg) Then
         blnValid = True
         dblTotal = dblTotal + varArg
      End If
   Next
   If blnValid Then ' One of the arguments was a number.
      RSum = dblTotal
   Else  ' Noo valid points to add.
      RSum = Null
   End If
End Function

Function RCount(ParamArray FieldValues()) As Variant
   '-------------------------------------------------
   ' Function RCount() will accept a variable number of arguments,
   ' and returns a count of arguments containing numbers.
   '-------------------------------------------------
   Dim lngCount As Long
   Dim varArg As Variant
   For Each varArg In FieldValues
      If IsNumeric(varArg) Then
         lngCount = lngCount + 1
      End If
   Next
   RCount = lngCount
End Function

Function RAvg(ParamArray FieldValues()) As Variant
   '----------------------------------------------------
   ' Function RAvg() will average all the numeric arguments passed to
   ' the function. If none of the arguments are numeric, it will
   ' return a null value.
   '-----------------------------------------------------
   Dim dblTotal As Double
   Dim lngCount As Long
   Dim varArg As Variant
   For Each varArg In FieldValues
      If IsNumeric(varArg) Then
         dblTotal = dblTotal + varArg
         lngCount = lngCount + 1
      End If
   Next
   If lngCount > 0 Then
      RAvg = dblTotal / lngCount
   Else
      RAvg = Null
   End If
End Function

Function RStDev(ParamArray FieldValues()) As Variant
   '---------------------------------------------------------
   ' Funzione RStDev() calcola la Deviazione Standard semplice
   '---------------------------------------------------------
   Dim dblSum As Double, dblSumOfSq As Double
   Dim n As Long
   Dim varArg As Variant
   For Each varArg In FieldValues
      If IsNumeric(varArg) Then
         dblSum = dblSum + varArg
         dblSumOfSq = dblSumOfSq + varArg * varArg
         n = n + 1
      End If
   Next
   If n > 1 Then ' Variance/StDev è applicabime se ho più di un valore
      RStDev = Sqr((n * dblSumOfSq - dblSum * dblSum) _
         / (n * (n - 1)))
   Else
      RStDev = Null
   End If
End Function

Function RStDevP(ParamArray FieldValues()) As Variant
   '-----------------------------------------------
   ' Function RStDevP() returns the Standard Deviation of the
   ' Population for all the arguments passed to it. The standard
   ' deviation of the population is only valid for one or more
   ' numeric values. If none of the arguments passed to
   ' the function contains a numeric value, the function will return
   ' a null.
   '-----------------------------------------------
   Dim dblSum As Double, dblSumOfSq As Double
   Dim n As Long
   Dim varArg As Variant
   For Each varArg In FieldValues
      If IsNumeric(varArg) Then
         dblSum = dblSum + varArg
         dblSumOfSq = dblSumOfSq + varArg * varArg
         n = n + 1
      End If
   Next
   If n > 0 Then 'only applies if points available
      RStDevP = Sqr((n * dblSumOfSq - dblSum * dblSum) / n / n)
   Else
      RStDevP = Null
   End If

End Function