Option Compare Database

Function InverseCDF(ByVal p As Double) As Double


' * INSERIRE LA PROBABILITA' E SI OTTIENE LA Z
'Define coefficients in rational approximations
Const A1 = -39.6968302866538
Const A2 = 220.946098424521
Const A3 = -275.928510446969
Const A4 = 138.357751867269
Const a5 = -30.6647980661472
Const a6 = 2.50662827745924

Const B1 = -54.4760987982241
Const B2 = 161.585836858041
Const B3 = -155.698979859887
Const b4 = 66.8013118877197
Const b5 = -13.2806815528857

Const C1 = -7.78489400243029E-03
Const C2 = -0.322396458041136
Const c3 = -2.40075827716184
Const c4 = -2.54973253934373
Const c5 = 4.37466414146497
Const c6 = 2.93816398269878

Const d1 = 7.78469570904146E-03
Const d2 = 0.32246712907004
Const d3 = 2.445134137143
Const d4 = 3.75440866190742

'Define break-points
Const p_low = 0.02425
Const p_high = 1 - p_low

'Define work variables
Dim q As Double, R As Double

'If argument out of bounds, raise error
'If p <= 0 Or p >= 1 Then Err.Raise 5

If p < p_low Then
  'Rational approximation for lower region
 ' q = Sqr(-2 * Log(p))
  InverseCDF = (((((C1 * q + C2) * q + c3) * q + c4) * q + c5) * q + c6) / _
    ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
ElseIf p <= p_high Then
  'Rational approximation for lower region
  q = p - 0.5
  R = q * q
  InverseCDF = (((((A1 * R + A2) * R + A3) * R + A4) * R + a5) * R + a6) * q / _
    (((((B1 * R + B2) * R + B3) * R + b4) * R + b5) * R + 1)
ElseIf p < 1 Then
  'Rational approximation for upper region
  q = Sqr(-2 * Log(1 - p))
  InverseCDF = -(((((C1 * q + C2) * q + c3) * q + c4) * q + c5) * q + c6) / _
    ((((d1 * q + d2) * q + d3) * q + d4) * q + 1)
End If

End Function