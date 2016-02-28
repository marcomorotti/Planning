Option Compare Database

Function func_normsdist(z As Double) As Double
'******************************************************************
'*  Adapted from http://lib.stat.cmu.edu/apstat/66
'*  Evaluates the tail area of the standardised normal curve
'*  from x to infinity if upper is .true. or
'*  from minus infinity to x if upper is .false.
'*  INSERIRE LA Z E SI OTTIENE LA PROBABILITA'
'******************************************************************

Const a0 = 0.5
Const A1 = 0.398942280444
Const A2 = 0.399903438505
Const A3 = 5.75885480458
Const A4 = 29.8213557808
Const a5 = 2.62433121679
Const a6 = 48.6959930692
Const a7 = 5.92885724438

Const b0 = 0.398942280385
Const B1 = 3.8052 * 10 ^ (-8)
Const B2 = 1.00000615302
Const B3 = 3.98064794 * 10 ^ (-4)
Const b4 = 1.98615381364
Const b5 = 0.151679116635
Const b6 = 5.29330324926
Const b7 = 4.8385912808
Const b8 = 15.1508972451
Const b9 = 0.742380924027
Const b10 = 30.789933034
Const b11 = 3.99019417011

Dim zabs As Double
Dim pdf As Double
Dim p As Double
Dim q As Double
Dim Y As Double
Dim Temp As Double

zabs = Abs(z)

If zabs <= 12.7 Then
    Y = a0 * z * z
    pdf = Exp(-Y) * b0
    If zabs <= 1.28 Then
        Temp = Y + A3 - A4 / (Y + a5 + a6 / (Y + a7))
        q = a0 - zabs * (A1 - A2 * Y / Temp)
    Else
        Temp = (zabs - b5 + b6 / (zabs + b7 - b8 / (zabs + b9 + b10 / (zabs + b11))))
        q = pdf / (zabs - B1 + (B2 / (zabs + B3 + b4 / Temp)))
    End If
Else
    pdf = 0
    q = 0
End If

If z < 0 Then
    func_normsdist = q
Else
    func_normsdist = 1 - q
End If

End Function