Attribute VB_Name = "modCDS"
Option Explicit

Function CDSwap(ByVal Nom1 As Double, ByVal T1 As Integer, ByVal d1 As Double, ByRef MProb1() As Variant, ByRef curva1() As propCurva, ByVal TC1 As Double, ByVal Tasa1 As Double, ByVal Per1 As Integer, ByVal Nom2 As Double, ByVal T2 As Integer, ByVal d2 As Integer, MProb2() As Variant, ByRef curva2() As propCurva, ByVal TC2 As Integer, ByVal Tasa2 As Double, ByVal Per2 As Integer)
Dim i As Integer
Dim pata1 As Double
Dim pata2 As Double
Dim t As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'NOM1      el monto nocional en caso de no default, en caso de default es el monto de rec
'T1        es el plazo en dias del primer flujo
'T2        es el plazo en dias del segundo flujo
'D1        es 1 si el pago se realiza cuando ocurre el default y 0 si son pagos periodicos siempre
'D2        similar a d1
'          y cuando no ocurra el default
'mprob1    es una matriz con la siguiente estructura
'1 col     el plazo
'2 col     la probabilidad de no default hasta t
'3 col     la prob condicional de default
'curva     son las curvas de descuento
'tc1 tc2   tipos de cambio
'tasa      es la tasa de los pagos periodicos en caso de no default
'per       periodicidad de los pagos


If d1 = 1 Then
    i = 1
    While MProb1(i + 1, 1) < T1
      pata1 = pata1 + Nom1 * ProbDefaultT(MProb1(i + 1, 1), MProb1) * FactDesc(MProb1(i + 1, 1), curva1)
      i = i + 1
    Wend
    pata1 = pata1 + Nom1 * ProbDefaultT(T1, MProb1) * FactDesc(T1, curva1)
ElseIf d1 = 0 Then
    t = T1
    While t > 0
      pata1 = pata1 + Nom1 * Tasa1 * (Per1 / 360) * ProbNDefault(t, MProb1) * FactDesc(t, curva1)
      t = t - Per1
    Wend
End If
    
If d2 = 1 Then
    i = 1
    While MProb2(i + 1, 1) < T2
       'se toman los plazos que trae el propio MProb2
        pata2 = pata2 + Nom2 * ProbDefaultT(MProb2(i + 1, 1), MProb2) * FactDesc(MProb2(i + 1, 1), curva2)
        i = i + 1
    Wend
    pata2 = pata2 + Nom2 * ProbDefaultT(T2, MProb2) * FactDesc(T2, curva2)
ElseIf d2 = 0 Then
    t = T2
    While t > 0
        pata2 = pata2 + Nom2 * Tasa2 * (Per2 / 360) * ProbNDefault(t, MProb2) * FactDesc(t, curva2)
        t = t - Per2
    Wend
End If

CDSwap = TC1 * pata1 + TC2 * pata2
'observaciones: esta formula supone que no hay amortizaciones del nocional

On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function ValCDSwap(ByRef curva1() As propCurva, ByRef curva2() As propCurva, ByRef curva3() As propCurva, ByRef curva4() As propCurva, Nom1, T1, d1, TC1, Tasa1, Per1, Nom2, T2, d2, TC2, Tasa2, Per2, trecp)
Dim probcal() As Variant
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'la misma que CDSwap, solo que conjunta 2 operaciones
' el calculo de probabilidades y
'la valuacion del cdswap
'curva 1 curva de valuacion 1
'curva 2 curva de valuacion 2
'curva 3 curva de valuacion 3
'curva 4 curva de valuacion 4

'primero se determinan las probabilidades de default con las curvas implicitas y
'la tasa de recuperacion
probcal = Calibra_Probabilidad(curva1, curva2, curva3, trecp, curva4)
ValCDSwap = CDSwap(Nom1, T1, d1, probcal, curva1, TC1, Tasa1, Per1, Nom2, T2, d2, probcal, curva2, TC2, Tasa2, Per2)
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function ValCDSwapFlujos(ByRef curva1() As propCurva, ByRef curva2() As propCurva, ByRef curva3() As propCurva, ByRef curva4() As propCurva, ByRef flujos1() As Variant, ByRef flujos2() As Variant, ByVal Nom1 As Double, ByVal T1 As Integer, ByVal d1 As Integer, ByVal TC1 As Integer, ByVal Tasa1 As Double, ByVal Per1 As Integer, ByVal Nom2 As Double, ByVal T2 As Integer, ByVal d2 As Integer, ByVal TC2 As Double, ByVal Tasa2 As Double, ByVal Per2 As Integer, ByVal trecp As Double)
Dim probcal() As Variant
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'la misma que ValCDSwap, solo que conjunta 2 operaciones
' el calculo de probabilidades y
'la valuacion del cdswap
probcal = Calibra_Probabilidad(curva1, curva2, curva3, trecp, curva4)
ValCDSwapFlujos = CDSwapFlujos(Nom1, flujos1, T1, d1, probcal, curva1, TC1, Tasa1, Per1, Nom2, flujos2, T2, d2, probcal, curva2, TC2, Tasa2, Per2)
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CDSwapFlujos(ByVal Nocional1 As Double, ByRef flujos1() As Variant, ByVal T1 As Integer, ByVal d1 As Integer, ByRef MProb1() As Variant, ByRef curva1() As propCurva, ByVal TC1 As Double, ByVal Tasa1 As Double, ByVal Per1 As Integer, ByVal Nocional2 As Double, ByRef flujos2() As Variant, ByVal T2 As Integer, ByVal d2 As Integer, ByRef MProb2() As Variant, ByRef curva2() As propCurva, ByVal TC2 As Double, ByVal Tasa2 As Double, ByVal Per2 As Integer)
Dim i As Integer
Dim p As Integer
Dim pata1 As Double
Dim pata2 As Double
Dim t As Integer
Dim mnocional As Double

'Nocional1 el monto nocional en caso de no default, en caso de default es el monto de rec
't1        es el plazo en dias
'd1        es 1 si el pago se realiza cuando ocurre el default y 0 si son pagos periodicos siempre
'd2        similar a d1
'          y cuando no ocurra el default
'mprob1    es una matriz donde
'1 col     el plazo
'2 col     la probabilidad de no default hasta t
'3 col     la prob condicional de default
'curva     son las curvas de descuento
'tc1 tc2   tipos de cambio
'tasa      es la tasa de los pagos periodicos en caso de no default
'per       periodicidad de los pagos


If d1 = 1 Then
    i = 1
    While MProb1(i + 1, 1) < T1
        pata1 = pata1 + Nocional1 * ProbDefaultT(MProb1(i + 1, 1), MProb1) * FactDesc(MProb1(i + 1, 1), curva1)
        i = i + 1
    Wend
    pata1 = pata1 + Nocional1 * ProbDefaultT(T1, MProb1) * FactDesc(T1, curva1)
ElseIf d1 = 0 Then
    t = T1
    While t > 0
        pata1 = pata1 + Nocional1 * Tasa1 * (Per1 / 360) * ProbNDefault(t, MProb1) * FactDesc(t, curva1)
        t = t - Per1
    Wend
End If
    
If d2 = 1 Then
    i = 1
    While MProb2(i + 1, 1) < T2
        'se toman los plazos que trae el propio MProb2
        pata2 = pata2 + Nocional2 * ProbDefaultT(MProb2(i + 1, 1), MProb2) * FactDesc(MProb2(i + 1, 1), curva2)
        i = i + 1
    Wend
    pata2 = pata2 + Nocional2 * ProbDefaultT(T2, MProb2) * FactDesc(T2, curva2)
ElseIf d2 = 0 Then
    t = T2
    While t > 0
        'se determina el nocional vigente hasta ese plazo
        For p = 1 To UBound(flujos2, 1)
         If t <= flujos2(p, 4) And t > flujos2(p - 1, 4) Then
          mnocional = flujos2(p, 2)
          Exit For
         End If
        Next p
        pata2 = pata2 + mnocional * Tasa2 * (Per2 / 360) * ProbNDefault(t, MProb2) * FactDesc(t, curva2)
        t = t - Per2
    Wend
End If

CDSwapFlujos = TC1 * pata1 + TC2 * pata2

End Function

Function FactDesc(ByVal plazo As Integer, ByRef curva() As propCurva)
Dim a As Integer
Dim i As Integer
Dim tasa As Double

If plazo <= curva(1).plazo Then
    a = 1
    FactDesc = (1 + curva(a).valor * curva(a).plazo / 360) ^ (-plazo / curva(a).plazo)
ElseIf plazo >= curva(UBound(curva, 1)).plazo Then
    a = UBound(curva, 1)
    FactDesc = (1 + curva(a).valor * curva(a).plazo / 360) ^ (-plazo / curva(a).plazo)
Else
    i = 2
    While plazo > curva(i).plazo
        i = i + 1
    Wend
    tasa = (curva(i).valor - curva(i - 1).valor) / (curva(i).plazo - curva(i - 1).plazo) * (plazo - curva(i - 1).plazo) + curva(i - 1).valor
    FactDesc = 1 / (1 + tasa * plazo / 360)
End If

End Function

Function ProbNDefault(ByVal plazo As Integer, ByRef M_Prob() As Variant)
Dim i As Integer
'Calcula la probabilidad de no caer en default hasta el momento Plazo, dada
'una matriz cuya primera columna es el plazo ti, la segunda la probabilidad
'de no caer en default hasta el tiempo ti y la tercera la probabilidad de caer
'en default entre ti y ti+1 dado que no se ha caido en default hasta ti

'primero busca el valor de i tal que
'M_Prob(i , 1) < Plazo <= M_Prob(i + 1, 1)
i = 1
While M_Prob(i + 1, 1) < plazo
   i = i + 1
Wend
If M_Prob(i + 1, 1) = plazo Then
    ProbNDefault = M_Prob(i + 1, 2)
Else
    ProbNDefault = M_Prob(i, 2) * (1 - M_Prob(i, 3)) ^ ((plazo - M_Prob(i, 1)) / (M_Prob(i + 1, 1) - M_Prob(i, 1)))
End If

End Function

Function ProbDefaultT(ByVal plazo As Integer, ByRef M_Prob() As Variant)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Dim i As Integer
'Calcula la probabilidad de caer en default entre ti y Plazo (ti<Plazo>ti+1), dada
'una matriz cuya primera columna es el plazo ti, la segunda la probabilidad
'de no caer en default hasta el tiempo ti y la tercera la probabilidad de caer
'en default entre ti y ti+1 dado que no se ha caido en default hasta ti

'primero busca el valor de i tal que
'M_Prob(i , 1) < Plazo <= M_Prob(i + 1, 1)
i = 1
While M_Prob(i + 1, 1) < plazo
   i = i + 1
Wend
If M_Prob(i + 1, 1) = plazo Then
    ProbDefaultT = M_Prob(i, 2) * M_Prob(i, 3)
Else
    ProbDefaultT = M_Prob(i, 2) * (1 - (1 - M_Prob(i, 3)) ^ ((plazo - M_Prob(i, 1)) / (M_Prob(i + 1, 1) - M_Prob(i, 1))))
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function Calibra_Probabilidad(ByRef Curva_ref() As propCurva, ByRef Curva_LR() As propCurva, ByRef Curva_desc() As propCurva, ByVal Recovery As Double, ByRef Curva_Spread() As propCurva)
Dim i As Integer
Dim h As Double
Dim epsilon As Double
Dim no_iteraciones As Integer
Dim Part As Integer
Dim ren As Integer
Dim TIni As Double
Dim TFin As Double
Dim Prob As Double
Dim Prob1 As Double
Dim fxo As Double
Dim fx1 As Double
Dim dfxo As Double
Dim kk As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'el objetivo de esta rutina es ajustar las probabilidades de default en tabla
'a los demas parametros introducidos
'Curva_ref             curva referencia
'Curva_LR              curva libre de riesgo
'curva_desc            curva descuento
'Curva_Spread          matriz con los spread
'recovery              tasa de recuperacion
'no_iteraciones        no de iteraciones
'h                     incremento para la derivada
'part                  intervalo de tiempo entre flujos
'
'mat_aux               tiene los siguientes campos
'1                     plazo i
'2                     probabilidad de no default hasta el plazo 7*(i-1)
'3                     probabilidad de default entre el plazo 7*(i-1) y el plazo 7*i

h = 10 ^ -8
epsilon = 10 ^ -10
no_iteraciones = 500

Part = 7      'convencion para semanas
ren = Int(Curva_Spread(UBound(Curva_Spread, 1), 2) / Part) 'el plazo maximo entre el tamaño de la particion


'se crean las matrices mat_aux, resul y resul1
ReDim Mat_Aux(1 To ren + 1, 1 To 3) As Variant
ReDim resul(1 To ren + 1, 1 To 3) As Variant
ReDim Resul1(1 To ren + 1, 1 To 3) As Variant

Mat_Aux(1, 1) = 0
resul(1, 1) = 0
Mat_Aux(1, 2) = 1
resul(1, 2) = 1
'se establecen las condiciones inciales para mat_aux
For i = 2 To ren + 1
    Mat_Aux(i, 1) = Part * (i - 1)    'plazo
    resul(i, 1) = Part * (i - 1)      'plazo
    Mat_Aux(i, 2) = (FactDesc(Part * (i - 1), Curva_ref) / FactDesc(Part * (i - 1), Curva_LR) - Recovery) / (1 - Recovery) 'probabilidad de sobrevivencia
    'implicita en las curva de referencia y la curva libre de riesgo
    Mat_Aux(i - 1, 3) = Mat_Aux(i - 1, 2) - Mat_Aux(i, 2) 'probabilidad de entrar en default en el periodo i
Next i

TIni = 0
'busca un objetivo
'se va alterando las probabilidades en plazos
'determinados
For i = 1 To UBound(Curva_Spread, 1) - 3 'se altero esta linea ¿por que?
    TFin = Curva_Spread(i).plazo     'se va avanzando en la curva de spreads
    'se fija una semilla
    Prob = 0.0001
    resul = AproxCurva(TIni, TFin, resul, Prob, Mat_Aux)
    'se valua un CDSwap hipotetico con prob
    fxo = CDSwap(-1, TFin, 0, resul, Curva_desc, 1, Curva_Spread(i).valor, 182, (1 - Recovery), TFin, 1, resul, Curva_desc, 1, 0, 0)
    'resul se modifica con la funcion aprox y un incremento h
    resul = AproxCurva(TIni, TFin, resul, Prob + h, Mat_Aux)
    'se valua otra vez para obtener el CDSwap similar
    fx1 = CDSwap(-1, TFin, 0, resul, Curva_desc, 1, Curva_Spread(i).valor, 182, (1 - Recovery), TFin, 1, resul, Curva_desc, 1, 0, 0)
    'se obtiene el valor del CDSwap en 2 puntos
    'se obtiene la pendiente de esta funcion en el punto prob
    dfxo = (fx1 - fxo) / h
    For kk = 1 To no_iteraciones
    'se determina el nuevo valor prob1
        Prob1 = Prob - (fxo / dfxo)
        resul = AproxCurva(TIni, TFin, resul, Prob1, Mat_Aux)
        'se valua un CDSwap con plazo hasta tfin
        fx1 = CDSwap(-1, TFin, 0, resul, Curva_desc, 1, Curva_Spread(i, 1), 182, (1 - Recovery), TFin, 1, resul, Curva_desc, 1, 0, 0)
        'se debe de anular el CDSwap
        If Abs(fx1) < epsilon Or Abs(Prob1 - Prob) < epsilon Then
            Prob = Prob1
            Exit For
        Else
        'se itera una vez mas el proceso de calibracion de las probabilidades
            Prob = Prob1
            fxo = fx1
            resul = AproxCurva(TIni, TFin, resul, Prob + h, Mat_Aux)
            fx1 = CDSwap(-1, TFin, 0, resul, Curva_desc, 1, Curva_Spread(i, 1), 182, (1 - Recovery), TFin, 1, resul, Curva_desc, 1, 0, 0)
            dfxo = (fx1 - fxo) / h
        End If
    Next kk
    resul = AproxCurva(TIni, TFin, resul, Prob, Mat_Aux)
    TIni = TFin
Next i
  If Prob < 0 Then MsgBox "Atencion, prob < 0"
Calibra_Probabilidad = resul

On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function AproxCurva(ByVal PIni As Integer, ByVal PFin As Integer, ByRef mat() As Variant, ByVal Prob As Double, ByRef Mat_A() As Variant)
Dim i As Integer
Dim j As Integer

i = 1
While mat(i, 1) < PIni 'se busca i donde mat(i, 1) >= PIni
    i = i + 1
Wend

j = i
'While Mat(j, 1) < PFin
'    If Mat(j, 1) <> Mat_A(j, 1) Then
'        Exit Function
'    End If
'    j = j + 1
'Wend
j = i

'se ajustan todas las probabilidades desde pini hasta pfin
While mat(j, 1) < PFin
    mat(j, 3) = Prob * Mat_A(j, 3) / Mat_A(i, 3)
'    se construye el valor p(i+1) ajustado
    mat(j + 1, 2) = mat(j, 2) * (1 - mat(j, 3))
    j = j + 1
Wend
AproxCurva = mat

End Function

Function Newton_Rapson(ByVal x0 As Double, ByVal fxo As Double, ByVal dfxo As Double, ByVal m As Integer, ByVal delta As Double, ByVal epsilon As Double)
Dim n As Integer
Dim x1 As Double

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

While Abs(fxo) > epsilon
    x1 = x0 - (fxo / dfxo)
    If Abs(x1 - x0) > delta Or n <= m Then
        x0 = x1
        n = n + 1
    End If
Wend

Newton_Rapson = x0

On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

