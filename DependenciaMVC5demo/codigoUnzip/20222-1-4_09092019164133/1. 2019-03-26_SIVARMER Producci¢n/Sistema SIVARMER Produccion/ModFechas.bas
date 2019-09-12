Attribute VB_Name = "ModFechas"
Option Explicit

Function LastDay(ByVal mes As Integer, ByVal año As Integer) As Integer
Select Case mes
    Case Is = 1, 3, 5, 7, 8, 10, 12
        LastDay = 31
    Case Is = 2
        If año Mod 4 = 0 Then
            LastDay = 29
        Else
            LastDay = 28
        End If
    Case Else
        LastDay = 30
End Select
        
End Function

Function JSanto(ByVal año As Integer) As Date
Dim a As Integer
Dim B As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer


a = año Mod 19
B = año Mod 4
c = año Mod 7
d = (19 * a + 24) Mod 30
e = (2 * B + 4 * c + 6 * d + 5) Mod 7
JSanto = DateSerial(año, 3, 15) + d + e + 4

End Function

Function VSanto(ByVal año As Integer) As Date
Dim a As Integer
Dim B As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer

a = año Mod 19
B = año Mod 4
c = año Mod 7
d = (19 * a + 24) Mod 30
e = (2 * B + 4 * c + 6 * d + 5) Mod 7
VSanto = DateSerial(año, 3, 15) + d + e + 5

End Function

Function FFechaP(ByVal j As Integer, ByVal dia As Integer, ByVal mes As Integer, ByVal año As Integer) As Date
Dim a As Date
Dim s As Integer

'Se obtiene la fecha del j-ésimo día de la semana del mes y año que se alimenta.
a = DateSerial(año, mes, 1)
s = dia - Weekday(a, vbMonday)
If s < 0 Then
    FFechaP = a + s + 7 * j
Else
    FFechaP = a + s + 7 * (j - 1)
End If

End Function

Function FechaU(ByVal dia As Integer, ByVal mes As Integer, ByVal año As Integer) As Date
Dim a As Date
Dim s As Integer
'Se obtiene la fecha del último día de la semana del mes y año que se alimenta

a = LastDay(mes, año)
s = dia - Weekday(a, vbMonday)
FechaU = a + s

End Function

Function SL_US(ByVal f_val As Date) As Date
Dim Temp As Integer
'Obtiene la f_val del siguiente lunes de una f_val no laborable en US
'cuando ésta cae en fin de semana, si no es la misma f_val que se alimenta.

Temp = Weekday(f_val, vbMonday)
If Temp > 5 Then
    SL_US = f_val + 8 - Temp
Else
    SL_US = f_val
End If

End Function

Function NoLabMX(ByVal f_val As Date) As Boolean
Dim a As Integer
Dim dia As Integer
Dim mes As Integer
Dim año As Integer
Dim lunes As Integer

a = Weekday(f_val, vbSunday)
If a = 1 Or a = 7 Then
    NoLabMX = True
Else
    dia = Day(f_val)
    mes = Month(f_val)
    año = Year(f_val)
    NoLabMX = False
    Select Case mes
        Case Is = 1
            If dia = 1 Then NoLabMX = True
        Case Is = 2
            lunes = plunes(f_val)
            If dia = lunes Then NoLabMX = True
        Case Is = 3
            If dia = 21 And año <= 2006 Then NoLabMX = True
            If dia = tlunes(f_val) And año > 2006 Then NoLabMX = True
            If f_val = JSanto(año) Then NoLabMX = True
            If f_val = VSanto(año) Then NoLabMX = True
        Case Is = 4
            If f_val = JSanto(año) Then NoLabMX = True
            If f_val = VSanto(año) Then NoLabMX = True
        Case Is = 5
            If dia = 1 Then NoLabMX = True
        Case Is = 9
            If dia = 16 Then NoLabMX = True
        Case Is = 11
            If dia = 2 Then NoLabMX = True
            If dia = tlunes(f_val) Then NoLabMX = True
        Case Is = 12
            If dia = 12 Or dia = 25 Then
                NoLabMX = True
            ElseIf dia = 1 And (año - 2000) Mod 6 = 0 Then
                NoLabMX = True
            End If
    End Select
End If

End Function

Function NolabUS(ByVal f_val As Date) As Boolean
Dim a As Integer
Dim año As Integer

NolabUS = False
a = Weekday(f_val, vbMonday)
If a > 5 Then
    NolabUS = True         'sabado o domingo
Else
    año = Year(f_val)
    Select Case f_val
        Case Is = SL_US(DateSerial(año, 1, 1)): NolabUS = True        'primero de enero
        Case Is = FFechaP(3, 1, 1, año): NolabUS = True               'dia de martin luther king
        Case Is = FFechaP(3, 1, 2, año): NolabUS = True               'dia del presidente
        Case Is = JSanto(año): NolabUS = True                         'el viernes santo
        Case Is = VSanto(año): NolabUS = True                         'el viernes santo
        Case Is = FechaU(1, 5, año): NolabUS = True                   'memorial day
        Case Is = SL_US(DateSerial(año, 7, 4)): NolabUS = True        'dia de la independencia
        Case Is = FFechaP(1, 1, 9, año): NolabUS = True               'labor day
        Case Is = FFechaP(2, 1, 10, año): NolabUS = True              'columbus day
        Case Is = SL_US(DateSerial(año, 11, 11)): NolabUS = True      'dia de los veteranos
        Case Is = FFechaP(4, 4, 11, año): NolabUS = True              'thanksgiving day
        Case Is = SL_US(DateSerial(año, 12, 25)): NolabUS = True      'navidad
    End Select
End If
End Function

Function NolabUK(ByVal f_val As Date) As Boolean
Dim a As Integer
Dim año As Integer

NolabUK = False
a = Weekday(f_val, vbMonday)
If a > 5 Then
    NolabUK = True         'sabado o domingo
Else
    año = Year(f_val)
    Select Case f_val
        Case Is = SL_US(DateSerial(año, 1, 1)): NolabUK = True          'primero de enero
        Case Is = JSanto(año): NolabUK = True                           'el viernes santo
        Case Is = VSanto(año): NolabUK = True                           'el viernes santo
        Case Is = VSanto(año) + 3: NolabUK = True                       'lunes de pascua
        Case Is = SL_US(DateSerial(año, 12, 25)): NolabUK = True        'navidad
    End Select
End If
End Function

Function NolabMXUS(ByVal f_val As Date) As Boolean
If NoLabMX(f_val) Then
    NolabMXUS = True
Else
    NolabMXUS = NolabUS(f_val)
End If
End Function

Function NolabMXUSUK(ByVal f_val As Date) As Boolean
If NoLabMX(f_val) Or NolabUS(f_val) Or NolabUS(f_val) Then
    NolabMXUSUK = True
Else
    NolabMXUSUK = False
End If
End Function

Function DAño(ByVal año As Integer) As Integer
'Calcula el número de días en el año
  If año Mod 4 = 0 Then
     DAño = 366
  Else
     DAño = 365
  End If
End Function

Function DefPlazo(ByVal FIni As Date, ByVal ffin As Date, ByVal Base As String) As Double
Dim a1 As Integer
Dim a2 As Integer
Dim m1 As Integer
Dim m2 As Integer
Dim d1 As Integer
Dim d2 As Integer
Dim fech1 As Date
Dim fech2 As Date

Select Case Base
    Case Is = "Actual/360"
        DefPlazo = (ffin - FIni) / 360
    Case Is = "Actual/365"
        DefPlazo = (ffin - FIni) / 365
    Case Is = "180/360"
        DefPlazo = 180 / 360
    Case Is = "180/360 ADJ"
        fech1 = FechaFinMesCercana(FIni)
        fech2 = FechaFinMesCercana(ffin)
        DefPlazo = (180 + ffin - FIni - (fech2 - fech1)) / 360
    Case Is = "30/360"
        DefPlazo = ((Year(ffin) - Year(FIni)) + 30 * (Month(ffin) - Month(FIni)) + Day(ffin) - Day(FIni)) / 360
    Case Is = "ACT/ACT"
        If Year(FIni) = Year(ffin) Then
            DefPlazo = (ffin - FIni) / DAño(Year(FIni))
        Else
            DefPlazo = (Year(ffin) - 1 - Year(FIni)) + (DateSerial(Year(FIni) + 1, 1, 1) - FIni) / DAño(Year(FIni)) + (ffin - DateSerial(Year(ffin) - 1, 12, 31)) / DAño(ffin)
        End If
    Case Is = "30/360US"
        a1 = Year(FIni)
        a2 = Year(ffin)
        m1 = Month(FIni)
        m2 = Month(ffin)
        d1 = Day(FIni)
        d2 = Day(ffin)
        If d1 = 31 Then d1 = 30
        If d2 = 31 Then
            If d1 = 30 Then
                d2 = 30
            Else
                d2 = 1
                If m2 = 12 Then
                    m2 = 1
                    a2 = a2 + 1
                Else
                    m2 = m2 + 1
                End If
            End If
        End If
        DefPlazo = 360 * (a2 - a1) + 30 * (m2 - m1) + (d2 - d1)
    Case Is = "30/360EU"
        a1 = Year(FIni)
        a2 = Year(ffin)
        m1 = Month(FIni)
        m2 = Month(ffin)
        d1 = Day(FIni)
        d2 = Day(ffin)
        If d1 = 31 Then d1 = 30
        If d2 = 31 Then d2 = 30
        DefPlazo = 360 * (a2 - a1) + 30 * (m2 - m1) + (d2 - d1)
    Case Is = "30/360IT"
        a1 = Year(FIni)
        a2 = Year(ffin)
        m1 = Month(FIni)
        m2 = Month(ffin)
        d1 = Day(FIni)
        d2 = Day(ffin)
        If m1 = 2 And d1 > 27 Then d1 = 30
        If m2 = 2 And d2 > 27 Then d2 = 30
        DefPlazo = 360 * (a2 - a1) + 30 * (m2 - m1) + (d2 - d1)
    End Select
    
End Function

Function FBD(ByVal f_val As Date, Calendario) As Date
Dim a As Boolean
Dim fechax As Date

fechax = f_val
'Convención Following Bussday
'La fechax que se obtiene es el PRIMER DIA LABORABLE posterior a la fechax introducida
'según el calendario de festivos especificado.

While True
      fechax = fechax + 1
      Select Case Calendario
             Case Is = "MX":       a = NoLabMX(fechax)
             Case Is = "US":       a = NolabUS(fechax)
             Case Is = "MXUS":     a = NolabMXUS(fechax)
      End Select
      If Not a Then
         FBD = fechax
         Exit Function
      End If
Wend

End Function

Function MFBD(ByVal f_val As Date, ByVal Calendario As String) As Date
Dim a As Boolean
Dim fechav1 As Date
'Convención Modified Following Bussday
'La f_val que se obtiene es el primer día laborable igual o posterior a la f_val introducida
'según el calendario de festivos especificado excepto en el caso que el día laborable esté en
'otro mes natural distinto al de la f_val introducida. En este caso se dará la última f_val
'laborable anterior a la f_val introducida.

Select Case Calendario
    Case Is = "MX":     a = NoLabMX(f_val)
    Case Is = "US":     a = NolabUS(f_val)
    Case Is = "MXUS":   a = NolabMXUS(f_val)
End Select
If Not a Then
    MFBD = f_val
Else
    fechav1 = f_val
    While a
        fechav1 = fechav1 + 1
        Select Case Calendario
            Case Is = "MX":     a = NoLabMX(fechav1)
            Case Is = "US":     a = NolabUS(fechav1)
            Case Is = "MXUS":   a = NolabMXUS(fechav1)
        End Select
    Wend
    If Month(fechav1) <> Month(f_val) Then
        a = True
        While a
            fechav1 = f_val - 1
            Select Case Calendario
                Case Is = "MX":     a = NoLabMX(fechav1)
                Case Is = "US":     a = NolabUS(fechav1)
                Case Is = "MXUS":   a = NolabMXUS(fechav1)
            End Select
        Wend
    End If
    MFBD = fechav1
End If

End Function

Function PBD(ByVal f_val As Date, ByVal Calendario As String) As Date
Dim a As Boolean
'Convención Preceding Bussines day
'La fechax que se obtiene es el primer día laborable anterior a la fechax introducida
'según el calendario de festivos especificado.

 Do While True
        Select Case Calendario
            Case Is = "MX":       a = NoLabMX(f_val)
            Case Is = "US":       a = NolabUS(f_val)
            Case Is = "MXUS":     a = NolabMXUS(f_val)
        End Select
        If Not a Then
            Exit Do
        End If
        f_val = f_val - 1
 Loop
 PBD = f_val
End Function

Function PBD1(ByVal f_val As Date, ByVal nd As Integer, ByVal Calendario As String) As Date
Dim a As Boolean
Dim fechax As Date
Dim contar As Integer

'Convención Preceding Bussines day
'Se obtiene el nd día laborable anterior a la f_val introducida
'según el calendario de festivos especificado.
fechax = f_val
contar = 1
    While contar <= nd
        fechax = fechax - 1
        Select Case Calendario
            Case Is = "MX":       a = NoLabMX(fechax)
            Case Is = "US":       a = NolabUS(fechax)
            Case Is = "MXUS":     a = NolabMXUS(fechax)
        End Select
        If Not a Then contar = contar + 1
    Wend
    PBD1 = fechax

End Function

Function MPBD(ByVal f_val As Date, ByVal Calendario As String) As Date
Dim a As Boolean
Dim fechav1 As Date

'Convención Modified Preceding Bussday
'La f_val que se obtiene es el primer día laborable igual o anterior a la f_val introducida
'según el calendario de festivos especificado excepto en el caso que el día laborable esté en
'otro mes natural distinto al de la f_val introducida. En este caso se dará la primera f_val
'laborable posterior a la f_val introducida.

Select Case Calendario
    Case Is = "MX":     a = NoLabMX(f_val)
    Case Is = "US":     a = NolabUS(f_val)
    Case Is = "MXUS":   a = NolabMXUS(f_val)
End Select
If Not a Then
    MPBD = f_val
Else
    fechav1 = f_val
    While a
        fechav1 = fechav1 - 1
        Select Case Calendario
            Case Is = "MX":     a = NoLabMX(fechav1)
            Case Is = "US":     a = NolabUS(fechav1)
            Case Is = "MXUS":   a = NolabMXUS(fechav1)
        End Select
    Wend
    If Month(fechav1) <> Month(f_val) Then
        a = True
        While a
            fechav1 = f_val + 1
            Select Case Calendario
                Case Is = "MX":     a = NoLabMX(fechav1)
                Case Is = "US":     a = NolabUS(fechav1)
                Case Is = "MXUS":   a = NolabMXUS(fechav1)
            End Select
        Wend
    End If
    MPBD = fechav1
End If

End Function

Function plunes(ByVal f_val As Date) As Integer
Dim año As Integer
Dim mes As Integer
Dim dia As Integer
Dim primerdia  As Date
Dim primerdiasem As Integer
Dim fechavn As Date

año = Year(f_val)
mes = Month(f_val)
dia = Day(f_val)


primerdia = (DateSerial(año, mes, 1))
primerdiasem = Weekday(primerdia)

If primerdiasem = 1 Then
fechavn = primerdia + 1
ElseIf primerdiasem = 2 Then
fechavn = primerdia
ElseIf primerdiasem = 3 Then
fechavn = primerdia + 6
ElseIf primerdiasem = 4 Then
fechavn = primerdia + 5
ElseIf primerdiasem = 5 Then
fechavn = primerdia + 4
ElseIf primerdiasem = 6 Then
fechavn = primerdia + 3
ElseIf primerdiasem = 7 Then
fechavn = primerdia + 2
End If
plunes = Day(fechavn)
End Function

Function tlunes(ByVal f_val As Date) As Integer
Dim año As Integer
Dim mes As Integer
Dim dia As Integer
Dim primerdia  As Date
Dim primerdiasem As Integer
Dim fechavn As Date

año = Year(f_val)
mes = Month(f_val)
dia = Day(f_val)


primerdia = (DateSerial(año, mes, 1))
primerdiasem = Weekday(primerdia)

If primerdiasem = 1 Then
fechavn = primerdia + 15
ElseIf primerdiasem = 2 Then
fechavn = primerdia + 14
ElseIf primerdiasem = 3 Then
fechavn = primerdia + 20
ElseIf primerdiasem = 4 Then
fechavn = primerdia + 19
ElseIf primerdiasem = 5 Then
fechavn = primerdia + 18
ElseIf primerdiasem = 6 Then
fechavn = primerdia + 17
ElseIf primerdiasem = 7 Then
fechavn = primerdia + 16
End If

tlunes = Day(fechavn)
End Function

Function DescDiasHabUS(ByVal fecha As Date, ByVal num As Integer) As Date
Dim fechax As Date
Dim contar As Integer

contar = 1
fechax = fecha
Do While contar <= num
fechax = fechax - 1
If Not NolabUS(fechax) Then
  contar = contar + 1
 End If
Loop
DescDiasHabUS = fechax
End Function

Function DescDiasHabUSUK(ByVal fecha As Date, ByVal num As Integer) As Date
Dim fechax As Date
Dim contar As Integer

contar = 1
fechax = fecha
Do While contar <= num
fechax = fechax - 1
If Not NolabUS(fechax) And Not NolabUK(fechax) Then
   contar = contar + 1
End If
Loop
DescDiasHabUSUK = fechax
End Function

Function esFinMes(ByVal fecha As Date) As Boolean
Dim fecha1 As Date
fecha1 = DateSerial(Year(fecha), Month(fecha) + 1, 1)
fecha1 = PBD1(fecha1, 1, "MX")
If fecha = fecha1 Then
   esFinMes = True
Else
   esFinMes = False
End If
End Function

Function Mestxt(ByVal mes As Integer) As String
If mes = 1 Then
   Mestxt = "enero"
ElseIf mes = 2 Then
   Mestxt = "febrero"
ElseIf mes = 3 Then
   Mestxt = "marzo"
ElseIf mes = 4 Then
   Mestxt = "abril"
ElseIf mes = 5 Then
   Mestxt = "mayo"
ElseIf mes = 6 Then
   Mestxt = "junio"
ElseIf mes = 7 Then
   Mestxt = "julio"
ElseIf mes = 8 Then
   Mestxt = "agosto"
ElseIf mes = 9 Then
   Mestxt = "septiembre"
ElseIf mes = 10 Then
   Mestxt = "octubre"
ElseIf mes = 11 Then
   Mestxt = "noviembre"
ElseIf mes = 12 Then
   Mestxt = "diciembre"
Else
   Mestxt = "desc"
End If
End Function
