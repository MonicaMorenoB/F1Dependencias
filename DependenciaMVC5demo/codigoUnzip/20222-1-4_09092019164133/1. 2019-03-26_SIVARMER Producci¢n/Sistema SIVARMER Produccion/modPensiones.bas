Attribute VB_Name = "modPensiones"
Option Explicit

Sub ImpPosPensiones(ByVal fecha As Date)
Dim i As Long
Dim j As Long
Dim txtfecha As String
Dim notabla1 As Integer
Dim notabla2 As Integer
Dim txtborra As String
Dim contar As Long
Dim mata() As Variant
Dim nreg1 As Long
Dim nreg2 As Long
Dim exito As Boolean
Dim txtnomarch1 As String
Dim txtnomarch2 As String

notabla1 = 7
notabla2 = 4
  ReDim nomtabla1(1 To notabla1) As String
  ReDim nomtabla2(1 To notabla2, 1 To 2) As String
  txtnomarch1 = "d:\fp\Fid. 2065 " & Format(fecha, "dd-mm-yyyy") & " (C).xls"
  txtnomarch2 = "d:\fp\Fid. 2160 " & Format(fecha, "dd-mm-yyyy") & " (C).xls"
  nomtabla1(1) = "BANOBRAS"
  nomtabla1(2) = "EVERCORE"
  nomtabla1(3) = "BANORTE"
  nomtabla1(4) = "SANTANDER"
  nomtabla1(5) = "GBM"
  nomtabla1(6) = "BANAMEX"
  nomtabla1(7) = "CB VECTOR"
  Open "d:\emisiones sin calificacion " & ClavePosPension1 & " " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #1
      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtborra = "DELETE FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension1 & " AND TIPOPOS = 1"
      ConAdo.Execute txtborra
      txtborra = "DELETE FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension1 & " AND TIPOPOS = 1"
      ConAdo.Execute txtborra
      contar = 0
      For j = 1 To notabla1
          mata = LeerHojaCalc1(txtnomarch1, nomtabla1(j))
          mata = depurartablafp1(mata, fecha, nomtabla1(j), 2065, contar)
          If UBound(mata, 1) <> 0 Then
             Call ImpPosFid(mata, 2065, nreg1, exito)
          End If
      Next j
  Close #1
  Open "d:\emisiones sin calificacion " & ClavePosPension2 & "  " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #1
  nomtabla2(1, 1) = "POSICION 988 Pensiones CD": nomtabla2(1, 2) = "988"
  nomtabla2(2, 1) = "POSICION 989 Oblig Post al R.": nomtabla2(2, 2) = "989"
  nomtabla2(3, 1) = "POSICION 990 Primas Ant.": nomtabla2(3, 2) = "990"
  nomtabla2(4, 1) = "POSICION 1111 Benef x Falle.": nomtabla2(4, 2) = "1111"
      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtborra = "DELETE FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension2 & " AND TIPOPOS = 1"
      ConAdo.Execute txtborra
      txtborra = "DELETE FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension2 & " AND TIPOPOS = 1"
      contar = 0
      ConAdo.Execute txtborra
      For j = 1 To notabla2
          mata = LeerHojaCalc2(txtnomarch2, nomtabla2(j, 1))
          mata = depurartablafp2(mata, fecha, nomtabla2(j, 2), contar)
          If UBound(mata, 1) <> 0 Then
             Call ImpPosFid(mata, 2160, nreg2, exito)
          End If
      Next j
  Close #1

End Sub

Sub ImpPosPensiones2(ByVal fecha As Date)
Dim i As Long
Dim j As Long
Dim txtfecha As String
Dim notabla1 As Integer
Dim notabla2 As Integer
Dim txtborra As String
Dim contar As Long
Dim mata() As Variant
Dim nreg1 As Long
Dim nreg2 As Long
Dim exito As Boolean
Dim txtnomarch1 As String
Dim txtnomarch2 As String

notabla1 = 7
notabla2 = 4
  ReDim nomtabla1(1 To notabla1) As String
  ReDim nomtabla2(1 To notabla2, 1 To 2) As String
  txtnomarch1 = "d:\fp\Fid. 2065 " & Format(fecha, "dd-mm-yyyy") & " (C).xls"
  txtnomarch2 = "d:\fp\Fid. 2160 " & Format(fecha, "dd-mm-yyyy") & " (C).xls"
  nomtabla1(1) = "BANOBRAS"
  nomtabla1(2) = "EVERCORE"
  nomtabla1(3) = "BANORTE"
  nomtabla1(4) = "SANTANDER"
  nomtabla1(5) = "GBM"
  nomtabla1(6) = "BANAMEX"
  nomtabla1(7) = "CB VECTOR"
  Open "d:\emisiones sin calificacion " & ClavePosPension1 & " " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #1
      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtborra = "DELETE FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension1 & " AND TIPOPOS = 1"
      ConAdo.Execute txtborra
      txtborra = "DELETE FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension1 & " AND TIPOPOS = 1"
      ConAdo.Execute txtborra
      contar = 0
      For j = 1 To notabla1
          mata = LeerHojaCalc3(txtnomarch1, nomtabla1(j))
          mata = depurartablafp1(mata, fecha, nomtabla1(j), 2065, contar)
          If UBound(mata, 1) <> 0 Then
             Call ImpPosFid(mata, 2065, nreg1, exito)
          End If
      Next j
  Close #1
  Open "d:\emisiones sin calificacion " & ClavePosPension2 & "  " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #1
  nomtabla2(1, 1) = "POSICION 988 Pensiones CD": nomtabla2(1, 2) = "988"
  nomtabla2(2, 1) = "POSICION 989 Oblig Post al R.": nomtabla2(2, 2) = "989"
  nomtabla2(3, 1) = "POSICION 990 Primas Ant.": nomtabla2(3, 2) = "990"
  nomtabla2(4, 1) = "POSICION 1111 Benef x Falle.": nomtabla2(4, 2) = "1111"
      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtborra = "DELETE FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension2 & " AND TIPOPOS = 1"
      ConAdo.Execute txtborra
      txtborra = "DELETE FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPension2 & " AND TIPOPOS = 1"
      contar = 0
      ConAdo.Execute txtborra
      For j = 1 To notabla2
          mata = LeerHojaCalc4(txtnomarch2, nomtabla2(j, 1))
          mata = depurartablafp2(mata, fecha, nomtabla2(j, 2), contar)
          If UBound(mata, 1) <> 0 Then
             Call ImpPosFid(mata, 2160, nreg2, exito)
          End If
      Next j
  Close #1

End Sub

Sub ValidarPosPension1(ByVal fecha As Date, ByVal cposicion As Integer)
Dim mata() As propPosMD
Dim matb() As propPosDiv
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim sihayarch As Boolean
Dim i As Long
Dim indice As Long
Dim indice1 As Long
Dim noreg As Long
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String

'rutina para leer la posicion del fondo de pensiones
    txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
    txtfiltro1 = "SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
    txtfiltro1 = txtfiltro1 & " AND TIPOPOS = 1"
    txtfiltro1 = txtfiltro1 & " AND CPOSICION = " & cposicion
    mata = LeerBaseMD(txtfiltro1)
    matvp = LeerVPrecios(fecha, mindvp)
    If UBound(mata, 1) <> 0 Then
       For i = 1 To UBound(mata, 1)
           indice = BuscarValorArray(mata(i).cEmisionMD, mindvp, 1)
           If indice <> 0 Then
           indice1 = mindvp(indice, 2)
           If mata(i).tValorMD = matvp(indice1, 3) And mata(i).emisionMD = matvp(indice, 4) And mata(i).serieMD = matvp(indice, 5) Then
              Call EncontrarInstTabValuacion(matvp(indice1).c_emision, matvp(indice1).tv, matvp(indice1).emision, matvp(indice1).serie, matvp(indice1).tmercado, matvp(indice1).st_colocacion, matvp(indice1).st_colocacion)
              Call EncontrarInstTabParametros(matvp(indice1).c_emision, matvp(indice1).tv, matvp(indice1).emision, matvp(indice1).serie, matvp(indice1).c_emision, matvp(indice1).tv)
              If mata(i).SiFlujosMD = "S" Then Call DetermSiHayFlujosEm(matvp(indice1).c_emision, matvp(indice1).tv, matvp(indice1).regla_cupon, matvp(indice1).femision, matvp(indice1).fvenc)
              Call DetermSiFRBD(matvp(indice, 3), matvp(indice1, 4), matvp(indice, 5))
            Else
              Print #8, "la clave de emision " & mata(i).cEmisionMD & "no esta bien construida"
           End If
          Else
             Print #8, "la emision " & mata(i).cEmisionMD & "no se encuentra en el vector de precios"
          End If
          DoEvents
        Next i
    End If
    txtfiltro2 = "SELECT * FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1"
    txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion
    matb = LeerPosDiv(txtfiltro2)
    If UBound(matb, 1) <> 0 Then
       For i = 1 To UBound(matb, 1)
           indice = BuscarValorArray(matb(i).CEmisionDiv, mindvp, 1)
           If indice <> 0 Then
              indice1 = mindvp(indice, 2)
              If matb(i).TValorDiv = matvp(indice, 3) And matb(i).EmisionDiv = matvp(indice, 4) And matb(i).SerieDiv = matvp(indice, 5) Then
                 Call EncontrarInstTabValuacion(matvp(indice1).c_emision, matvp(indice1).tv, matvp(indice1).emision, matvp(indice1).serie, matvp(indice1).c_emision, matvp(indice1).c_emision, matvp(indice1).c_emision)
                 Call EncontrarInstTabParametros(matvp(indice1).c_emision, matvp(indice1).c_emision, matvp(indice1).c_emision, matvp(indice1).c_emision, matvp(indice1).c_emision, matvp(indice1).c_emision)
                 Call DetermSiFRBD(matvp(indice1).tv, matvp(indice1).emision, matvp(indice1).serie)
              Else
                 Print #8, "la clave de emision " & cposicion & " " & matb(i).CEmisionDiv & "no esta bien construida"
              End If
           Else
              Print #8, "la emision " & matb(i).CEmisionDiv & "no se encuentra en el vector de precios"
           End If
           DoEvents
       Next i
    End If
End Sub

Sub ValidarPosPension2(ByVal fecha As Date, ByVal cposicion As Integer, ByVal nop As Integer)
Dim mata() As propPosMD
Dim matb() As propPosDiv
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim sihayarch As Boolean
Dim i As Long
Dim indice As Long
Dim noreg As Long
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String

'rutina para leer la posicion del fondo de pensiones
    txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
    txtfiltro1 = "SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
    txtfiltro1 = txtfiltro1 & " AND TIPOPOS = 1"
    txtfiltro1 = txtfiltro1 & " AND CPOSICION = " & cposicion
    mata = LeerBaseMD(txtfiltro1)
    matvp = LeerVPrecios(fecha, mindvp)
    If UBound(mata, 1) <> 0 Then
       For i = 1 To UBound(mata, 1)
           indice = BuscarValorArray(mata(i).cEmisionMD, matvp, 22)
           If indice <> 0 Then
              Call DeterminaSiHayFR(matvp(indice, 3), matvp(indice, 4), matvp(indice, 5), #1/1/2015#, fecha, nop)
           Else
             Print #nop, "la emision " & mata(i).cEmisionMD & "no se encuentra en el vector de precios"
          End If
          DoEvents
        Next i
    End If
    txtfiltro2 = "SELECT * FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1"
    txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion
    matb = LeerPosDiv(txtfiltro2)
    If UBound(matb, 1) <> 0 Then
       For i = 1 To UBound(matb, 1)
           indice = BuscarValorArray(matb(i).CEmisionDiv, matvp, 22)
           If indice <> 0 Then
              Call DeterminaSiHayFR(matvp(indice, 3), matvp(indice, 4), matvp(indice, 5), #1/1/2016#, fecha, nop)
          Else
            Print #nop, "la emision " & matb(i).CEmisionDiv & "no se encuentra en el vector de precios"
          End If
          DoEvents
        Next i
    End If
End Sub

Sub ValidarPosPension3(ByVal fecha As Date, ByVal cposicion As Integer, ByVal nop As Integer)
Dim mata() As propPosMD
Dim matb() As propPosDiv
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim sihayarch As Boolean
Dim i As Long
Dim indice As Long
Dim noreg As Long
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String

'rutina para leer la posicion del fondo de pensiones
    txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
    txtfiltro1 = "SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
    txtfiltro1 = txtfiltro1 & " AND TIPOPOS = 1"
    txtfiltro1 = txtfiltro1 & " AND CPOSICION = " & cposicion
    mata = LeerBaseMD(txtfiltro1)
    matvp = LeerVPrecios(fecha, mindvp)
    If UBound(mata, 1) <> 0 Then
       For i = 1 To UBound(mata, 1)
           indice = BuscarValorArray(mata(i).cEmisionMD, matvp, 22)
           If indice <> 0 Then
              Call DeterminaSiHayFR2(matvp(indice, 3), matvp(indice, 4), matvp(indice, 5), #1/1/2015#, fecha, nop)
           Else
             Print #nop, "la emision " & mata(i).cEmisionMD & "no se encuentra en el vector de precios"
          End If
          DoEvents
        Next i
    End If
    txtfiltro2 = "SELECT * FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1"
    txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion
    matb = LeerPosDiv(txtfiltro2)
    If UBound(matb, 1) <> 0 Then
       For i = 1 To UBound(matb, 1)
           indice = BuscarValorArray(matb(i).CEmisionDiv, matvp, 22)
           If indice <> 0 Then
              Call DeterminaSiHayFR2(matvp(indice, 3), matvp(indice, 4), matvp(indice, 5), #1/1/2015#, fecha, nop)
          Else
            Print #nop, "la emision " & matb(i).CEmisionDiv & "no se encuentra en el vector de precios"
          End If
          DoEvents
        Next i
    End If
End Sub


Function ValidarPosPension4(ByRef mata() As Variant, ByVal fecha As Date) As Variant()
Dim i As Long
Dim noreg As Long
Dim nocampos As Long
Dim fcompra As Date
Dim fven As Date
Dim fvenrep As Date
Dim fpos As Date
Dim tipooper As Integer
Dim txtfpos As String
Dim txtfven As String
Dim txtfcompra As String
Dim emision As String
Dim tv As String
Dim serie As String
Dim tcupon As Double
Dim tpremio As Double
Dim tmercado As String
Dim txtfecha As String
Dim txtfvenrep As String
Dim toper As Integer
Dim tvalor As String


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se procede a verificar los datos de la posicion
'ademas se procede a clasificarlos y guardar otros parametros para su uso
'corrige algunos errores de la tabla de datos

'1 fecha pos
'2 intencion
'3 tipo de operacion
'4 tipo valor
'5 emision
'6 serie
'7 no titulos
'8 fecha vencimiento
'  campo1
'  campo2
'  campo3
'12 tasa reporto
'13 precio pactado
'14 propietario
'15 fecha compra

 '1  FECHAPOSISION
 '2  PROPIETARIO
 '3  TIPOPOSICION
 '4  TIPOOPERACION
 '5  NUMTPAPEL
 '6  TIPOVALOR
 '7  NOMEMI
 '8  serie
 '9  TITULOS
 '10 FECHAVENC
 '11 ValorNominal
 '12 PERIODOCUPON
 '13 TASACUPON
 '14 TASAPREMIO
 '15 precio
 '16 FECHACOMPRA
 '17 tmercado
 '18 ?

If IsArray(mata) Then
noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
ReDim matb(1 To noreg, 1 To nocampos + 3) As Variant

For i = 1 To noreg
fcompra = 0
fven = 0
fvenrep = 0
fpos = 0
tipooper = mata(i, 3)             'tipo de operacion
txtfpos = Trim(mata(i, 1))        'fecha de la posicion
txtfven = Trim(mata(i, 8))        'fecha de vencimiento de la operacion
txtfcompra = Trim(mata(i, 15))    'fecha compra
emision = Trim(mata(i, 5))        'emision
tv = UCase(Trim(mata(i, 4)))      'tipo valor
serie = UCase(Trim(mata(i, 6)))   'serie
tcupon = Val(mata(i, 11)) / 100
tpremio = Val(mata(i, 12))
tmercado = ""
If Format(Val(txtfpos), "00000000") = txtfpos Then 'fecha de posicion
 txtfecha = Mid(txtfpos, 7, 2) & "/" & Mid(txtfpos, 5, 2) & "/" & Mid(txtfpos, 1, 4)
 If IsDate(txtfecha) Then fpos = CDate(txtfecha)
End If
If Format(Val(txtfven), "00000000") = txtfven And txtfven <> "99999999" Then
'fecha de vencimiento
 txtfecha = Mid(txtfven, 7, 2) & "/" & Mid(txtfven, 5, 2) & "/" & Mid(txtfven, 1, 4)
 If IsDate(txtfecha) Then fven = CDate(txtfecha)
End If
If Format(Val(txtfcompra), "00000000") = txtfcompra And txtfcompra <> "99999999" Then
'fecha de compra
 txtfecha = Mid(txtfcompra, 7, 2) & "/" & Mid(txtfcompra, 5, 2) & "/" & Mid(txtfcompra, 1, 4)
 If IsDate(txtfecha) Then fcompra = CDate(txtfecha)
End If
If Format(Val(txtfvenrep), "00000000") = txtfvenrep And txtfvenrep <> "99999999" Then
 txtfecha = Mid(txtfvenrep, 7, 2) & "/" & Mid(txtfvenrep, 5, 2) & "/" & Mid(txtfvenrep, 1, 4)
 If IsDate(txtfecha) Then fvenrep = CDate(txtfecha)
End If
If tipooper = "D" Then
 toper = 1
ElseIf tipooper = "R" Then
 toper = 2
ElseIf tipooper = "E" Then
 toper = 1
End If
If tv = "B" Then
 emision = "CETES"
 tmercado = "MD"
End If
If tv = "BI" Then
 emision = "CETES IMP"
 tmercado = "MD"
End If
If tv = "LP" Then
 emision = "BONDE91"
End If
If tv = "LD" Then
 tmercado = "MD"
End If
If tv = "LT" Then
 emision = "BONDEST"
End If
If tv = "LS" Then
 emision = "BOND182"
 tmercado = "MD"
End If
If tv = "IP" Then
 tmercado = "MD"
End If
If tv = "IT" Then
 tmercado = "MD"
End If
If tv = "2U" Then
 tmercado = "MD"
End If
If tv = "M" Or tv = "M0" Or tv = "M2" Or tv = "M3" Or tv = "M5" Or tv = "M7" Then
 emision = "BONOS"
 tmercado = "MD"
End If
If tv = "M2" Then
 tvalor = "M3"
 emision = "BONOS"
End If
If tv = "PI" Then
 tmercado = "MD"
End If
If tv = "S0" Or tv = "S3" Or tv = "S5" Or tv Then
 tmercado = "MD"
End If
If tv = "V" And emision = "WLEASE" Then
 tv = "91"
 emision = "VWLEASE"
 tmercado = "MD"
End If
If tv = "90" Then
 tmercado = "MD"
End If
If tv = "XA" Then
 emision = "BREMS"
 tmercado = "MD"
End If
If emision = "FINASA" And serie = "1-96" Then
 serie = "ENE-96"
End If
If emision = "BONDE91" And serie = "040429" Then
 tv = "BONDEST"
End If
If Left(tv, 2) = "D1" Then
 tmercado = "MD"
End If
If Left(tv, 2) = "D1" And emision = "UMS09V" And serie = "FRN2009" Then
 emision = "MEXR92"
 serie = "090113"
End If

If tv = "CH" Then
 tv = "90"
 tmercado = "MD"
End If


If tv = "M0" And serie = "131219" Then
 tv = "M"
End If



If Left(tv, 1) = "I" And (emision = "BACOMER" Or emision = "BANEJER" Or emision = "BACMEXT" Or emision = "BANSAN" Or emision = "NAFIN" Or emision = "INBURSA" Or emision = "BANORTE" Or emision = "SHF") Then
 tmercado = "MD"
End If

'1fecha pos
'2intencion
'3tipo de operacion
'4tipo valor
'5emision
'6serie
'7no titulos
'8fecha vencimiento
'campo1
'campo2
'campo3
'12tasa reporto
'13precio pactado
'14propietario
'15fecha compra
 

' If tmercado = "" Then Call MostrarMensajeSistema("Validando la posicion del fondo de pensiones.  No se definio el mercado " & tv, frmprogreso.label2, 0.1, Date, Time, NomUsuario)
 
 matb(i, 1) = fpos
 matb(i, 2) = mata(i, 2)
 matb(i, 3) = toper
 matb(i, 4) = tv           'tipo valor
 matb(i, 5) = emision      'emision
 matb(i, 6) = serie        'serie
 matb(i, 7) = mata(i, 7)   'no titulos
 matb(i, 8) = fven         'fecha vencimiento
 matb(i, 9) = mata(i, 9)
 matb(i, 10) = mata(i, 10)
 matb(i, 11) = mata(i, 11)
 matb(i, 12) = tpremio     'tasa reporto
 matb(i, 13) = mata(i, 13) 'precio pactado
 matb(i, 14) = mata(i, 14) 'propietario
 matb(i, 15) = fcompra     'fecha compra
 matb(i, 16) = tmercado    'tipo mercado
 matb(i, 17) = 4           'mesa o posicion
 
AvanceProc = i / noreg
MensajeProc = "Verificando el F Pensiones del " & fecha & ": " & Format(AvanceProc, "#,##0.00 %")

Next i
ValidarPosPension4 = matb
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub ProcCalculoCVaRFondoP(ByVal fecha As Date, ByVal txtport As String, ByRef matport() As String, ByVal impflujos As Boolean, ByVal dirflujos As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef txtmsg As String)
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim txtportfr As String
Dim txttvar As String
Dim i As Long
Dim nconf As Double
Dim valor As Double
Dim exito As Boolean

txtportfr = "Normal"
txttvar = "CVARH"
nconf = 0.97


If impflujos Then
   Call ImportFlujosPIP(fecha, dirflujos)
End If
MatCurvasT = LeerCurvaCompleta(fecha, exito)
Call RutinaValPort(fecha, fecha, fecha, txtport, matpos, matposmd, 1, txtmsg, exito)
If exito Then Call GuardarResValPort(fecha, fecha, fecha, txtport, txtportfr, matpos, matposmd, 1, exito)
If exito Then
   Call GenValPortPosPension(fecha, fecha, fecha, txtport, txtport, txtportfr, 1, exito)
   For i = 1 To UBound(matport, 1)
       Call GenValPortPosPension(fecha, fecha, fecha, txtport, matport(i, 1), txtportfr, 1, exito)
   Next i
   Call SubprocCalculoPyGPort(fecha, fecha, fecha, txtport, txtportfr, noesc, htiempo, txtmsg, exito)
   Call ConsolidaPyGSubport(fecha, fecha, fecha, txtport, txtportfr, txtport, noesc, htiempo, txtmsg, exito)
   valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, txtport, noesc, htiempo, 1 - nconf, exito)
   If exito Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, txtport, txtportfr, txttvar, noesc, htiempo, 0, 1 - nconf, 0, valor)
   For i = 1 To UBound(matport, 1)
       Call ConsolidaPyGSubport(fecha, fecha, fecha, txtport, txtportfr, matport(i, 1), noesc, htiempo, txtmsg, exito)
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, matport(i, 1), noesc, htiempo, 1 - nconf, exito)
       If exito Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, matport(i, 1), txtportfr, txttvar, noesc, htiempo, 0, 1 - nconf, 0, valor)
   Next i
   Call GenVaRMark2(fecha, txtport, txtportfr, matport, noesc, htiempo, nconf, txtmsg, exito)
   Call ObtenerResFondoPen(fecha, txtport, noesc, htiempo, matport)
   Call LeerPyGHistPortPos2(fecha, txtport, txtportfr, matport, noesc, htiempo)
End If
End Sub

Function DefinePortFP1()

ReDim matport1(1 To 25, 1 To 2) As String
matport1(1, 1) = "BANAMEX 2065"
matport1(2, 1) = "BANOBRAS 2065"
matport1(3, 1) = "BANORTE 2065"
matport1(4, 1) = "EVERCORE 2065"
matport1(5, 1) = "GBM 2065"
matport1(6, 1) = "SANTANDER 2065"
matport1(7, 1) = "VECTOR 2065"
matport1(8, 1) = "983"
matport1(9, 1) = "984"
matport1(10, 1) = "985"
matport1(11, 1) = "986"
matport1(12, 1) = "987"
matport1(13, 1) = "Capitales y SI 2065"
matport1(14, 1) = "CBICS Y PIC 2065"
matport1(15, 1) = "Papel Guber 2065"
matport1(16, 1) = "Papel Privado 2065"
matport1(17, 1) = "Reporto PG 2065"
matport1(18, 1) = "Reporto PP 2065"
matport1(19, 1) = "AAA 2065"
matport1(20, 1) = "AA 2065"
matport1(21, 1) = "A 2065"
matport1(22, 1) = "BBB+ 2065"
matport1(23, 1) = "Corto plazo 2065"
matport1(24, 1) = "NA 2065"
matport1(25, 1) = "Gobierno Fed 2065"

DefinePortFP1 = matport1
End Function

Function DefinePortFPA1(nofid)

ReDim matport1(1 To 30, 1 To 2) As String
matport1(1, 1) = "BI " & nofid
matport1(2, 1) = "LD " & nofid
matport1(3, 1) = "IM " & nofid
matport1(4, 1) = "IQ " & nofid
matport1(5, 1) = "IS " & nofid
matport1(6, 1) = "M " & nofid
matport1(7, 1) = "S " & nofid
matport1(8, 1) = "PI " & nofid
matport1(9, 1) = "2U " & nofid
matport1(10, 1) = "D2 " & nofid
matport1(11, 1) = "D8 " & nofid
matport1(12, 1) = "90 " & nofid
matport1(13, 1) = "91 " & nofid
matport1(14, 1) = "92 " & nofid
matport1(15, 1) = "93 " & nofid
matport1(16, 1) = "94 " & nofid
matport1(17, 1) = "95 " & nofid
matport1(18, 1) = "F " & nofid
matport1(19, 1) = "I " & nofid
matport1(20, 1) = "JE " & nofid
matport1(21, 1) = "JI " & nofid
matport1(22, 1) = "CD " & nofid
matport1(23, 1) = "F " & nofid
matport1(24, 1) = "1 " & nofid
matport1(25, 1) = "1A " & nofid
matport1(26, 1) = "1I " & nofid
matport1(27, 1) = "CF " & nofid
matport1(28, 1) = "41 " & nofid
matport1(29, 1) = "51 " & nofid
matport1(30, 1) = "52 " & nofid

matport1(1, 2) = "BI"
matport1(2, 2) = "LD"
matport1(3, 2) = "IM"
matport1(4, 2) = "IQ"
matport1(5, 2) = "IS"
matport1(6, 2) = "M"
matport1(7, 2) = "S"
matport1(8, 2) = "PI"
matport1(9, 2) = "2U"
matport1(10, 2) = "D2"
matport1(11, 2) = "D8"
matport1(12, 2) = "90"
matport1(13, 2) = "91"
matport1(14, 2) = "92"
matport1(15, 2) = "93"
matport1(16, 2) = "94"
matport1(17, 2) = "95"
matport1(18, 2) = "F"
matport1(19, 2) = "I"
matport1(20, 2) = "JE"
matport1(21, 2) = "JI"
matport1(22, 2) = "CD"
matport1(23, 2) = "F"
matport1(24, 2) = "1"
matport1(25, 2) = "1A"
matport1(26, 2) = "1I"
matport1(27, 2) = "CF"
matport1(28, 2) = "41"
matport1(29, 2) = "51"
matport1(30, 2) = "52"

DefinePortFPA1 = matport1
End Function


Function DefinePortFP2()

ReDim matport2(1 To 22, 1 To 2) As String
matport2(1, 1) = "ACTINVER 2160"
matport2(2, 1) = "BANAMEX 2160"
matport2(3, 1) = "BANOBRAS 2160"
matport2(4, 1) = "GBM 2160"
matport2(5, 1) = "VECTOR 2160"
matport2(6, 1) = "988"
matport2(7, 1) = "989"
matport2(8, 1) = "990"
matport2(9, 1) = "1111"

matport2(10, 1) = "Capitales y SI 2160"
matport2(11, 1) = "CBICS y PIC 2160"
matport2(12, 1) = "Papel Guber 2160"
matport2(13, 1) = "Papel Privado 2160"
matport2(14, 1) = "Reporto PG 2160"
matport2(15, 1) = "Reporto PP 2160"
matport2(16, 1) = "AAA 2160"
matport2(17, 1) = "AA 2160"
matport2(18, 1) = "A 2160"
matport2(19, 1) = "BBB+ 2160"
matport2(20, 1) = "Corto plazo 2160"
matport2(21, 1) = "NA 2160"
matport2(22, 1) = "Gobierno Fed 2160"

DefinePortFP2 = matport2
End Function

Sub ImportFlujosPIP(ByVal fecha As Date, ByVal dirarch As String)
Dim nomarch2 As String
Dim nomarch3 As String
Dim sihayarch2 As Boolean
Dim sihayarch3 As Boolean
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtfiltro As String
Dim txtinserta As String
Dim matpos() As propPosMD
Dim noreg As Long
Dim noreg2 As Long
Dim nocampos2 As Long
Dim nocampos As Long
Dim nocampos1 As Long
Dim i As Long
Dim j As Long
Dim noreg1 As Long
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim contar1 As Long
Dim contar2 As Long
Dim txtborra As String
Dim mata() As Variant

nomarch2 = dirarch & "\BanobrasFlujos" & Format(fecha, "yyyymmdd") & ".xls"
nomarch3 = dirarch & "\BanobrasFlujos" & Format(fecha, "yyyymmdd") & "_2.xls"
sihayarch2 = VerifAccesoArch(nomarch2)
sihayarch3 = VerifAccesoArch(nomarch3)
If sihayarch2 And sihayarch3 Then
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro = "SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
   txtfiltro = txtfiltro & " AND (TOPERACION =1 OR TOPERACION =4)"
   matpos = LeerBaseMD(txtfiltro)
   noreg = UBound(matpos, 1)
   ReDim mata(1 To noreg, 1 To 1) As Variant
   For i = 1 To noreg
      mata(i, 1) = matpos(i).cEmisionMD
   Next i
   mata = ObtFactUnicos(mata, 1)
   noreg = UBound(mata, 1)
   Set base1 = OpenDatabase(nomarch2, dbDriverNoPrompt, False, VersExcel)
   Set registros1 = base1.OpenRecordset("Sheet1$", dbOpenDynaset)
   registros1.MoveLast
   noreg1 = registros1.RecordCount
   nocampos1 = registros1.Fields.Count
   registros1.MoveFirst
   ReDim matb(1 To noreg1, 1 To nocampos1 + 1) As Variant
   For i = 1 To noreg1
       For j = 1 To nocampos1
           matb(i, j) = LeerTAccess(registros1, j - 1, i)
       Next j
       registros1.MoveNext
   Next i
   registros1.Close
   base1.Close
   For i = 1 To noreg1
       matb(i, nocampos1 + 1) = Trim(matb(i, 2)) & Trim(matb(i, 3)) & Trim(matb(i, 4))
   Next i
   Set base1 = OpenDatabase(nomarch3, dbDriverNoPrompt, False, VersExcel)
   Set registros1 = base1.OpenRecordset("Sheet1$", dbOpenDynaset)
   registros1.MoveLast
   noreg2 = registros1.RecordCount
   nocampos2 = registros1.Fields.Count
   registros1.MoveFirst
   ReDim matc(1 To noreg2, 1 To nocampos2 + 1) As Variant
   For i = 1 To noreg2
       For j = 1 To nocampos2
           matc(i, j) = LeerTAccess(registros1, j - 1, i)
       Next j
       registros1.MoveNext
   Next i
   registros1.Close
   base1.Close
   For i = 1 To noreg2
       matc(i, nocampos2 + 1) = Trim(matc(i, 2)) & Trim(matc(i, 3)) & Trim(matc(i, 4))
   Next i
   For i = 1 To noreg
       AvanceProc = i / noreg
       MensajeProc = "Buscando flujos de la emisión " & mata(i, 1) & " " & Format(AvanceProc, "##0.00 %")
       contar1 = 0
       ReDim matd(1 To 8, 1 To 1) As Variant
         For j = 1 To noreg1
             If mata(i, 1) = matb(j, nocampos1 + 1) And Val(Trim(matb(j, 14))) <> 0 And Val(Trim(matb(j, 16))) <> 0 Then
                contar1 = contar1 + 1
                ReDim Preserve matd(1 To 8, 1 To contar1) As Variant
                matd(1, contar1) = mata(i, 1)                  'clave de la emision
                matd(2, contar1) = fecha                       'fecha de registro
                matd(3, contar1) = matb(j, 12)                 'fecha de inicio
                matd(4, contar1) = matb(j, 13)                 'fecha final
                matd(5, contar1) = matb(j, 16)                 'saldo insoluto
        End If
    Next j
    If contar1 <> 0 Then
       matd = MTranV(matd)
       Call RutinaOrden(matd, 3, SRutOrden)
       For j = 1 To contar1
       If j <> contar1 Then
          matd(j, 6) = matd(j, 5) - matd(j + 1, 5)
       Else
          matd(j, 6) = matd(j, 5)
       End If
       matd(j, 7) = 0
       matd(j, 8) = matd(j, 4) - matd(j, 3)
       Next j
    End If
    contar2 = 0
    ReDim mate(1 To 8, 1 To 1) As Variant
    For j = 1 To noreg2
        If mata(i, 1) = matc(j, nocampos2 + 1) And Val(Trim(matc(j, 14))) <> 0 And Val(Trim(matc(j, 16))) <> 0 Then
           contar2 = contar2 + 1
           ReDim Preserve mate(1 To 8, 1 To contar2) As Variant
           mate(1, contar2) = mata(i, 1)                 'clave de la emision
           mate(2, contar2) = fecha                       'fecha de registro
           mate(3, contar2) = matc(j, 12)                 'fecha de inicio
           mate(4, contar2) = matc(j, 13)                 'fecha final
           mate(5, contar2) = matc(j, 16)                 'saldo insoluto
        End If
    Next j
    If contar2 <> 0 Then
       mate = MTranV(mate)
       Call RutinaOrden(mate, 3, SRutOrden)
       For j = 1 To contar2
       If j <> contar2 Then
          mate(j, 6) = mate(j, 5) - mate(j + 1, 5)
       Else
          mate(j, 6) = mate(j, 5)
       End If
       mate(j, 7) = 0
       mate(j, 8) = mate(j, 4) - mate(j, 3)
       Next j
    End If
    If contar1 <> 0 Then
       txtborra = "DELETE FROM " & TablaFlujosMD & " WHERE EMISION = '" & mata(i, 1) & "' and FREGISTRO = " & txtfecha
       ConAdo.Execute txtborra
       For j = 1 To contar1
           txtfecha1 = "to_date('" & Format(matd(j, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtfecha2 = "to_date('" & Format(matd(j, 3), "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtfecha3 = "to_date('" & Format(matd(j, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtinserta = "INSERT INTO " & TablaFlujosMD & " VALUES("
           txtinserta = txtinserta & "'" & matd(j, 1) & "',"          'emision
           txtinserta = txtinserta & txtfecha1 & ","                  'fecha de registro
           txtinserta = txtinserta & txtfecha2 & ","                  'fecha inicio
           txtinserta = txtinserta & txtfecha3 & ","                  'fecha final
           txtinserta = txtinserta & matd(j, 5) & ","                 'saldo
           txtinserta = txtinserta & matd(j, 6) & ","                 'AMORTIZACION
           txtinserta = txtinserta & matd(j, 7) & ","                 'tasa
           txtinserta = txtinserta & matd(j, 8) & ")"                 'plazo cupon
           ConAdo.Execute txtinserta
       Next j
    End If
    If contar2 <> 0 Then
       txtborra = "DELETE FROM " & TablaFlujosMD & " WHERE EMISION = '" & mata(i, 1) & "' AND FREGISTRO = " & txtfecha
       ConAdo.Execute txtborra
       For j = 1 To contar2
           txtfecha1 = "to_date('" & Format(mate(j, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtfecha2 = "to_date('" & Format(mate(j, 3), "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtfecha3 = "to_date('" & Format(mate(j, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtinserta = "INSERT INTO " & TablaFlujosMD & " VALUES("
           txtinserta = txtinserta & "'" & mate(j, 1) & "',"           'emision
           txtinserta = txtinserta & txtfecha1 & ","                  'fecha de registro
           txtinserta = txtinserta & txtfecha2 & ","                  'fecha inicio
           txtinserta = txtinserta & txtfecha3 & ","                  'fecha final
           txtinserta = txtinserta & mate(j, 5) & ","                 'saldo
           txtinserta = txtinserta & mate(j, 6) & ","                 'AMORTIZACION
           txtinserta = txtinserta & mate(j, 7) & ","                 'tasa
           txtinserta = txtinserta & mate(j, 8) & ")"                 'plazo cupon
           ConAdo.Execute txtinserta
       Next j
    End If
    
 Next i
Close #1
Else
 MsgBox "no existe los archivos de flujos"
End If

End Sub

Sub ObtenerResFondoPen(ByVal fecha As Date, ByVal txtport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByRef matport() As String)
Dim nomarch As String
Dim valor() As Double
Dim var As Double
Dim txtcadena As String
Dim i As Long
Dim exito As Boolean
nomarch = DirResVaR & "\Resultados " & txtport & " " & Format(fecha, "yyyymmdd") & ".txt"
nomarch = "d:\Resultados " & txtport & " " & Format(fecha, "yyyymmdd") & ".txt"
'CommonDialog1.FileName = nomarch
'CommonDialog1.ShowSave
'nomarch = CommonDialog1.FileName
Open nomarch For Output As #1
Print #1, "PORTAFOLIO" & Chr(9) & "VALUACION" & Chr(9) & "CVAR"
valor = LeerResValPort(fecha, txtport, txtport, 1)
var = LeerResVaR(fecha, txtport, "Normal", txtport, noesc, htiempo, 0, 0.03, 0, "CVARH", exito)
txtcadena = txtport & Chr(9) & valor(1) & Chr(9) & var
Print #1, txtcadena
For i = 1 To UBound(matport, 1)
    valor = LeerResValPort(fecha, txtport, matport(i, 1), 1)
    var = LeerResVaR(fecha, txtport, "Normal", matport(i, 1), noesc, htiempo, 0, 0.03, 0, "CVARH", exito)
    If UBound(valor, 1) <> 0 Then
       txtcadena = matport(i, 1) & Chr(9) & valor(1) & Chr(9) & var
    Else
       txtcadena = matport(i, 1) & Chr(9) & 0 & Chr(9) & var
    End If
    Print #1, txtcadena
Next i
Close #1
End Sub

Sub LeerResultadosSimPos(ByVal fecha As Date, ByVal txtnompos As String, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim nomarch As String
Dim valor() As Double
Dim var As Double
Dim txtcadena As String
Dim i As Long
Dim noreg As Long
Dim exito As Boolean
Dim coperacion As String
Dim cposicion As Long
Dim notitulos As Double
Dim valor1 As Double
Dim val_pip_s As Double
Dim valactiva As Double
Dim valpasiva As Double
Dim duract As Double
Dim durpas As Double
Dim dv01_act As Double
Dim dv01_pas As Double
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
nomarch = DirResVaR & "\Resultados sim " & txtnompos & " " & Format(fecha, "yyyy-mm-dd") & ".txt"
'CommonDialog1.FileName = nomarch
'CommonDialog1.ShowSave
'nomarch = CommonDialog1.FileName
'Open nomarch For Output As #1
Print #1, "Clave de operacion" & Chr(9) & "Clave de posicion" & Chr(9) & "No de titulos" & Chr(9) & "Valuacion" & Chr(9) & "Val pata activa" & Chr(9) & "Val pata pasiva" & Chr(9) & "Duracion act" & Chr(9) & "Duracion pas" & Chr(9) & "DV01 act" & Chr(9) & "DV01 pas"
txtfiltro2 = "SELECT * FROM " & TablaValPos & " WHERE NOMPOS = '" & txtnompos & "' AND FECHAP = " & txtfecha
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       coperacion = rmesa.Fields("COPERACION").value
       cposicion = rmesa.Fields("CPOSICION").value
       notitulos = rmesa.Fields("NO_TITULOS_").value
       valor1 = rmesa.Fields("P_SUCIO").value
       valactiva = rmesa.Fields("VAL_ACT_S").value
       valpasiva = rmesa.Fields("VAL_PAS_S").value
       val_pip_s = rmesa.Fields("VAL_PIP_S").value
       duract = rmesa.Fields("DUR_ACT").value
       durpas = rmesa.Fields("DUR_PAS").value
       dv01_act = rmesa.Fields("DV01_ACT").value
       dv01_pas = rmesa.Fields("DV01_PAS").value
       rmesa.MoveNext
       txtcadena = coperacion & Chr(9) & cposicion & Chr(9) & notitulos & Chr(9) & valor1 & Chr(9) & val_pip_s & Chr(9) & valactiva & Chr(9) & valpasiva & Chr(9) & duract & Chr(9) & durpas & Chr(9) & dv01_act & Chr(9) & dv01_pas
       Print #1, txtcadena
   Next i
   rmesa.Close
End If
Close #1
Call LeerPyGPosSim(fecha, txtnompos, "Normal", noesc, htiempo)

End Sub


Sub CrearPortEmxContrap(ByVal fecha As Date, ByRef mata() As String, ByVal opcion As Integer)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtport As String
Dim cposicion As Integer
Dim coperacion As String
Dim txtcadena As String
Dim txtborra As String
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim tipopos As Integer
Dim contar As Long
Dim rmesa As New ADODB.recordset
tipopos = 1
contar = 0
ReDim mata(1 To 1) As String
For i = 1 To UBound(MatContrapartes, 1)
    If opcion = 1 Then
       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro2 = "SELECT " & TablaPosMD & ".CPOSICION," & TablaPosMD & ".COPERACION"
       txtfiltro2 = txtfiltro2 & " FROM " & TablaPosMD & " JOIN " & PrefijoBD & TablaEmxContrap & " ON "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".EMISION = " & PrefijoBD & TablaEmxContrap & ".EMISION WHERE "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".FECHAREG = " & txtfecha & " AND "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".TIPOPOS = " & tipopos & " AND "
       txtfiltro2 = txtfiltro2 & "(" & TablaPosMD & ".CPOSICION = " & ClavePosMD
       txtfiltro2 = txtfiltro2 & " OR " & TablaPosMD & ".CPOSICION = " & ClavePosTeso & ") AND "
       txtfiltro2 = txtfiltro2 & PrefijoBD & TablaEmxContrap & ".IDCONTRAP = " & MatContrapartes(i, 1)
    ElseIf opcion = 2 Then
       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro2 = "SELECT " & TablaPosMD & ".CPOSICION," & TablaPosMD & ".COPERACION"
       txtfiltro2 = txtfiltro2 & " FROM " & TablaPosMD & " JOIN " & PrefijoBD & TablaEmxContrap & " ON "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".EMISION = " & PrefijoBD & TablaEmxContrap & ".EMISION WHERE "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".FECHAREG = " & txtfecha & " AND "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".TIPOPOS = " & tipopos & " AND "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".CPOSICION = " & ClavePosPIDV & " AND "
       txtfiltro2 = txtfiltro2 & PrefijoBD & TablaEmxContrap & ".IDCONTRAP = " & MatContrapartes(i, 1)
    ElseIf opcion = 3 Then
       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro2 = "SELECT " & TablaPosMD & ".CPOSICION," & TablaPosMD & ".COPERACION"
       txtfiltro2 = txtfiltro2 & " FROM " & TablaPosMD & " JOIN " & PrefijoBD & TablaEmxContrap & " ON "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".EMISION = " & PrefijoBD & TablaEmxContrap & ".EMISION WHERE "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".FECHAREG = " & txtfecha & " AND "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".TIPOPOS = " & tipopos & " AND "
       txtfiltro2 = txtfiltro2 & TablaPosMD & ".CPOSICION = " & ClavePosPICV & " AND "
       txtfiltro2 = txtfiltro2 & PrefijoBD & TablaEmxContrap & ".IDCONTRAP = " & MatContrapartes(i, 1)
    End If
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       If opcion = 1 Then
          txtport = "Em x contrap " & MatContrapartes(i, 1) & " MD"
       ElseIf opcion = 2 Then
          txtport = "Em x contrap " & MatContrapartes(i, 1) & " PIDV"
       ElseIf opcion = 3 Then
          txtport = "Em x contrap " & MatContrapartes(i, 1) & " PICV"
       End If
       contar = contar + 1
       ReDim Preserve mata(1 To contar) As String
       mata(contar) = MatContrapartes(i, 1)
       rmesa.Open txtfiltro2, ConAdo
       txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "'"
       ConAdo.Execute txtborra
       For j = 1 To noreg
           cposicion = rmesa.Fields("CPOSICION")
           coperacion = rmesa.Fields("COPERACION")
           txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
           txtcadena = txtcadena & txtfecha & ","
           txtcadena = txtcadena & "'" & txtport & "',"
           txtcadena = txtcadena & tipopos & ","
           txtcadena = txtcadena & txtfecha & ","
           txtcadena = txtcadena & "'Real',"
           txtcadena = txtcadena & "'000000',"
           txtcadena = txtcadena & cposicion & ","
           txtcadena = txtcadena & "'" & coperacion & "')"
           ConAdo.Execute txtcadena
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
Next i
End Sub

