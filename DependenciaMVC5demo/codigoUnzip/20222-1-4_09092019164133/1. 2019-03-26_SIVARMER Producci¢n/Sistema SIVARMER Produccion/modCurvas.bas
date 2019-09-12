Attribute VB_Name = "modCurvas"
Option Explicit

Sub ImpNdCurvas(ByVal fecha As Date, ByVal txtcurva As String, ByRef noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim siaccion As Boolean
Dim i As Integer
Dim noreg1 As Integer

exito = True
txtmsg = "El proceso finalizo correctamente"
'se lee el archivo de curvas del proveedor de precios
noreg = 0
'primero se verifica si la fecha es un dia habil
siaccion = NoLabMX(fecha)
If Not siaccion Then
  'abre el archivo de curvas del proveedor de precios PIP
    If EsVariableVacia(txtcurva) Then
       For i = 1 To UBound(MatCatCurvas, 1)
          Call LeerNodosCurva(fecha, MatCatCurvas(i, 1), MatCatCurvas(i, 2), noreg1, exito1)
          noreg = noreg + noreg1
          exito = exito And exito1
          If Not exito1 Then txtmsg = "No se encontro los datos de la curva " & MatCatCurvas(i, 2)
       Next i
       MensajeProc = "Se guardaron en " & TablaFRiesgoO & " " & noreg & " registros del archivo de curvas del " & fecha
    Else
       For i = 1 To UBound(MatCatCurvas, 1)
           If MatCatCurvas(i, 2) = txtcurva Then
              Call LeerNodosCurva(fecha, MatCatCurvas(i, 1), MatCatCurvas(i, 2), noreg1, exito1)
              noreg = noreg + noreg1
              exito = exito And exito1
              Exit For
           End If
       Next i
       MensajeProc = "Se guardaron en " & TablaFRiesgoO & " " & noreg & " registros del archivo de curvas del " & fecha
      
   End If
Else
  exito = False
End If
End Sub

Sub LeerNodosCurva(ByVal fecha As Date, ByVal idcurva As Integer, ByVal txtcurva As String, ByVal noreg1 As Long, ByRef exito As Boolean)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim valor As String
Dim matc() As String
Dim i As Long
Dim matv() As Double
Dim txtborra As String
Dim contar As Integer
Dim indice As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

'rutina para importar datos de PIP
'se completan algunos datos para que la curva no
'tenga huecos
'
'primero se carga la curva completa
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaCurvas & " WHERE FECHA = " & txtfecha & " AND IDCURVA = " & idcurva
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      valor = rmesa.Fields(2).GetChunk(rmesa.Fields(2).ActualSize)
      rmesa.Close
      matc = EncontrarSubCadenas(valor, ",")
      If UBound(matc, 1) = 12000 Then
      ReDim matv(1 To UBound(matc, 1)) As Double
      For i = 1 To UBound(matc, 1)
          matv(i) = CDbl(matc(i))
      Next i
      txtborra = "DELETE FROM " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtcurva & "'"
      ConAdo.Execute txtborra
      contar = 0
      For i = 1 To UBound(MatNodosCurvas, 1)
          If MatNodosCurvas(i, 1) = txtcurva Then
             indice = CLng(fecha) & txtcurva & Format(MatNodosCurvas(i, 2), "0000000")
             txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
             txtcadena = txtcadena & txtfecha & ","
             txtcadena = txtcadena & "'" & txtcurva & "',"
             txtcadena = txtcadena & MatNodosCurvas(i, 2) & ","
             txtcadena = txtcadena & matv(MatNodosCurvas(i, 2)) & ","
             txtcadena = txtcadena & "'" & indice & "')"
             ConAdo.Execute txtcadena
             contar = contar + 1
          End If
          AvanceProc = i / UBound(MatNodosCurvas, 1)
          MensajeProc = "Guardando el nodo de " & txtcurva & " del " & fecha & " " & Format(AvanceProc, "##0.00 %")
          DoEvents
      Next i
      noreg1 = contar
      exito = True
   Else
    noreg1 = 0
    exito = False
   End If
Else
   noreg1 = 0
   exito = False
End If
End Sub

Sub LeerNodosCurvas()
'se cargan los plazos que se van a leer en la rutina de importacion de
'datos del archivo de curvas de proveedor de precios
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Integer
Dim rmesa As New ADODB.recordset
'====================================================
txtfiltro2 = "SELECT * from " & PrefijoBD & TablaNodosCurvas & " ORDER BY CURVA, NODO"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim MatNodosCurvas(1 To noreg, 1 To 2) As Variant
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
        MatNodosCurvas(i, 1) = rmesa.Fields("CURVA")
        MatNodosCurvas(i, 2) = rmesa.Fields("NODO")
        rmesa.MoveNext
   Next i
   rmesa.Close
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub BorrarDiasFestO()
Dim txtfiltro As String
Dim noreg As Integer
Dim i As Integer
Dim mata() As Variant
Dim txtfecha As String
Dim rmesa As New ADODB.recordset

MsgBox "Esta rutina borra los días que no deben de ir en la tabla de factores de riesgo"
'se leen los dias festivos del calendario

'se procede a determinar que dia es fin de semana y se procede a borrarlo
txtfiltro = "select count(distinct FECHA) from " & TablaFRiesgoO & ""
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
txtfiltro = "select FECHA from " & TablaFRiesgoO & " GROUP BY FECHA ORDER BY FECHA"
rmesa.Open txtfiltro, ConAdo
ReDim mata(1 To noreg) As Variant
For i = 1 To noreg
 mata(i) = rmesa.Fields("FECHA")
 rmesa.MoveNext
Next i
rmesa.Close
'borra los dias sabados y domingos de la tabla de datos
For i = 1 To noreg
If Weekday(mata(i)) = 1 Or Weekday(mata(i)) = 7 Then
   txtfecha = "TO_DATE('" & Format(mata(i), "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro = "DELETE FROM " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha
   ConAdo.Execute txtfiltro
End If
Next i
MsgBox "Se borraron registros de dias festivos de la tabla " & TablaFRiesgoO
'se borraron los factores de riesgo no validos
End Sub

Sub ObtenerTIIE(ByVal fecha As Date, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim siesfv As Boolean
Dim sieshabil As Boolean
Dim fechax As Date
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim txtconcepto As String
Dim txtfecha1 As String
Dim indice1 As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

siesfv = EsFechaVaR(fecha)
sieshabil = Not NoLabMX(fecha)
fechax = PBD1(fecha, 1, "MX")
If fechax <> 0 And sieshabil Then
   txtfecha = "TO_DATE('" & Format(fechax, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro2 = "SELECT * FROM " & TablaFRiesgoO & " WHERE FECHA =  " & txtfecha & " AND CONCEPTO = 'TIIE28 PIP'"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 2) As Variant
      rmesa.MoveFirst
      For i = 1 To noreg
          mata(i, 1) = fecha
          mata(i, 2) = rmesa.Fields(3)
          rmesa.MoveNext
      Next i
      rmesa.Close
      txtconcepto = "TIIE 28"
      For i = 1 To noreg
          txtfecha1 = "TO_DATE('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          ConAdo.Execute "DELETE FROM " & TablaFRiesgoO & " WHERE fecha = " & txtfecha1 & " AND CONCEPTO = '" & txtconcepto & "'"
          indice1 = CLng(mata(i, 1)) & txtconcepto & "0000000"
          txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
          txtcadena = txtcadena & txtfecha1 & ","
          txtcadena = txtcadena & "'" & txtconcepto & "',"
          txtcadena = txtcadena & "0,"
          txtcadena = txtcadena & mata(i, 2) & ","
          txtcadena = txtcadena & "'" & indice1 & "')"
          ConAdo.Execute txtcadena
      Next i
      exito1 = True
   Else
      exito1 = False
   End If
   txtfecha = "TO_DATE('" & Format(fechax, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro2 = "SELECT * FROM " & TablaFRiesgoO & " WHERE FECHA =  " & txtfecha & " AND CONCEPTO = 'TIIE91 PIP'"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 2) As Variant
      rmesa.MoveFirst
      For i = 1 To noreg
          mata(i, 1) = fecha
          mata(i, 2) = rmesa.Fields(3)
          rmesa.MoveNext
      Next i
      rmesa.Close
      txtconcepto = "TIIE 91"
      For i = 1 To noreg
          txtfecha1 = "TO_DATE('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          ConAdo.Execute "DELETE FROM " & TablaFRiesgoO & " WHERE fecha = " & txtfecha1 & " AND CONCEPTO = '" & txtconcepto & "'"
          indice1 = CLng(mata(i, 1)) & txtconcepto & "0000000"
          txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
          txtcadena = txtcadena & txtfecha1 & ","
          txtcadena = txtcadena & "'" & txtconcepto & "',"
          txtcadena = txtcadena & "0,"
          txtcadena = txtcadena & mata(i, 2) & ","
          txtcadena = txtcadena & "'" & indice1 & "')"
          ConAdo.Execute txtcadena
      Next i
      exito2 = True
   Else
      exito2 = False
   End If
   If exito1 And exito2 Then
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      exito = False
   End If
End If

End Sub

Sub ObtenerLibort2(ByVal fecha As Date, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim siesfv As Boolean
Dim sieshabil As Boolean
Dim fechax As Date
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim txtconcepto As String
Dim txtfecha1 As String
Dim indice1 As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

siesfv = EsFechaVaR(fecha)
sieshabil = Not NoLabMX(fecha)
fechax = PBD1(fecha, 2, "MX")
If fechax <> 0 And sieshabil Then
   txtfecha = "TO_DATE('" & Format(fechax, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro2 = "SELECT * FROM " & TablaFRiesgoO & " WHERE FECHA =  " & txtfecha & " AND CONCEPTO = 'LIBOR 3M PIP'"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      ReDim mata(1 To noreg, 1 To 2) As Variant
      rmesa.MoveFirst
      For i = 1 To noreg
          mata(i, 1) = fecha
          mata(i, 2) = rmesa.Fields(3)
          rmesa.MoveNext
      Next i
      rmesa.Close
      txtconcepto = "LIBOR 3M t-2"
      For i = 1 To noreg
          txtfecha1 = "TO_DATE('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          ConAdo.Execute "DELETE FROM " & TablaFRiesgoO & " WHERE fecha = " & txtfecha1 & " AND CONCEPTO = '" & txtconcepto & "'"
          indice1 = CLng(mata(i, 1)) & txtconcepto & "0000000"
          txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
          txtcadena = txtcadena & txtfecha1 & ","
          txtcadena = txtcadena & "'" & txtconcepto & "',"
          txtcadena = txtcadena & "0,"
          txtcadena = txtcadena & mata(i, 2) & ","
          txtcadena = txtcadena & "'" & indice1 & "')"
          ConAdo.Execute txtcadena
      Next i
      exito1 = True
   Else
      exito1 = False
   End If
   txtfecha = "TO_DATE('" & Format(fechax, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro2 = "SELECT * FROM " & TablaFRiesgoO & " WHERE FECHA =  " & txtfecha & " AND CONCEPTO = 'LIBOR 6M PIP'"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      ReDim mata(1 To noreg, 1 To 2) As Variant
      rmesa.MoveFirst
      For i = 1 To noreg
          mata(i, 1) = fecha
          mata(i, 2) = rmesa.Fields(3)
          rmesa.MoveNext
      Next i
      rmesa.Close
      txtconcepto = "LIBOR 6M t-2"
      For i = 1 To noreg
          txtfecha1 = "TO_DATE('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          ConAdo.Execute "DELETE FROM " & TablaFRiesgoO & " WHERE fecha = " & txtfecha1 & " AND CONCEPTO = '" & txtconcepto & "'"
          indice1 = CLng(mata(i, 1)) & txtconcepto & "0000000"
          txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
          txtcadena = txtcadena & txtfecha1 & ","
          txtcadena = txtcadena & "'" & txtconcepto & "',"
          txtcadena = txtcadena & "0,"
          txtcadena = txtcadena & mata(i, 2) & ","
          txtcadena = txtcadena & "'" & indice1 & "')"
          ConAdo.Execute txtcadena
      Next i
      exito2 = True
   Else
      exito2 = False
   End If
   If exito1 And exito2 Then
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      exito = False
   End If
End If

End Sub


Sub ObtTRefPIPO(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef nr1 As Long, ByRef txtmsg As String, exito As Boolean)
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim j As Integer
Dim fecha As Date
Dim txtfecha As String
Dim txtindice As String
Dim txtcadena As String
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim contar As Integer
Dim dif As Double

   exito = True
   noreg = 2
   ReDim mata(1 To noreg, 1 To 3) As Variant
   mata(1, 1) = "IM"
   mata(1, 2) = "TR BPAG28"
   mata(2, 1) = "IQ"
   mata(2, 2) = "TR BPAG91"
   fecha = fecha1
   Do While fecha <= fecha2
      matvp = LeerVPrecios(fecha, mindvp)
      noreg1 = UBound(matvp, 1)
      If noreg1 <> 0 Then
         For i = 1 To noreg
             contar = 0
             dif = 0
             For j = 1 To noreg1
                 If matvp(j).tv = mata(i, 1) Then
                    dif = dif + (matvp(j).yield / 100 - matvp(j).tasa_st / 100)
                    contar = contar + 1
                 End If
             Next j
             If contar <> 0 Then
                dif = dif / contar
             Else
                dif = 0
             End If
             mata(i, 3) = dif
         Next i
         For i = 1 To noreg
             If mata(i, 3) <> 0 Then
              txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
              txtcadena = "DELETE FROM " & TablaFRiesgoO & " WHERE CONCEPTO = '" & mata(i, 2) & "' AND PLAZO = 0 AND FECHA = " & txtfecha
              ConAdo.Execute txtcadena
              txtindice = CLng(fecha) & mata(i, 2) & "0000000"
              txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
              txtcadena = txtcadena & txtfecha & ","
              txtcadena = txtcadena & "'" & mata(i, 2) & "',"
              txtcadena = txtcadena & "0,"
              txtcadena = txtcadena & mata(i, 3) & ","
              txtcadena = txtcadena & "'" & txtindice & "')"
              ConAdo.Execute txtcadena
             End If
         Next i
         txtmsg = "El proceso finalizo correctamente"
      Else
         txtmsg = "No se encontraron los datos"
         exito = False
      End If
      fecha = fecha + 1
   Loop

End Sub

Sub ObtTCambioMPIP(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal nr1 As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim i As Integer
Dim j As Integer
Dim txtcadena As String
Dim txtfiltro As String
Dim noreg2 As Integer
Dim noreg3 As Integer
Dim indice As String
Dim fecha As Date
Dim txtfecha As String
Dim rmesa As New ADODB.recordset

'los tipos de cambio de banxico mensuales
exito = True
txtfecha1 = "TO_DATE('" & Format(fecha1 - 50, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "TO_DATE('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
noreg2 = 3
ReDim matb(1 To noreg2, 1 To 3) As Variant
matb(1, 1) = "*CBMXPEURBEUR"
matb(1, 2) = "EURO BM"
matb(2, 1) = "*CBMXPJPYBJPY"
matb(2, 2) = "YEN BM"
matb(3, 1) = "*CBMXPUSDBUSD"
matb(3, 2) = "DOLAR BM"
For i = 1 To noreg2
    txtcadena = "SELECT * FROM " & TablaVecPrecios & " WHERE CLAVE_EMISION = '" & matb(i, 1) & "' AND FECHA >=  " & txtfecha1 & " AND FECHA <= " & txtfecha2 & "  ORDER BY FECHA"
    txtfiltro = "SELECT COUNT(*) FROM(" & txtcadena & ")"
    rmesa.Open txtfiltro, ConAdo
    noreg3 = rmesa.Fields(0)
    rmesa.Close
    If noreg3 <> 0 Then
       rmesa.Open txtcadena, ConAdo
   ReDim matc(1 To noreg3, 1 To 2)
       rmesa.MoveFirst
       For j = 1 To noreg3
           matc(j, 1) = rmesa.Fields(0)
           matc(j, 2) = rmesa.Fields(5)
           rmesa.MoveNext
       Next j
       rmesa.Close
       fecha = fecha1
       Do While fecha <= fecha2
          indice = BuscarValorArray(fecha, MatFechasFR, 1)
          If indice <> 0 Then
             matb(i, 3) = 0
             If fecha < matc(1, 1) Then
                matb(i, 3) = 0
             ElseIf fecha >= matc(noreg3, 1) Then
                matb(i, 3) = matc(noreg3, 2)
             Else
                For j = 1 To noreg3
                    If fecha >= matc(j, 1) And fecha < matc(j + 1, 1) Then
                       matb(i, 3) = matc(j, 2)
                       Exit For
                    End If
                Next j
             End If
             If matb(i, 3) <> 0 Then
                txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
                txtcadena = "DELETE FROM " & TablaFRiesgoO & " WHERE CONCEPTO = '" & matb(i, 2) & "' AND FECHA = " & txtfecha
                ConAdo.Execute txtcadena
                indice = CLng(fecha) & matb(i, 2) & "0000000"
                txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
                txtcadena = txtcadena & txtfecha & ","
                txtcadena = txtcadena & "'" & matb(i, 2) & "',"
                txtcadena = txtcadena & "0,"
                txtcadena = txtcadena & matb(i, 3) & ","
                txtcadena = txtcadena & "'" & indice & "')"
                ConAdo.Execute txtcadena
             End If
             txtmsg = "El proceso finalizo correctamente"
      End If
   fecha = fecha + 1
   Loop
End If


Next i

End Sub

Sub ImportarFactVP(ByVal fecha As Date, ByVal texto As String, ByRef nr1 As Long, ByRef txtmsg As String, ByRef exito As Boolean)
'rutina que lee los indices del vector de precios
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim noreg As Integer
Dim noregvp As Integer
Dim contar As Integer
Dim i As Integer
Dim indice As Long
Dim indice1 As Long
Dim txtfecha As String
Dim txtcadena As String

ReDim matb(1 To 6, 0 To 0) As Variant
   exito = True
   txtmsg = ""
   matvp = LeerVPrecios(fecha, mindvp)
   noreg = UBound(MatIndVPrecios, 1)
   noregvp = UBound(matvp, 1)
   If noregvp <> 0 Then
      contar = 0
      ReDim Preserve matb(1 To 6, 0 To 0) As Variant
      If EsVariableVacia(texto) Then
         For i = 1 To noreg
             If MatIndVPrecios(i, 1) <= fecha And IsNull(MatIndVPrecios(i, 2)) Then
                indice = BuscarValorArray(MatIndVPrecios(i, 6), mindvp, 1)
                If indice <> 0 Then
                   indice1 = mindvp(indice, 2)
                   contar = contar + 1
                   ReDim Preserve matb(1 To 6, 0 To contar) As Variant
                   matb(1, contar) = matvp(indice1).c_emision
                   matb(2, contar) = MatIndVPrecios(i, 3)    'nombre de factor
                   matb(3, contar) = MatIndVPrecios(i, 4)    'tipo de factor
                   matb(4, contar) = MatIndVPrecios(i, 5)    'columna
                   matb(5, contar) = fecha
                   If MatIndVPrecios(i, 4) = "INDICE" Then
                      If MatIndVPrecios(i, 5) = 7 Then
                         matb(6, contar) = matvp(indice1).psucio
                      ElseIf MatIndVPrecios(i, 5) = 17 Then
                         matb(6, contar) = matvp(indice1).yield
                      End If
                   Else
                      If MatIndVPrecios(i, 5) = 7 Then
                         matb(6, contar) = matvp(indice1).psucio / 100
                      ElseIf MatIndVPrecios(i, 5) = 17 Then
                         matb(6, contar) = matvp(indice1).yield / 100
                      End If
                   End If
                Else
                   MensajeProc = "Falta el factor " & MatIndVPrecios(i, 3)
                   txtmsg = txtmsg & "," & MensajeProc
                   exito = False
                End If
             ElseIf MatIndVPrecios(i, 1) <= fecha And fecha < MatIndVPrecios(i, 2) Then
                indice = BuscarValorArray(MatIndVPrecios(i, 6), mindvp, 1)
                If indice <> 0 Then
                   indice1 = mindvp(indice, 2)
                   contar = contar + 1
                    ReDim Preserve matb(1 To 6, 0 To contar) As Variant
                   matb(1, contar) = matvp(indice1).c_emision
                   matb(2, contar) = MatIndVPrecios(i, 3)    'nombre de factor
                   matb(3, contar) = MatIndVPrecios(i, 4)    'tipo de factor
                   matb(4, contar) = MatIndVPrecios(i, 5)    'columna
                   matb(5, contar) = fecha
                   If MatIndVPrecios(i, 4) = "INDICE" Then
                      matb(6, contar) = matvp(indice1).psucio
                   Else
                      matb(6, contar) = matvp(indice1).yield / 100
                   End If
                Else
                    MensajeProc = "Falta el factor " & MatIndVPrecios(i, 3)
                    txtmsg = txtmsg & "," & MensajeProc
                    exito = False
                End If
             End If
             AvanceProc = i / noreg
             MensajeProc = "Buscando la Informacion de " & MatIndVPrecios(i, 3) & " " & Format(AvanceProc, "##0.00 %")
             DoEvents
         Next i
      Else
         For i = 1 To noreg
             If MatIndVPrecios(i, 3) = texto Then
                If MatIndVPrecios(i, 1) <= fecha And IsNull(MatIndVPrecios(i, 2)) Then
                   indice1 = BuscarValorArray(MatIndVPrecios(i, 6), matvp, 18)
                   If indice1 <> 0 Then
                      contar = contar + 1
                      ReDim Preserve matb(1 To 6, 0 To contar) As Variant
                      matb(1, contar) = matvp(indice1, 22)
                      matb(2, contar) = MatIndVPrecios(i, 3)    'nombre de factor
                      matb(3, contar) = MatIndVPrecios(i, 4)    'tipo de factor
                      matb(4, contar) = MatIndVPrecios(i, 5)    'columna
                      matb(5, contar) = fecha
                      If MatIndVPrecios(i, 4) = "INDICE" Then
                         matb(6, contar) = matvp(indice1, MatIndVPrecios(i, 5))
                      Else
                         matb(6, contar) = matvp(indice1, MatIndVPrecios(i, 5)) / 100
                      End If
                    Else
                      MensajeProc = "Falta el factor " & MatIndVPrecios(i, 3)
                      txtmsg = txtmsg & "," & MensajeProc
                      exito = False
                    End If
                ElseIf MatIndVPrecios(i, 1) <= fecha And fecha < MatIndVPrecios(i, 2) Then
                    indice1 = BuscarValorArray(MatIndVPrecios(i, 6), matvp, 18)
                    If indice1 <> 0 Then
                       contar = contar + 1
                       ReDim Preserve matb(1 To 6, 0 To contar) As Variant
                       matb(1, contar) = matvp(indice1, 22)
                       matb(2, contar) = MatIndVPrecios(i, 3)    'nombre de factor
                       matb(3, contar) = MatIndVPrecios(i, 4)    'tipo de factor
                       matb(4, contar) = MatIndVPrecios(i, 5)    'columna
                       matb(5, contar) = fecha
                       If MatIndVPrecios(i, 4) = "INDICE" Then
                          matb(6, contar) = matvp(indice1, MatIndVPrecios(i, 5))
                       Else
                          matb(6, contar) = matvp(indice1, MatIndVPrecios(i, 5)) / 100
                       End If
                    Else
                       MensajeProc = "Falta el factor " & MatIndVPrecios(i, 3)
                       txtmsg = txtmsg & "," & MensajeProc
                       exito = False
                    End If
                End If
                 AvanceProc = i / noreg
                 MensajeProc = "Buscando la Informacion " & Format(AvanceProc, "##0.00 %")
                 DoEvents
             End If
         Next i
      End If

         noreg = UBound(matb, 2)
         For i = 1 To noreg
             txtfecha = "TO_DATE('" & Format(matb(5, i), "dd/mm/yyyy") & "','dd/mm/yyyy')"
             txtcadena = "DELETE FROM " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & matb(2, i) & "' AND PLAZO = 0"
             ConAdo.Execute txtcadena
             txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
             txtcadena = txtcadena & txtfecha & ",'"                 'fecha
             txtcadena = txtcadena & matb(2, i) & "',"               'concepto
             txtcadena = txtcadena & 0 & ","                         'plazo
             txtcadena = txtcadena & matb(6, i) & ","                'valor
             txtcadena = txtcadena & "'" & CLng(matb(5, i)) & Trim(matb(2, i)) & Trim(Format(0, "0000000")) & "')"
             ConAdo.Execute txtcadena
             MensajeProc = "Guardando " & matb(2, i) & " del " & matb(5, i)
             DoEvents
         Next i
         MensajeProc = "Se actualizaron datos del vector de precios en la tabla " & TablaFRiesgoO & " de la fecha " & fecha
   Else
      MensajeProc = "No hay vector de precios"
      exito = False
      txtmsg = MensajeProc
 End If
 If exito Then txtmsg = "El proceso finalizo correctamente"
nr1 = noreg
End Sub



