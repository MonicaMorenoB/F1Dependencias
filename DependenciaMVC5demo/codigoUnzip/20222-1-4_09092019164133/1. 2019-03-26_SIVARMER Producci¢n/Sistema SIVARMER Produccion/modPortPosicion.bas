Attribute VB_Name = "modPortPosicion"
Option Explicit

Function TraducirCadenaSQL(ByVal sqltexto As String, ByVal txtfechareg As String, ByVal tipopos As Integer)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_tipopos", tipopos)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_txtfechareg", txtfechareg)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_TablaPosMD", TablaPosMD)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_TablaPosDiv", TablaPosDiv)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_TablaPosSwaps", TablaPosSwaps)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_TablaPosFwd", TablaPosFwd)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_TablaCatContrap", PrefijoBD & TablaContrapartes)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_TablaPosDeuda", TablaPosDeuda)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_ClavePosMD", ClavePosMD)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_ClavePosTeso", ClavePosTeso)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_ClavePosMC", ClavePosMC)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_ClavePosDeriv", ClavePosDeriv)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_ClavePosPIDV", ClavePosPIDV)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_ClavePosPenMD", ClavePosPenMD)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_ClavePosPICV", ClavePosPICV)
sqltexto = ReemplazaCadenaTexto(sqltexto, "sql_ClavePosPID", ClavePosPID)
TraducirCadenaSQL = sqltexto
End Function

Sub ProcGenPortafolios(ByVal fecha As Date, ByVal id_subproc1 As Integer, ByVal id_subproc2 As Integer, ByVal id_subproc3 As Integer, ByVal opcion As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim txtborra As String
Dim txtfecha As String
Dim i As Integer
Dim txttabla As String
txttabla = DetermTablaSubproc(opcion)

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & txttabla & " WHERE FECHAP = " & txtfecha
txtborra = txtborra & " AND ID_SUBPROCESO = " & id_subproc1
ConAdo.Execute txtborra
exito = True
'se generan todos los subportafolios definidos en el catalogo de portafolios
For i = 1 To UBound(MatSQLPort, 1)
    Call GenSubGenPort(fecha, i, id_subproc1, opcion, txtmsg, exito1)
    exito = exito And exito1
    If exito1 = False Then
       MsgBox "alguna cadena sql esta ma construida"
       Exit Sub
    End If
Next i
'se construye el portafolio de estructurales y relacionados
Call DeterminaSubportEstructural(fecha)
Call GenPortPIDVDeriv(fecha)
Call GenPortDerivPIDV(fecha)
Call GenPortEmisionyPos(fecha)
'la posicion de derivados por contraparte
Call DeterminaPosDerivContrap(fecha, 1, id_subproc2, opcion)
'la posicion de swaps por contraparte
Call GeneraPosSwapsxContrap(fecha, 1, id_subproc3, opcion)
Call DeterminaContrapPosFwd(fecha, "Fwds Contrap", 1)
DoEvents

End Sub

Sub GenPortEmisionyPos(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfechar As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim noreg1 As Integer
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim txtport As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = " SELECT CPOSICION,C_EMISION,TOPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND TIPOPOS = 1"
txtfiltro2 = txtfiltro2 & "  GROUP BY CPOSICION,C_EMISION,TOPERACION ORDER BY C_EMISION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matem(1 To noreg, 1 To 3) As String
   For i = 1 To noreg
       matem(i, 1) = rmesa.Fields("C_EMISION")
       matem(i, 2) = rmesa.Fields("CPOSICION")
       matem(i, 3) = rmesa.Fields("TOPERACION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD
       txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & matem(i, 2)
       txtfiltro2 = txtfiltro2 & " AND TOPERACION = " & matem(i, 3)
       txtfiltro2 = txtfiltro2 & " AND C_EMISION = '" & matem(i, 1) & "' AND TIPOPOS = 1"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg1 = rmesa.Fields(0)
       rmesa.Close
       If noreg1 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          txtport = "EM " & matem(i, 1) & " POS " & matem(i, 2) & " T_OP " & matem(i, 3)
          For j = 1 To noreg1
              tipopos = rmesa.Fields("TIPOPOS")
              fechareg = rmesa.Fields("FECHAREG")
              txtnompos = rmesa.Fields("NOMPOS")
              horareg = rmesa.Fields("HORAREG")
              cposicion = rmesa.Fields("CPOSICION")
              coperacion = rmesa.Fields("COPERACION")
              txtfechar = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
              txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
              txtcadena = txtcadena & txtfecha & ","
              txtcadena = txtcadena & "'" & txtport & "',"
              txtcadena = txtcadena & tipopos & ","
              txtcadena = txtcadena & txtfechar & ","
              txtcadena = txtcadena & "'" & txtnompos & "',"
              txtcadena = txtcadena & "'" & horareg & "',"
              txtcadena = txtcadena & cposicion & ","
              txtcadena = txtcadena & "'" & coperacion & "')"
              ConAdo.Execute txtcadena
              rmesa.MoveNext
          Next j
          rmesa.Close
       End If
    Next i
End If
End Sub

Sub GenPortDerivPIDV(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfechar As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim noreg1 As Integer
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim txtport As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = " SELECT C_EMISION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND TIPOPOS = 1"
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & ClavePosPIDV & " GROUP BY C_EMISION ORDER BY C_EMISION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matem(1 To noreg) As String
   For i = 1 To noreg
       matem(i) = rmesa.Fields("C_EMISION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosSwaps
       txtfiltro2 = txtfiltro2 & " WHERE (CPOSICION,COPERACION,FECHAREG) IN "
       txtfiltro2 = txtfiltro2 & " (SELECT CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG FROM " & TablaPosSwaps
       txtfiltro2 = txtfiltro2 & " WHERE FECHAREG <= " & txtfecha & " AND C_EM_PIDV = '" & matem(i) & "' AND TIPOPOS = 1"
       txtfiltro2 = txtfiltro2 & " GROUP BY CPOSICION,COPERACION) AND TIPOPOS = 1"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg1 = rmesa.Fields(0)
       rmesa.Close
       If noreg1 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          txtport = "DERIVADOS PIDV " & matem(i)
          For j = 1 To noreg1
              tipopos = rmesa.Fields("TIPOPOS")
              fechareg = rmesa.Fields("FECHAREG")
              txtnompos = rmesa.Fields("NOMPOS")
              horareg = rmesa.Fields("HORAREG")
              cposicion = rmesa.Fields("CPOSICION")
              coperacion = rmesa.Fields("COPERACION")
              txtfechar = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
              txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
              txtcadena = txtcadena & txtfecha & ","
              txtcadena = txtcadena & "'" & txtport & "',"
              txtcadena = txtcadena & tipopos & ","
              txtcadena = txtcadena & txtfechar & ","
              txtcadena = txtcadena & "'" & txtnompos & "',"
              txtcadena = txtcadena & "'" & horareg & "',"
              txtcadena = txtcadena & cposicion & ","
              txtcadena = txtcadena & "'" & coperacion & "')"
              ConAdo.Execute txtcadena
              rmesa.MoveNext
          Next j
          rmesa.Close
       End If
   Next i
End If
End Sub


Sub GenPortPIDVDeriv(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfechar As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim noreg1 As Integer
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim txtport As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = " SELECT C_EMISION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND TIPOPOS = 1"
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & ClavePosPIDV & " GROUP BY C_EMISION ORDER BY C_EMISION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matem(1 To noreg) As String
   For i = 1 To noreg
       matem(i) = rmesa.Fields("C_EMISION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD
       txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPIDV
       txtfiltro2 = txtfiltro2 & " AND C_EMISION = '" & matem(i) & "' AND TIPOPOS = 1"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg1 = rmesa.Fields(0)
       rmesa.Close
       If noreg1 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          txtport = "PIDV " & matem(i)
          For j = 1 To noreg1
              tipopos = rmesa.Fields("TIPOPOS")
              fechareg = rmesa.Fields("FECHAREG")
              txtnompos = rmesa.Fields("NOMPOS")
              horareg = rmesa.Fields("HORAREG")
              cposicion = rmesa.Fields("CPOSICION")
              coperacion = rmesa.Fields("COPERACION")
              txtfechar = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
              txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
              txtcadena = txtcadena & txtfecha & ","
              txtcadena = txtcadena & "'" & txtport & "',"
              txtcadena = txtcadena & tipopos & ","
              txtcadena = txtcadena & txtfechar & ","
              txtcadena = txtcadena & "'" & txtnompos & "',"
              txtcadena = txtcadena & "'" & horareg & "',"
              txtcadena = txtcadena & cposicion & ","
              txtcadena = txtcadena & "'" & coperacion & "')"
              ConAdo.Execute txtcadena
              rmesa.MoveNext
          Next j
          rmesa.Close
       End If
       txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosMD
       txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfecha & " AND CPOSICION = " & ClavePosPIDV
       txtfiltro2 = txtfiltro2 & " AND C_EMISION = '" & matem(i) & "' AND TIPOPOS =1"
       txtfiltro2 = txtfiltro2 & " UNION "
       txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosSwaps
       txtfiltro2 = txtfiltro2 & " WHERE (CPOSICION,COPERACION,FECHAREG) IN "
       txtfiltro2 = txtfiltro2 & " (SELECT CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG FROM " & TablaPosSwaps
       txtfiltro2 = txtfiltro2 & " WHERE FECHAREG <= " & txtfecha & " AND C_EM_PIDV = '" & matem(i) & "' AND TIPOPOS = 1"
       txtfiltro2 = txtfiltro2 & " GROUP BY CPOSICION,COPERACION) AND TIPOPOS = 1"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg1 = rmesa.Fields(0)
       rmesa.Close
       If noreg1 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          txtport = "PIDV " & matem(i) & "+DERIV"
          For j = 1 To noreg1
              tipopos = rmesa.Fields("TIPOPOS")
              fechareg = rmesa.Fields("FECHAREG")
              txtnompos = rmesa.Fields("NOMPOS")
              horareg = rmesa.Fields("HORAREG")
              cposicion = rmesa.Fields("CPOSICION")
              coperacion = rmesa.Fields("COPERACION")
              txtfechar = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
              txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
              txtcadena = txtcadena & txtfecha & ","
              txtcadena = txtcadena & "'" & txtport & "',"
              txtcadena = txtcadena & tipopos & ","
              txtcadena = txtcadena & txtfechar & ","
              txtcadena = txtcadena & "'" & txtnompos & "',"
              txtcadena = txtcadena & "'" & horareg & "',"
              txtcadena = txtcadena & cposicion & ","
              txtcadena = txtcadena & "'" & coperacion & "')"
              ConAdo.Execute txtcadena
              rmesa.MoveNext
          Next j
          rmesa.Close
       End If
   Next i
End If
End Sub

Sub GenSubGenPort(ByVal fecha As Date, ByVal indice As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim nveces() As Integer
Dim txtesc() As String
Dim exito1 As Boolean
Dim contar As Long
Dim txtfecha As String
Dim txtborra As String
Dim txtfiltro As String
Dim txtcadena As String

exito = True
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
   contar = DeterminaMaxRegSubproc(id_tabla)
   contar = contar + 1
   txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Generación de portafolio", indice, 1, "", "", "", "", "", "", "", "", "", "", id_tabla)
   ConAdo.Execute txtcadena
txtmsg = "El proceso finalizo correctamente"
exito = True
End Sub


Sub DeterminaPosSwapsCobyPrim(ByVal fecha As Date, ByVal fechar As Date, ByVal txtport As String, ByVal tipopos As Integer)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim fechareg As Date
Dim noreg As Long
Dim i As Long
Dim cposicion As Integer
Dim coperacion As String
Dim txtnompos As String
Dim horareg As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosSwaps & " WHERE (NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG)"
txtfiltro2 = txtfiltro2 & " IN (SELECT NOMPOS,HORAREG,CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos
txtfiltro2 = txtfiltro2 & " GROUP BY NOMPOS,HORAREG,CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha & " AND TIPOPOS = " & tipopos & " AND INTENCION = 'C'"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosDeuda & " WHERE (NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG)"
txtfiltro2 = txtfiltro2 & " IN (SELECT NOMPOS,HORAREG,CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPosDeuda & " WHERE FECHAREG <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos
txtfiltro2 = txtfiltro2 & " GROUP BY NOMPOS,HORAREG,CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha & " AND TIPOPOS = " & tipopos
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
ConAdo.Execute txtborra
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       fechareg = rmesa.Fields("FECHAREG")                     'fecha de registro
       cposicion = rmesa.Fields("CPOSICION")                   'clave de posicion
       coperacion = rmesa.Fields("COPERACION")                 'clave de operacion
       txtnompos = rmesa.Fields("NOMPOS")                      'NOMBRE DE LA POSICION
       horareg = rmesa.Fields("HORAREG")                       'HORA DE REGISTRO
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","                        'la fecha del portafolio
       txtcadena = txtcadena & "'" & txtport & "',"                  'el nombre del portafolio
       txtcadena = txtcadena & "'" & tipopos & "',"                  'tipo de posicion
       txtcadena = txtcadena & txtfecha1 & ","                       'la fecha de registro
       txtcadena = txtcadena & "'" & txtnompos & "',"                'NOMBRE DE LA POSICION
       txtcadena = txtcadena & "'" & horareg & "',"                  'la hora de registro
       txtcadena = txtcadena & cposicion & ","                       'la clave de posicion
       txtcadena = txtcadena & "'" & coperacion & "')"               'la clave de OPERACION
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If


End Sub

Sub DeterminaPort(ByVal fecha As Date, ByVal fechar As Date, ByVal indice As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
On Error GoTo hayerror
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfechareg As String
Dim i As Long
Dim noreg As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim txtborra As String
Dim tipopos As Integer
Dim txtnompos As String
Dim fechareg As Date
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechareg = "to_date('" & Format(fechar, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = TraducirCadenaSQL(MatSQLPort(indice, 3), txtfechareg, 1)
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
   txtborra = txtborra & " PORTAFOLIO = '" & MatSQLPort(indice, 2) & "'"
   txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
   ConAdo.Execute txtborra
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 6) As Variant
   rmesa.MoveFirst
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")                                'TIPO DE POSICION
       fechareg = rmesa.Fields("FECHAREG")                               'FECHA DE registro
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")                              'clave de la posicion
       coperacion = rmesa.Fields("COPERACION")                             'clave de operacion
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","                              'la fecha del portafolio
       txtcadena = txtcadena & "'" & MatSQLPort(indice, 2) & "',"          'el nombre del portafolio
       txtcadena = txtcadena & tipopos & ","                               'tipo de posicion
       txtcadena = txtcadena & txtfecha1 & ","                             'la fecha de registro
       txtcadena = txtcadena & "'" & txtnompos & "',"                      'nombre de la posicion
       txtcadena = txtcadena & "'" & horareg & "',"                        'hora de registro
       txtcadena = txtcadena & cposicion & ","                             'clave de posicion
       txtcadena = txtcadena & "'" & coperacion & "')"                     'la clave de OPERACION
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
exito = True
Exit Sub
hayerror:
exito = False
txtmsg = "La cadena sql  " & MatSQLPort(indice, 2) & " esta mal construida " & error(Err())
If Err() = -2147467259 Then
   txtmsg = "La cadena sql  " & MatSQLPort(indice, 2) & " esta mal construida " & error(Err())
   exito = False
End If
End Sub

Sub DeterminaSubportEstructural(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfecha1 As String
Dim noport As Integer
Dim txtport1 As String
Dim txtport2 As String
Dim txtport3 As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim i As Integer
Dim j As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

noport = UBound(MatPortEstruct, 1)
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
For i = 1 To noport
    txtport1 = MatPortEstruct(i)
    txtport2 = MatPortEstruct(i) & " Deriv"
    txtport3 = MatPortEstruct(i) & " Oper"
    txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosSwaps & " WHERE"
    txtfiltro2 = txtfiltro2 & " (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN(SELECT TIPOPOS,MAX(FECHAREG)"
    txtfiltro2 = txtfiltro2 & " AS FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosSwaps & " WHERE"
    txtfiltro2 = txtfiltro2 & " (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & PrefijoBD & TablaPortPosEstructural
    txtfiltro2 = txtfiltro2 & " WHERE PORTAFOLIO = '" & txtport1 & "') AND FECHAREG <= " & txtfecha & " GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION)"
    txtfiltro2 = txtfiltro2 & " UNION "
    txtfiltro2 = txtfiltro2 & " SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDeuda & " WHERE"
    txtfiltro2 = txtfiltro2 & " (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN(SELECT TIPOPOS,MAX(FECHAREG)"
    txtfiltro2 = txtfiltro2 & " AS FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDeuda & " WHERE"
    txtfiltro2 = txtfiltro2 & " (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & PrefijoBD & TablaPortPosEstructural
    txtfiltro2 = txtfiltro2 & " WHERE PORTAFOLIO = '" & txtport1 & "') AND FECHAREG <= " & txtfecha & " GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION)"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg
           tipopos = rmesa.Fields("TIPOPOS")
           fechareg = rmesa.Fields("FECHAREG")
           txtnompos = rmesa.Fields("NOMPOS")
           horareg = rmesa.Fields("HORAREG")
           cposicion = rmesa.Fields("CPOSICION")
           coperacion = rmesa.Fields("COPERACION")
           txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtcadena = "INSERT INTO " & TablaPortPosicion & " VALUES("
           txtcadena = txtcadena & txtfecha & ","
           txtcadena = txtcadena & "'" & txtport1 & "',"
           txtcadena = txtcadena & tipopos & ","
           txtcadena = txtcadena & txtfecha1 & ","
           txtcadena = txtcadena & "'" & txtnompos & "',"
           txtcadena = txtcadena & "'" & horareg & "',"
           txtcadena = txtcadena & cposicion & ","
           txtcadena = txtcadena & "'" & coperacion & "')"
           ConAdo.Execute txtcadena
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
    txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosSwaps & " WHERE"
    txtfiltro2 = txtfiltro2 & " (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN(SELECT TIPOPOS,MAX(FECHAREG)"
    txtfiltro2 = txtfiltro2 & " AS FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosSwaps & " WHERE"
    txtfiltro2 = txtfiltro2 & " (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & PrefijoBD & TablaPortPosEstructural
    txtfiltro2 = txtfiltro2 & " WHERE PORTAFOLIO = '" & txtport1 & "') AND FECHAREG <= " & txtfecha & " GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION)"
    'txtfiltro2 = txtfiltro2 & " UNION "
    'txtfiltro2 = txtfiltro2 & " SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDeuda & " WHERE"
    'txtfiltro2 = txtfiltro2 & " (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN(SELECT TIPOPOS,MAX(FECHAREG)"
    'txtfiltro2 = txtfiltro2 & " AS FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDeuda & " WHERE"
    'txtfiltro2 = txtfiltro2 & " (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & PrefijoBD & TablaPortPosEstructural
    'txtfiltro2 = txtfiltro2 & " WHERE PORTAFOLIO = '" & txtport2 & "') AND FECHAREG <= " & txtfecha & " GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION)"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg
           tipopos = rmesa.Fields("TIPOPOS")
           fechareg = rmesa.Fields("FECHAREG")
           txtnompos = rmesa.Fields("NOMPOS")
           horareg = rmesa.Fields("HORAREG")
           cposicion = rmesa.Fields("CPOSICION")
           coperacion = rmesa.Fields("COPERACION")
           txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtcadena = "INSERT INTO " & TablaPortPosicion & " VALUES("
           txtcadena = txtcadena & txtfecha & ","
           txtcadena = txtcadena & "'" & txtport2 & "',"
           txtcadena = txtcadena & tipopos & ","
           txtcadena = txtcadena & txtfecha1 & ","
           txtcadena = txtcadena & "'" & txtnompos & "',"
           txtcadena = txtcadena & "'" & horareg & "',"
           txtcadena = txtcadena & cposicion & ","
           txtcadena = txtcadena & "'" & coperacion & "')"
           ConAdo.Execute txtcadena
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If

    'txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosSwaps & " WHERE"
    'txtfiltro2 = txtfiltro2 & " (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN(SELECT TIPOPOS,MAX(FECHAREG)"
    'txtfiltro2 = txtfiltro2 & " AS FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosSwaps & " WHERE"
    'txtfiltro2 = txtfiltro2 & " (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " &PrefijoBD & TablaPortPosEstructural
    'txtfiltro2 = txtfiltro2 & " WHERE PORTAFOLIO = '" & txtport3 & "') AND FECHAREG <= " & txtfecha & " GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION)"
    'txtfiltro2 = txtfiltro2 & " UNION "
    txtfiltro2 = " SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDeuda & " WHERE"
    txtfiltro2 = txtfiltro2 & " (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN(SELECT TIPOPOS,MAX(FECHAREG)"
    txtfiltro2 = txtfiltro2 & " AS FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDeuda & " WHERE"
    txtfiltro2 = txtfiltro2 & " (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & PrefijoBD & TablaPortPosEstructural
    txtfiltro2 = txtfiltro2 & " WHERE PORTAFOLIO = '" & txtport1 & "') AND FECHAREG <= " & txtfecha & " GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION)"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg
           tipopos = rmesa.Fields("TIPOPOS")
           fechareg = rmesa.Fields("FECHAREG")
           txtnompos = rmesa.Fields("NOMPOS")
           horareg = rmesa.Fields("HORAREG")
           cposicion = rmesa.Fields("CPOSICION")
           coperacion = rmesa.Fields("COPERACION")
           txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
           txtcadena = "INSERT INTO " & TablaPortPosicion & " VALUES("
           txtcadena = txtcadena & txtfecha & ","
           txtcadena = txtcadena & "'" & txtport3 & "',"
           txtcadena = txtcadena & tipopos & ","
           txtcadena = txtcadena & txtfecha1 & ","
           txtcadena = txtcadena & "'" & txtnompos & "',"
           txtcadena = txtcadena & "'" & horareg & "',"
           txtcadena = txtcadena & cposicion & ","
           txtcadena = txtcadena & "'" & coperacion & "')"
           ConAdo.Execute txtcadena
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
Next i

End Sub

Sub DeterminaSubPortArray(ByVal fecha As Date, ByVal txtport As String, ByRef mata() As Variant, ByVal tipopos As Integer)
Dim noreg As Integer
Dim i As Integer
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtcadena As String
Dim txtborra As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
ConAdo.Execute txtborra
   noreg = UBound(mata, 1)
   For i = 1 To noreg
       txtfecha1 = "to_date('" & Format(mata(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","                     'fecha de la posicion
       txtcadena = txtcadena & "'" & txtport & "',"               'nombre del portafolio
       txtcadena = txtcadena & "'" & tipopos & "',"               'tipo de posicion
       txtcadena = txtcadena & txtfecha1 & ","                    'fecha de registro
       txtcadena = txtcadena & "'Real',"                          'nombre de la posicion
       txtcadena = txtcadena & "'000000',"                        'hora de registro
       txtcadena = txtcadena & mata(i, 1) & ","                   'clave de la posicion
       txtcadena = txtcadena & "'" & mata(i, 3) & "')"            'clave de operacion
       ConAdo.Execute txtcadena
   Next i
End Sub

Sub DeterminaPortDer2(ByVal fecha As Date, ByVal fechar As Date, ByVal txtport As String, ByVal intencion As String, ByVal estruc As String, ByVal reclasif As String)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfechareg As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim txtborra As String
Dim noreg As Long
Dim i As Long
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

tipopos = 1
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechareg = "to_date('" & Format(fechar, "dd/mm/yyyy") & "','dd/mm/yyyy')"
If intencion = "N" And estruc = "N" And reclasif = "N" Then
   txtfiltro2 = "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosSwaps
   txtfiltro2 = txtfiltro2 & " WHERE (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG)"
   txtfiltro2 = txtfiltro2 & " IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
   txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha & " AND INTENCION = 'N' AND ESTRUCTURAL = 'N'"
   txtfiltro2 = txtfiltro2 & " UNION "
   txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosFwd
   txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfechareg
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND INTENCION = 'N' AND ESTRUCTURAL = 'N' AND RECLASIFICA = 'N' AND TIPOPOS ='" & tipopos & "'"
ElseIf intencion = "N" And estruc = "S" And reclasif = "N" Then
   txtfiltro2 = "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosSwaps
   txtfiltro2 = txtfiltro2 & " WHERE (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG)"
   txtfiltro2 = txtfiltro2 & " IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
   txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha & " AND INTENCION = 'N' AND ESTRUCTURAL = 'S'"
   txtfiltro2 = txtfiltro2 & " UNION "
   txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosFwd
   txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfechareg
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND INTENCION = 'N' AND ESTRUCTURAL = 'S' AND RECLASIFICA = 'N' AND TIPOPOS ='" & tipopos & "'"
ElseIf intencion = "C" And estruc = "N" And reclasif = "N" Then
   txtfiltro2 = "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosSwaps
   txtfiltro2 = txtfiltro2 & " WHERE (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG)"
   txtfiltro2 = txtfiltro2 & " IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
   txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha & " AND INTENCION = 'C'"
   txtfiltro2 = txtfiltro2 & " UNION "
   txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosFwd
   txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfechareg
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND INTENCION = 'C' AND ESTRUCTURAL = 'N' AND RECLASIFICA = 'N' AND TIPOPOS ='" & tipopos & "'"
ElseIf intencion = "*" And estruc = "*" And reclasif = "*" Then
   txtfiltro2 = "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosSwaps
   txtfiltro2 = txtfiltro2 & " WHERE (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG)"
   txtfiltro2 = txtfiltro2 & " IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
   txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " UNION "
   txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosFwd
   txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfechareg
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS ='" & tipopos & "'"
End If
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
   txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
   txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
   ConAdo.Execute txtborra
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")                               'tipo de posicion
       fechareg = rmesa.Fields("FECHAREG")                             'fecha de registro
       txtnompos = rmesa.Fields("NOMPOS")                              'nombre de la posicion
       horareg = rmesa.Fields("HORAREG")                               'HORA DE registro
       cposicion = rmesa.Fields("CPOSICION")                           'clave de posicion
       coperacion = rmesa.Fields("COPERACION")                         'clave de operacion
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","                           'la fecha del portafolio
       txtcadena = txtcadena & "'" & txtport & "',"                     'el nombre del portafolio
       txtcadena = txtcadena & "'" & tipopos & "',"                     'tipo de posicion
       txtcadena = txtcadena & txtfecha1 & ","                          'la fecha de registro
       txtcadena = txtcadena & "'" & txtnompos & "',"                   'nombre de la posicion
       txtcadena = txtcadena & "'" & horareg & "',"                     'hora de registro
       txtcadena = txtcadena & cposicion & ","                          'clave de posicion
       txtcadena = txtcadena & "'" & coperacion & "')"                  'la clave de OPERACION
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If

End Sub

Sub DeterminaPosDerivContrap(ByVal fecha As Date, ByVal tipopos As Integer, ByVal id_subproc As Integer, ByVal opcion As Integer)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim noreg As Long
Dim i As Long
Dim mata() As Long
Dim txttabla As String
Dim rmesa As New ADODB.recordset

txttabla = DetermTablaSubproc(opcion)
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT ID_CONTRAP FROM " & TablaPosFwd & " WHERE"
txtfiltro2 = txtfiltro2 & " FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & ClavePosDeriv
txtfiltro2 = txtfiltro2 & " GROUP BY ID_CONTRAP"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT ID_CONTRAP FROM " & TablaPosSwaps
txtfiltro2 = txtfiltro2 & " WHERE (CPOSICION,COPERACION,FECHAREG)"
txtfiltro2 = txtfiltro2 & " IN (SELECT CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos & " AND CPOSICION = " & ClavePosDeriv & " GROUP BY CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha & " AND TIPOPOS = " & tipopos & " GROUP BY ID_CONTRAP"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtborra = "DELETE FROM " & txttabla & " WHERE FECHAP = " & txtfecha
   txtborra = txtborra & " AND ID_SUBPROCESO = " & id_subproc
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg) As Long
   For i = 1 To noreg
       mata(i) = rmesa.Fields("ID_CONTRAP")
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       Call GenSubprocPortDerivContrap(fecha, mata(i), "Deriv Contrap " & mata(i), tipopos, id_subproc, opcion)
   Next i
End If
End Sub

  Sub GenSubprocPortDerivContrap(ByVal dtfecha As Date, ByVal id_contrap As Long, ByVal txtport As String, ByVal tipopos As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer)
  Dim txtfecha1 As String
  Dim txtcadena As String
  Dim txtfiltro As String
  Dim contar As Long
    txtfecha1 = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    contar = contar + 1
    txtcadena = CrearCadInsSub(dtfecha, id_subproc, contar, "Generacion port deriv contrap", id_contrap, txtport, tipopos, "", "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
  End Sub
  
  Sub GenSubprocPortDerivContrapyTOP(ByVal dtfecha As Date, ByVal id_contrap As Long, ByVal tipopos As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer)
  Dim txtfecha1 As String
  Dim txtcadena As String
  Dim txtfiltro As String
  Dim contar As Long
    txtfecha1 = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    contar = contar + 1
    txtcadena = CrearCadInsSub(dtfecha, id_subproc, contar, "Gen. port swaps x contrap y t op", id_contrap, tipopos, "", "", "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
  End Sub

Sub GenPortDerivContrap(ByVal fecha As Date, ByVal idcontrap As Long, ByVal txtport As String, ByVal tipopos As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtinserta As String
Dim txtborra As String
Dim fechareg As Date
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT CPOSICION,COPERACION,FECHAREG FROM " & TablaPosSwaps & " WHERE (CPOSICION,COPERACION,FECHAREG)"
txtfiltro2 = txtfiltro2 & " IN (SELECT CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos & " GROUP BY CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
txtfiltro2 = txtfiltro2 & " AND ID_CONTRAP = " & idcontrap
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT CPOSICION,COPERACION,FECHAREG FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos
txtfiltro2 = txtfiltro2 & " AND ID_CONTRAP = " & idcontrap
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
ConAdo.Execute txtborra

If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 3) As Variant
   For i = 1 To noreg
       cposicion = Val(rmesa.Fields("CPOSICION"))    'clave de la posicion
       fechareg = rmesa.Fields("FECHAREG")          'fecha de registro
       coperacion = rmesa.Fields("COPERACION")        'clave de operacion
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtinserta = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtinserta = txtinserta & txtfecha & ","                          'fecha del portafolio
       txtinserta = txtinserta & "'" & txtport & "',"                    'nombre del portafolio
       txtinserta = txtinserta & "'" & tipopos & "',"                    'tipo de posicion
       txtinserta = txtinserta & txtfecha1 & ","                         'fecha de registro
       txtinserta = txtinserta & "'Real',"                               'nombre de la posicion
       txtinserta = txtinserta & "'000000',"                             'hora de registro
       txtinserta = txtinserta & cposicion & ","                         'clave de la posicion
       txtinserta = txtinserta & "'" & coperacion & "')"                 'clave de operacion
       ConAdo.Execute txtinserta
       rmesa.MoveNext
   Next i
   rmesa.Close
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "No hay datos para este portafolio"
   exito = False
End If
End Sub

Sub DeterminaContrapPosFwd(ByVal fecha As Date, ByVal txtport As String, ByVal tipopos As Integer)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT ID_CONTRAP FROM " & TablaPosFwd & " WHERE"
txtfiltro2 = txtfiltro2 & " FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos
txtfiltro2 = txtfiltro2 & " GROUP BY ID_CONTRAP ORDER BY ID_CONTRAP"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg) As Long
   For i = 1 To noreg
       mata(i) = rmesa.Fields(0)
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       Call DeterPosFwdContrap(fecha, mata(i), txtport & " " & mata(i), tipopos)
   Next i
End If
End Sub

Sub GeneraPosSwapsxContrap(ByVal fecha As Date, ByVal tipopos As Integer, ByVal id_subproc As Integer, ByVal opcion As Integer)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT ID_CONTRAP FROM " & TablaPosSwaps & " WHERE (CPOSICION,COPERACION,FECHAREG)"
txtfiltro2 = txtfiltro2 & " IN (SELECT CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND CPOSICION = " & ClavePosDeriv & " GROUP BY CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha & " AND TIPOPOS ='" & tipopos & "' GROUP BY ID_CONTRAP ORDER BY ID_CONTRAP"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg) As Long
   For i = 1 To noreg
       mata(i) = rmesa.Fields("ID_CONTRAP")
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       Call GenSubprocPortDerivContrapyTOP(fecha, mata(i), tipopos, id_subproc, opcion)
   Next i
End If
End Sub

Sub DeterPosFwdContrap(ByVal fecha As Date, ByVal idcontrap As Long, ByVal txtport As String, ByVal tipopos As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtinserta As String
Dim txtborra As String
Dim cposicion As Integer
Dim coperacion As String
Dim horareg As String
Dim fechareg As Date
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPosFwd & " WHERE"
txtfiltro2 = txtfiltro2 & " TIPOPOS = " & tipopos
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & ClavePosDeriv
txtfiltro2 = txtfiltro2 & " AND FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND ID_CONTRAP = " & idcontrap
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
ConAdo.Execute txtborra

If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 3) As Variant
   For i = 1 To noreg
       cposicion = Val(rmesa.Fields("CPOSICION"))             'clave de la posicion
       fechareg = rmesa.Fields("FECHAREG")                    'fecha de registro
       coperacion = rmesa.Fields("COPERACION")                'clave de operacion
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtinserta = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtinserta = txtinserta & txtfecha & ","                                'fecha del portafolio
       txtinserta = txtinserta & "'" & txtport & "',"                          'nombre del portafolio
       txtinserta = txtinserta & "'" & tipopos & "',"                          'tipo de posicion
       txtinserta = txtinserta & txtfecha1 & ","                               'fecha de registro
       txtinserta = txtinserta & "'Real',"                                     'nombre de la posicion
       txtinserta = txtinserta & "'000000',"                                   'hora de registro
       txtinserta = txtinserta & cposicion & ","                               'clave de la posicion
       txtinserta = txtinserta & "'" & coperacion & "')"                       'clave de operacion
       ConAdo.Execute txtinserta
       rmesa.MoveNext
   Next i
   rmesa.Close
End If


End Sub

Sub DeterPosSwapsxContrap(ByVal fecha As Date, ByVal idcontrap As Long, ByVal txtport As String, ByVal tipopos As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtinserta As String
Dim txtborra As String
Dim cposicion As Integer
Dim coperacion As String
Dim fechareg As Date
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & " WHERE (CPOSICION,COPERACION,FECHAREG)"
txtfiltro2 = txtfiltro2 & " IN (SELECT CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos & " GROUP BY CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
txtfiltro2 = txtfiltro2 & " AND ID_CONTRAP = " & idcontrap
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
ConAdo.Execute txtborra
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 3) As Variant
   For i = 1 To noreg
       cposicion = Val(rmesa.Fields("CPOSICION"))    'clave de la posicion
       fechareg = rmesa.Fields("FECHAREG")           'fecha de registro
       coperacion = rmesa.Fields("COPERACION")       'clave de operacion
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtinserta = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtinserta = txtinserta & txtfecha & ","                          'fecha de la posicion
       txtinserta = txtinserta & "'" & txtport & "',"                    'nombre del portafolio
       txtinserta = txtinserta & "'" & tipopos & "',"                    'tipo de posicion
       txtinserta = txtinserta & txtfecha1 & ","                         'fecha de registro
       txtinserta = txtinserta & "'Real',"                               'nombre de la posicion
       txtinserta = txtinserta & "'000000',"                             'hora de registro
       txtinserta = txtinserta & cposicion & ","                         'clave de la posicion
       txtinserta = txtinserta & "'" & coperacion & "')"                 'clave de operacion
       ConAdo.Execute txtinserta
       rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
   Next i
End If


End Sub

Sub DeterPosSwapsContrapTOper(ByVal fecha As Date, ByVal idcontrap As Long, ByVal txtprod As String, ByVal txtport As String, ByVal tipopos As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtinserta As String
Dim txtborra As String
Dim cposicion As Integer
Dim coperacion As String
Dim fechareg As Date
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & " WHERE (CPOSICION,COPERACION,FECHAREG)"
txtfiltro2 = txtfiltro2 & " IN (SELECT CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos & " GROUP BY CPOSICION,COPERACION) AND FINICIO <= " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos
txtfiltro2 = txtfiltro2 & " AND ID_CONTRAP = " & idcontrap
txtfiltro2 = txtfiltro2 & " AND FVALUACION LIKE '" & txtprod & "%'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
ConAdo.Execute txtborra
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 3) As Variant
   For i = 1 To noreg
       cposicion = Val(rmesa.Fields("CPOSICION"))    'clave de la posicion
       fechareg = rmesa.Fields("FECHAREG")           'fecha de registro
       coperacion = rmesa.Fields("COPERACION")       'clave de operacion
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtinserta = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtinserta = txtinserta & txtfecha & ","                          'fecha de la posicion
       txtinserta = txtinserta & "'" & txtport & "',"                    'nombre del portafolio
       txtinserta = txtinserta & "'" & tipopos & "',"                    'tipo de posicion
       txtinserta = txtinserta & txtfecha1 & ","                         'fecha de registro
       txtinserta = txtinserta & "'Real',"                               'nombre de la posicion
       txtinserta = txtinserta & "'000000',"                             'hora de registro
       txtinserta = txtinserta & cposicion & ","                         'clave de la posicion
       txtinserta = txtinserta & "'" & coperacion & "')"                 'clave de operacion
       ConAdo.Execute txtinserta
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
End Sub

Sub DeterminaPortCons(ByVal fecha As Date, ByVal fechar As Date, ByVal txtport As String)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfechareg As String
Dim noreg As Integer
Dim i As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim txtborra As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechareg = "to_date('" & Format(fechar, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosMD
txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfechareg
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & 1
txtfiltro2 = txtfiltro2 & " AND (CPOSICION = " & ClavePosMD
txtfiltro2 = txtfiltro2 & " OR CPOSICION = " & ClavePosTeso & ")"
txtfiltro2 = txtfiltro2 & " AND INTENCION = 'N'"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosDiv & " WHERE "
txtfiltro2 = txtfiltro2 & " TIPOPOS = " & 1
txtfiltro2 = txtfiltro2 & " AND FECHAREG = " & txtfechareg
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & ClavePosMC
txtfiltro2 = txtfiltro2 & " AND INTENCION = 'N'"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE (TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG)"
txtfiltro2 = txtfiltro2 & " IN (SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,MAX(FECHAREG) AS FECHAREG"
txtfiltro2 = txtfiltro2 & " FROM " & TablaPosSwaps & " WHERE FECHAREG <= " & txtfechareg
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & 1 & " GROUP BY TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION) AND FINICIO <= " & txtfechareg
txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfechareg & " AND INTENCION = 'N'"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,NOMPOS,HORAREG,CPOSICION,COPERACION,FECHAREG FROM " & TablaPosFwd
txtfiltro2 = txtfiltro2 & " WHERE TIPOPOS = " & 1
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & ClavePosDeriv
txtfiltro2 = txtfiltro2 & " AND FECHAREG = " & txtfechareg
txtfiltro2 = txtfiltro2 & " AND INTENCION = 'N'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
   txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
   txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
   ConAdo.Execute txtborra
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")                               'tipo de posicion
       fechareg = rmesa.Fields("FECHAREG")                             'fecha de registro
       txtnompos = rmesa.Fields("NOMPOS")                              'nombre de la posicion
       horareg = rmesa.Fields("HORAREG")                               'HORA DE registro
       cposicion = rmesa.Fields("CPOSICION")                           'clave de posicion
       coperacion = rmesa.Fields("COPERACION")                         'clave de operacion
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","                           'la fecha del portafolio
       txtcadena = txtcadena & "'" & txtport & "',"                     'el nombre del portafolio
       txtcadena = txtcadena & "'" & tipopos & "',"                     'tipo de posicion
       txtcadena = txtcadena & txtfecha1 & ","                          'la fecha de registro
       txtcadena = txtcadena & "'" & txtnompos & "',"                   'nombre de la posicion
       txtcadena = txtcadena & "'" & horareg & "',"                     'hora de registro
       txtcadena = txtcadena & cposicion & ","                          'clave de posicion
       txtcadena = txtcadena & "'" & coperacion & "')"                  'la clave de OPERACION
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If

End Sub

Sub DeterminaPortMD(ByVal fecha As Date, ByVal fechar As Date, ByVal txtport As String, ByVal intencion As String)
Dim txtfecha As String
Dim txtfechareg As String
Dim txtfecha1 As String
Dim noreg As Integer
Dim i As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim txtborra As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

tipopos = 1

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechareg = "to_date('" & Format(fechar, "dd/mm/yyyy") & "','dd/mm/yyyy')"
If intencion = "N" Then
   txtfiltro2 = "SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfechareg & " AND TIPOPOS = " & tipopos
   txtfiltro2 = txtfiltro2 & " AND (CPOSICION =  " & ClavePosMD & " OR CPOSICION = " & ClavePosTeso & ")"
   txtfiltro2 = txtfiltro2 & " AND INTENCION ='N'"
ElseIf intencion = "*" Then
   txtfiltro2 = "SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfechareg & " AND TIPOPOS = " & tipopos
   txtfiltro2 = txtfiltro2 & " AND (CPOSICION =  " & ClavePosMD & " OR CPOSICION = " & ClavePosTeso & ")"
End If
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
   txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
   txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
   ConAdo.Execute txtborra
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")                               'tipo de posicion
       fechareg = rmesa.Fields("FECHAREG")                             'fecha de registro
       txtnompos = rmesa.Fields("NOMPOS")                              'nombre de la posicion
       horareg = rmesa.Fields("HORAREG")                               'HORA DE registro
       cposicion = rmesa.Fields("CPOSICION")                           'clave de posicion
       coperacion = rmesa.Fields("COPERACION")                         'clave de operacion
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","                           'la fecha del portafolio
       txtcadena = txtcadena & "'" & txtport & "',"                     'el nombre del portafolio
       txtcadena = txtcadena & "'" & tipopos & "',"                     'tipo de posicion
       txtcadena = txtcadena & txtfecha1 & ","                          'la fecha de registro
       txtcadena = txtcadena & "'" & txtnompos & "',"                   'nombre de la posicion
       txtcadena = txtcadena & "'" & horareg & "',"                     'hora de registro
       txtcadena = txtcadena & cposicion & ","                          'clave de posicion
       txtcadena = txtcadena & "'" & coperacion & "')"                  'la clave de OPERACION
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If

End Sub


Sub DeterminaPortMC(ByVal fecha As Date, ByVal fechar As Date, ByVal txtport As String, ByVal intencion As String)
Dim txtfecha As String
Dim txtfechareg As String
Dim txtfecha1 As String
Dim noreg As Integer
Dim i As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim txtborra As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

tipopos = 1
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechareg = "to_date('" & Format(fechar, "dd/mm/yyyy") & "','dd/mm/yyyy')"
If intencion = "N" Then
   txtfiltro2 = "SELECT * FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfechareg & " AND TIPOPOS = " & tipopos & " AND INTENCION = 'N'"
Else
   txtfiltro2 = "SELECT * FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfechareg & " AND TIPOPOS = " & tipopos
End If
txtfiltro2 = txtfiltro2 & " AND CPOSICION =  " & ClavePosMC
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
   txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
   txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
   ConAdo.Execute txtborra
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 2) As Variant
   rmesa.MoveFirst
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")                               'tipo de posicion
       fechareg = rmesa.Fields("FECHAREG")                             'fecha de registro
       txtnompos = rmesa.Fields("NOMPOS")                              'nombre de la posicion
       horareg = rmesa.Fields("HORAREG")                               'HORA DE registro
       cposicion = rmesa.Fields("CPOSICION")                           'clave de posicion
       coperacion = rmesa.Fields("COPERACION")                         'clave de operacion
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","                           'la fecha del portafolio
       txtcadena = txtcadena & "'" & txtport & "',"                     'el nombre del portafolio
       txtcadena = txtcadena & "'" & tipopos & "',"                     'tipo de posicion
       txtcadena = txtcadena & txtfecha1 & ","                          'la fecha de registro
       txtcadena = txtcadena & "'" & txtnompos & "',"                   'nombre de la posicion
       txtcadena = txtcadena & "'" & horareg & "',"                     'hora de registro
       txtcadena = txtcadena & cposicion & ","                          'clave de posicion
       txtcadena = txtcadena & "'" & coperacion & "')"                  'la clave de OPERACION
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If

End Sub

Sub DeterminaPortFwds2(ByVal fecha As Date, ByVal fechar As Date, ByVal txtport As String, ByVal intencion As String, ByVal estruc As String, ByVal reclasif As String)
Dim txport As String
Dim txtfecha As String
Dim txtfechareg As String
Dim txtfecha1 As String
Dim noreg As Integer
Dim i As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim txtborra As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

tipopos = 1

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechareg = "to_date('" & Format(fechar, "dd/mm/yyyy") & "','dd/mm/yyyy')"
If intencion = "N" And estruc = "S" And reclasif = "N" Then
   txtfiltro2 = "SELECT * FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfechareg
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos & " AND INTENCION = 'N' AND ESTRUCTURAL = 'S' AND RECLASIFICA = 'N'"
ElseIf intencion = "N" And estruc = "N" And reclasif = "N" Then
   txtfiltro2 = "SELECT * FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfechareg
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos & " AND INTENCION = 'N' AND ESTRUCTURAL = 'N' AND RECLASIFICA ='N'"
ElseIf intencion = "C" Then
   txtfiltro2 = "SELECT * FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfechareg
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos & " AND INTENCION = 'C' AND ESTRUCTURAL = 'N' AND RECLASIFICA ='N'"
ElseIf intencion = "N" And estruc = "N" And reclasif = "S" Then
   txtfiltro2 = "SELECT * FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfechareg
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos & " AND INTENCION = 'N' AND ESTRUCTURAL = 'N' AND RECLASIFICA ='S'"
ElseIf intencion = "*" And estruc = "*" And reclasif = "*" Then
   txtfiltro2 = "SELECT * FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfechareg
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS = " & tipopos
End If
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
txtborra = "DELETE FROM " & TablaPortPosicion & "  WHERE"
txtborra = txtborra & " PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND FECHA_PORT = " & txtfecha
ConAdo.Execute txtborra
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")                               'tipo de posicion
       fechareg = rmesa.Fields("FECHAREG")                             'fecha de registro
       txtnompos = rmesa.Fields("NOMPOS")                              'nombre de la posicion
       horareg = rmesa.Fields("HORAREG")                               'HORA DE registro
       cposicion = rmesa.Fields("CPOSICION")                           'clave de posicion
       coperacion = rmesa.Fields("COPERACION")                         'clave de operacion
       txtfecha1 = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPortPosicion & "  VALUES("
       txtcadena = txtcadena & txtfecha & ","                           'la fecha del portafolio
       txtcadena = txtcadena & "'" & txtport & "',"                     'el nombre del portafolio
       txtcadena = txtcadena & "'" & tipopos & "',"                     'tipo de posicion
       txtcadena = txtcadena & txtfecha1 & ","                          'la fecha de registro
       txtcadena = txtcadena & "'" & txtnompos & "',"                   'nombre de la posicion
       txtcadena = txtcadena & "'" & horareg & "',"                     'hora de registro
       txtcadena = txtcadena & cposicion & ","                          'clave de posicion
       txtcadena = txtcadena & "'" & coperacion & "')"                  'la clave de OPERACION
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If

End Sub


