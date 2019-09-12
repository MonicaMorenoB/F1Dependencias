Attribute VB_Name = "ModVaRExtremo"
Option Explicit

Sub Rutina2EscEstres(ByVal fecha As Date, _
                     ByVal txtport As String, _
                     ByRef matpos() As propPosRiesgo, _
                     ByRef matposmd() As propPosMD, _
                     ByRef matposdiv() As propPosDiv, _
                     ByRef matposswaps() As propPosSwaps, _
                     ByRef matposfwd() As propPosFwd, _
                     ByRef matflswap() As estFlujosDeuda, _
                     ByRef matposdeuda() As propPosDeuda, _
                     ByRef matfldeuda() As estFlujosDeuda, _
                     ByRef exito As Boolean)
Dim tasas1() As Double
Dim tasas2() As Double
  If EsArrayVacio(matEscEstres) Then
     Call CrearEscenariosTasas(fecha, exito)
  End If
  Call CalculoEscStress(fecha, txtport, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda)
End Sub

Sub GenSubProcValPosSubPort(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
   Dim txtborra As String
   Dim txtfecha As String
   Dim exito1 As Boolean
   Dim exito2 As Boolean
   Dim txttabla As String
   
   txttabla = DetermTablaSubproc(id_tabla)
   txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
   txtborra = "DELETE FROM " & txttabla & " WHERE ID_SUBPROCESO = " & id_subproc
   txtborra = txtborra & " AND FECHAP = " & txtfecha
   ConAdo.Execute txtborra
   Call GenSubProcValPosPort(fecha, txtport, txtportfr, txtgrupoport, 1, id_subproc, id_tabla, txtmsg, exito1)
   Call GenSubProcValPosPort(fecha, txtport, txtportfr, txtgrupoport, 2, id_subproc, id_tabla, txtmsg, exito2)
   Call GenProcValEmPos(fecha, txtport, txtportfr, 1, id_subproc, id_tabla)
   Call GenProcValEmPos(fecha, txtport, txtportfr, 2, id_subproc, id_tabla)
   Call GenProcValPIDV(fecha, txtport, txtportfr, 1, id_subproc, id_tabla)
   Call GenProcValPIDV(fecha, txtport, txtportfr, 2, id_subproc, id_tabla)
   Call GenProcValEstructural(fecha, txtport, txtportfr, 1, id_subproc, id_tabla)
   Call GenProcValEstructural(fecha, txtport, txtportfr, 2, id_subproc, id_tabla)
   If exito1 And exito2 Then
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      txtmsg = "Algo salio mal"
      exito = False
   End If

End Sub

Sub GenProcValEmPos(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal id_val As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim matem() As String
Dim txtcadena As String
Dim txttabla As String
Dim txtsubport As String
Dim rmesa As New ADODB.recordset

txttabla = DetermTablaSubproc(id_tabla)
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT C_EMISION,CPOSICION,TOPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1  GROUP BY C_EMISION,CPOSICION,TOPERACION ORDER BY CPOSICION,C_EMISION,TOPERACION"
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
   contar = DeterminaMaxRegSubproc(id_tabla)
   For i = 1 To UBound(matem, 1)
       txtsubport = "EM " & matem(i, 1) & " POS " & matem(i, 2) & " T_OP " & matem(i, 3)
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Valuación de subport", txtport, txtsubport, txtportfr, id_val, "", "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
   Next i
   
End If
End Sub

Sub GenProcValPIDV(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal id_val As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim matem() As String
Dim txtcadena As String
Dim txttabla As String
Dim rmesa As New ADODB.recordset

txttabla = DetermTablaSubproc(id_tabla)
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT C_EMISION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND CPOSICION = " & ClavePosPIDV & " GROUP BY C_EMISION ORDER BY C_EMISION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matem(1 To noreg, 1 To 3) As String
   For i = 1 To noreg
       matem(i, 1) = rmesa.Fields("C_EMISION")
       matem(i, 2) = "PIDV " & matem(i, 1)
       matem(i, 3) = "PIDV " & matem(i, 1) & "+DERIV"
       rmesa.MoveNext
   Next i
   rmesa.Close
   contar = DeterminaMaxRegSubproc(id_tabla)
   For i = 1 To UBound(matem, 1)
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Valuación de subport", txtport, matem(i, 2), txtportfr, id_val, "", "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Valuación de subport", txtport, matem(i, 3), txtportfr, id_val, "", "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
   Next i
   
End If
End Sub

Function DeterminaMaxRegSubproc(ByVal opcion As Integer)
Dim txtfiltro As String
Dim txttabla As String
Dim rmesa As New ADODB.recordset
txttabla = DetermTablaSubproc(opcion)
   txtfiltro = "select max(folio) from " & txttabla
   rmesa.Open txtfiltro, ConAdo
   If rmesa.RecordCount <> 0 Then
      If Not EsVariableVacia(rmesa.Fields(0)) Then
         DeterminaMaxRegSubproc = rmesa.Fields(0)
      Else
         DeterminaMaxRegSubproc = 0
      End If
   Else
      DeterminaMaxRegSubproc = 0
   End If
   rmesa.Close
End Function

Sub GenProcValEstructural(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal id_val As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer)
Dim txtfiltro As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim txtcadena As String
Dim txtport1 As String
Dim txtport2 As String
Dim txtport3 As String

   contar = DeterminaMaxRegSubproc(id_tabla)
   noreg = UBound(MatPortEstruct, 1)
   For i = 1 To noreg
       txtport1 = MatPortEstruct(i)
       txtport2 = MatPortEstruct(i) & " Deriv"
       txtport3 = MatPortEstruct(i) & " Oper"
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Valuación de subport", txtport, txtport1, txtportfr, id_val, "", "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Valuación de subport", txtport, txtport2, txtportfr, id_val, "", "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Valuación de subport", txtport, txtport3, txtportfr, id_val, "", "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
   Next i
   
End Sub

Sub GenProcPyGEstructural(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer)
Dim txtfiltro As String
Dim noreg As Long
Dim contar As Long
Dim txtport1 As String
Dim txtport2 As String
Dim txtport3 As String
Dim i As Long
Dim txtcadena As String
   contar = DeterminaMaxRegSubproc(id_tabla)
   noreg = UBound(MatPortEstruct, 1)
   For i = 1 To noreg
       txtport1 = MatPortEstruct(i)
       txtport2 = MatPortEstruct(i) & " Deriv"
       txtport3 = MatPortEstruct(i) & " Oper"
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Calculo de pyg de subport", txtport, txtportfr, txtport1, noesc, htiempo, "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Calculo de pyg de subport", txtport, txtportfr, txtport2, noesc, htiempo, "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Calculo de pyg de subport", txtport, txtportfr, txtport3, noesc, htiempo, "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
   Next i
   
End Sub

Sub GenProcPyGEmPos(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim matem() As String
Dim txtcadena As String
Dim txtsubport As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT C_EMISION,CPOSICION,TOPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 GROUP BY C_EMISION,CPOSICION,TOPERACION ORDER BY CPOSICION,C_EMISION,TOPERACION"
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
   contar = DeterminaMaxRegSubproc(id_tabla)
   For i = 1 To UBound(matem, 1)
       txtsubport = "EM " & matem(i, 1) & " POS " & matem(i, 2) & " T_OP " & matem(i, 3)
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Calculo de pyg de subport", txtport, txtportfr, txtsubport, noesc, htiempo, "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
   Next i
   
End If
End Sub


Sub GenProcPyGPIDV(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim matem() As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT C_EMISION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND CPOSICION = " & ClavePosPIDV & " GROUP BY C_EMISION ORDER BY C_EMISION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matem(1 To noreg, 1 To 3) As String
   For i = 1 To noreg
       matem(i, 1) = rmesa.Fields("C_EMISION")
       matem(i, 2) = "PIDV " & matem(i, 1)
       matem(i, 3) = "PIDV " & matem(i, 1) & "+DERIV"
       rmesa.MoveNext
   Next i
   rmesa.Close
   contar = DeterminaMaxRegSubproc(id_tabla)
   For i = 1 To UBound(matem, 1)
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Calculo de pyg de subport", txtport, txtportfr, matem(i, 2), noesc, htiempo, "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Calculo de pyg de subport", txtport, txtportfr, matem(i, 3), noesc, htiempo, "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
   Next i
   
End If
End Sub

Sub GenProcCVaREstructural(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByVal id_subproc As Integer)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim matem() As String
Dim txtcadena As String
Dim txttvar As String
Dim exito As Boolean
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim txtborra As String
Dim valor As Double
Dim txtmsg As String
Dim txtport1 As String
Dim txtport2 As String
Dim txtport3 As String


txttvar = "CVARH"
exito = False
noreg = UBound(MatPortEstruct, 1)
If noreg <> 0 Then
   txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
   For i = 1 To noreg
       txtport1 = MatPortEstruct(i)
       txtport2 = MatPortEstruct(i) & " Deriv"
       txtport3 = MatPortEstruct(i) & " Oper"
    
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & txtport1 & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & 1 - nconf
       ConAdo.Execute txtborra, noreg
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & txtport1 & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & nconf
       ConAdo.Execute txtborra, noreg
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & txtport2 & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & 1 - nconf
       ConAdo.Execute txtborra, noreg
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & txtport2 & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & nconf
       ConAdo.Execute txtborra, noreg
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & txtport3 & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & 1 - nconf
       ConAdo.Execute txtborra, noreg
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & txtport3 & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & nconf
       ConAdo.Execute txtborra, noreg
      
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, txtport1, noesc, htiempo, 1 - nconf, exito1)
       If exito1 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, txtport1, txtportfr, txttvar, noesc, htiempo, 0, 1 - nconf, 0, valor)
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, txtport1, noesc, htiempo, nconf, exito2)
       If exito2 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, txtport1, txtportfr, txttvar, noesc, htiempo, 0, nconf, 0, valor)
       
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, txtport2, noesc, htiempo, 1 - nconf, exito1)
       If exito1 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, txtport2, txtportfr, txttvar, noesc, htiempo, 0, 1 - nconf, 0, valor)
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, txtport2, noesc, htiempo, nconf, exito2)
       If exito2 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, txtport2, txtportfr, txttvar, noesc, htiempo, 0, nconf, 0, valor)
       
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, txtport3, noesc, htiempo, 1 - nconf, exito1)
       If exito1 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, txtport3, txtportfr, txttvar, noesc, htiempo, 0, 1 - nconf, 0, valor)
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, txtport3, noesc, htiempo, nconf, exito2)
       If exito2 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, txtport3, txtportfr, txttvar, noesc, htiempo, 0, nconf, 0, valor)

   Next i
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "El portafolio no esta definido"
   exito = False
End If

End Sub

Sub GenProcCVaREmPos(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByVal id_subproc As Integer)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim matem() As String
Dim txtcadena As String
Dim txttvar As String
Dim exito As Boolean
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim txtborra As String
Dim valor As Double
Dim txtmsg As String
Dim txtsubport As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT C_EMISION,CPOSICION,TOPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1  GROUP BY C_EMISION,CPOSICION,TOPERACION ORDER BY C_EMISION"
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
   txttvar = "CVARH"
   exito = False
   For i = 1 To UBound(matem, 1)
       txtsubport = "EM " & matem(i, 1) & " POS " & matem(i, 2) & " T_OP " & matem(i, 3)
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & txtsubport & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & 1 - nconf
       ConAdo.Execute txtborra, noreg
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & txtsubport & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & nconf
       ConAdo.Execute txtborra, noreg
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, txtsubport, noesc, htiempo, 1 - nconf, exito1)
       If exito1 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, txtsubport, txtportfr, txttvar, noesc, htiempo, 0, 1 - nconf, 0, valor)
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, txtsubport, noesc, htiempo, nconf, exito2)
       If exito2 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, txtsubport, txtportfr, txttvar, noesc, htiempo, 0, nconf, 0, valor)
   Next i
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "El portafolio no esta definido"
   exito = False
End If

  

End Sub

Sub GenProcCVaRPIDV(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nconf As Double, ByVal id_subproc As Integer)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim matem() As String
Dim txtcadena As String
Dim txttvar As String
Dim exito As Boolean
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim txtborra As String
Dim valor As Double
Dim txtmsg As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT C_EMISION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND CPOSICION = " & ClavePosPIDV & " GROUP BY C_EMISION ORDER BY C_EMISION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matem(1 To noreg, 1 To 3) As String
   For i = 1 To noreg
       matem(i, 1) = rmesa.Fields("C_EMISION")
       matem(i, 2) = "PIDV " & matem(i, 1)
       matem(i, 3) = "PIDV " & matem(i, 1) & "+DERIV"
       rmesa.MoveNext
   Next i
   rmesa.Close
   txttvar = "CVARH"
   exito = False
   For i = 1 To UBound(matem, 1)
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & matem(i, 2) & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & 1 - nconf
       ConAdo.Execute txtborra, noreg
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & matem(i, 2) & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & nconf
       ConAdo.Execute txtborra, noreg
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & matem(i, 3) & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & 1 - nconf
       ConAdo.Execute txtborra, noreg
       txtborra = "DELETE FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT  = '" & matem(i, 3) & "' AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo & " AND TVAR ='" & txttvar & "' AND NCONF = " & nconf
       ConAdo.Execute txtborra, noreg
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, matem(i, 2), noesc, htiempo, 1 - nconf, exito1)
       If exito1 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, matem(i, 2), txtportfr, txttvar, noesc, htiempo, 0, 1 - nconf, 0, valor)
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, matem(i, 2), noesc, htiempo, nconf, exito2)
       If exito2 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, matem(i, 2), txtportfr, txttvar, noesc, htiempo, 0, nconf, 0, valor)
       
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, matem(i, 3), noesc, htiempo, 1 - nconf, exito1)
       If exito1 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, matem(i, 3), txtportfr, txttvar, noesc, htiempo, 0, 1 - nconf, 0, valor)
       valor = CalcularCVaRPyG(fecha, fecha, fecha, txtport, txtportfr, matem(i, 3), noesc, htiempo, nconf, exito2)
       If exito2 Then Call InsertaRegVaR(fecha, fecha, fecha, txtport, matem(i, 3), txtportfr, txttvar, noesc, htiempo, 0, nconf, 0, valor)
   Next i
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "El portafolio no esta definido"
   exito = False
End If

  

End Sub


Sub GenSubCalcPyGOper(ByVal f_pos As Date, ByVal txtport As String, ByVal txtnompos As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)

Dim i As Integer
Dim noreg As Long
Dim exito1 As Boolean
Dim contar As Long
Dim txtfecha As String
Dim txtborra As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

exito = True

txtfecha = "TO_DATE('" & Format$(f_pos, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtborra = "DELETE FROM " & TablaPLHistOper & " WHERE F_POSICION = " & txtfecha
   txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
   txtborra = txtborra & " AND ESC_FACTORES = '" & txtportfr & "'"
   txtborra = txtborra & " AND NOESC = " & noesc
   txtborra = txtborra & " AND HTIEMPO = " & htiempo
   ConAdo.Execute txtborra
   txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha
   txtborra = txtborra & " AND ID_SUBPROCESO = " & id_subproc
   txtborra = txtborra & " AND PARAMETRO1 = '" & txtport & "'"
   txtborra = txtborra & " AND PARAMETRO2 = '" & txtportfr & "'"
   txtborra = txtborra & " AND PARAMETRO5 = '" & txtnompos & "'"
   txtborra = txtborra & " AND PARAMETRO9 = '" & noesc & "'"
   txtborra = txtborra & " AND PARAMETRO10 = '" & htiempo & "'"
   ConAdo.Execute txtborra
   contar = DeterminaMaxRegSubproc(id_tabla)
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       contar = contar + 1
       txtcadena = CrearCadInsSub(f_pos, id_subproc, contar, "Cálculo de PyG x op", txtport, txtportfr, tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, noesc, htiempo, "", "", id_tabla)
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close

txtmsg = "El proceso finalizo correctamente"
exito = True
End If
End Sub

Sub GenSubCalcPyGOperVR(ByVal fecha As Date, ByVal txtport As String, ByVal txtnompos As String, ByVal txtportfr As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal pfwd As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)

Dim i As Integer
Dim noreg As Long
Dim exito1 As Boolean
Dim contar As Long
Dim txtfecha As String
Dim txtborra As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

exito = True

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtborra = "DELETE FROM " & TablaPLHistOperVR & " WHERE FECHA = " & txtfecha
   txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
   txtborra = txtborra & " AND ESC_FACTORES = '" & txtportfr & "'"
   txtborra = txtborra & " AND NOESC = " & noesc
   txtborra = txtborra & " AND HTIEMPO = " & htiempo
   txtborra = txtborra & " AND P_FWD = " & pfwd
   ConAdo.Execute txtborra
   txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha
   txtborra = txtborra & " AND ID_SUBPROCESO = " & id_subproc
   txtborra = txtborra & " AND PARAMETRO1 = '" & txtport & "'"
   txtborra = txtborra & " AND PARAMETRO2 = '" & txtportfr & "'"
   txtborra = txtborra & " AND PARAMETRO5 = '" & txtnompos & "'"
   txtborra = txtborra & " AND PARAMETRO9 = '" & noesc & "'"
   txtborra = txtborra & " AND PARAMETRO10 = '" & htiempo & "'"
   txtborra = txtborra & " AND PARAMETRO11 = '" & pfwd & "'"
   ConAdo.Execute txtborra
   contar = DeterminaMaxRegSubproc(id_tabla)
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Cálculo de PyG x op VR", txtport, txtportfr, tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, noesc, htiempo, pfwd, "", id_tabla)
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close

txtmsg = "El proceso finalizo correctamente"
exito = True
End If
End Sub

Sub GenSubCalcSensibOper(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim i As Long
Dim noreg As Long
Dim contar As Long
Dim txtborra As String
Dim txtcadena As String
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim cposicion As Long
Dim fechareg As Date
Dim coperacion As String
Dim tipopos As Integer
Dim horareg As String
Dim txtnompos As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtborra = "DELETE FROM " & TablaSensibN & " WHERE FECHA = " & txtfecha
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
txtborra = txtborra & " AND PORT_FR = '" & txtportfr & "'"
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha
txtborra = txtborra & " AND ID_SUBPROCESO = " & id_subproc
txtborra = txtborra & " AND PARAMETRO1 = '" & txtport & "'"
ConAdo.Execute txtborra
txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   contar = DeterminaMaxRegSubproc(id_tabla)
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Cálculo de sensib de oper", txtport, txtportfr, tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   exito = False
End If
End Sub

Sub GenSubProcValPosPort(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal id_val As Integer, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim nveces() As Integer
Dim txtesc() As String
Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim exito1 As Boolean
Dim noesc As Integer
Dim contar As Long
Dim txtfecha As String
Dim txtborra As String
Dim txtfiltro As String
Dim txtcadena As String
Dim txttabla As String

txttabla = DetermTablaSubproc(id_tabla)
exito = True
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
If UBound(MatGruposPortPos, 1) <> 0 Then
   contar = DeterminaMaxRegSubproc(id_tabla)
   For i = 1 To UBound(MatGruposPortPos, 1)
    contar = contar + 1
    txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Valuación de subport", txtport, MatGruposPortPos(i, 3), txtportfr, id_val, "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
   Next i
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   txtmsg = "El portafolio no esta definido"
   exito = False
End If
End Sub

Sub GenSubProcEscenEstres(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim nveces() As Integer
Dim txtesc() As String
Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim noreg As Long
Dim exito1 As Boolean
Dim noesc As Integer
Dim contar As Long
Dim txtfecha As String
Dim txtborra As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim cposicion As Integer
Dim coperacion As String
Dim fechareg As Date
Dim tipopos As Integer
Dim txtnompos As String
Dim horareg As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtborra = "DELETE FROM " & TablaResEscEstres & " WHERE FECHA = " & txtfecha
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha
txtborra = txtborra & " AND ID_SUBPROCESO = " & id_subproc
txtborra = txtborra & " AND PARAMETRO1 = '" & txtport & "'"
ConAdo.Execute txtborra
txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   contar = DeterminaMaxRegSubproc(id_tabla)
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       tipopos = rmesa.Fields("TIPOPOS")
       fechareg = rmesa.Fields("FECHAREG")
       txtnompos = rmesa.Fields("NOMPOS")
       horareg = rmesa.Fields("HORAREG")
       cposicion = rmesa.Fields("CPOSICION")
       coperacion = rmesa.Fields("COPERACION")
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Cálculo esc de estres oper", txtport, txtportfr, tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   exito = False
End If

End Sub

Sub GenSubProcConsolEscEstres(ByVal fecha As Date, ByVal txtport As String, ByVal txtgrupoport As String, ByVal id_subproc As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim nveces() As Integer
Dim txtesc() As String
Dim i As Integer
Dim j As Integer
Dim l As Integer
Dim exito1 As Boolean
Dim noesc As Integer
Dim contar As Long
Dim txtfecha As String
Dim txtborra As String
Dim txtfiltro As String
Dim txtcadena As String

exito = True
noesc = UBound(MatFechasEstres, 1) + 3

ReDim txtesc(1 To noesc) As String
txtesc(1) = "3 desv est"
txtesc(2) = "Ad Hoc 1"
txtesc(3) = "Ad Hoc 2"
For i = 1 To UBound(MatFechasEstres, 1)
    txtesc(i + 3) = MatFechasEstres(i, 3)
Next i
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE ID_SUBPROCESO = " & id_subproc
txtborra = txtborra & " AND FECHAP = " & txtfecha
ConAdo.Execute txtborra
txtborra = "DELETE FROM " & TablaResEscEstresPort & " WHERE FECHA = " & txtfecha
txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
ConAdo.Execute txtborra

contar = DeterminaMaxRegSubproc(id_tabla)
For i = 1 To noesc
    contar = contar + 1
    txtcadena = CrearCadInsSub(fecha, id_subproc, contar, "Esc de estrés x subport", txtport, txtgrupoport, txtesc(i), "", "", "", "", "", "", "", "", "", id_tabla)
    ConAdo.Execute txtcadena
Next i
txtmsg = "El proceso finalizo correctamente"
exito = True
End Sub

Sub CalculoEscStress(ByVal fecha As Date, _
                     ByVal txtport As String, _
                     ByRef matpos() As propPosRiesgo, _
                     ByRef matposmd() As propPosMD, _
                     ByRef matposdiv() As propPosDiv, _
                     ByRef matposswaps() As propPosSwaps, _
                     ByRef matposfwd() As propPosFwd, _
                     ByRef matflswap() As estFlujosDeuda, _
                     ByRef matposdeuda() As propPosDeuda, _
                     ByRef matfldeuda() As estFlujosDeuda)
                     
Dim parval As ParamValPos
Dim mrvalflujo() As resValFlujo
Dim matprecios1() As resValIns
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim txtmsg As String
Dim exito As Boolean

'No de elementos en la posicion
'mfriesgo1      son las mfriesgo1 originales
'mfriesgo1 1    son las mfriesgo1 estresadas 1
'mfriesgo1 2    son las mfriesgo1 estresadas 2
ReDim mfriesgo2(1 To NoFactores, 1 To 1) As Double
'se guardan los resultados en MatVARExtremo
'se calculan los precios con la mfriesgo1 originales
Set parval = DeterminaPerfilVal("ESTRES")
MatPrecios = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg, exito)
noreg = UBound(matpos, 1)
For i = 1 To UBound(matEscEstres, 1)
    For j = 1 To NoFactores
         mfriesgo2(j, 1) = matEscEstres(i, j)
    Next j
'se calculan los precios con las mfriesgo1 estresadas
    ReDim MatVARExtremo(1 To noreg, 1 To 1) As Double
    matprecios1 = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, mfriesgo2, MatCurvasT, parval, mrvalflujo, txtmsg, exito)
    ReDim MatVARExtremo(1 To noreg, 1 To 1) As Double
    For j = 1 To noreg
        MatVARExtremo(j, 1) = matprecios1(j).mtm_sucio - MatPrecios(j).mtm_sucio
    Next j
    Call GuardarEscEstres(fecha, txtport, matNomEscEstres(i), matpos, MatVARExtremo)
Next i

End Sub

Sub CrearEscenariosTasas(ByVal fecha As Date, ByRef exito As Boolean)
exito = True
Dim indice As Long
Dim matx() As Variant
Dim matx1() As Double
Dim matrends() As Double
Dim siesfv As Boolean
Dim i As Integer
Dim j As Integer
Dim htiempo As Integer
Dim incfact As Double
Dim mata1() As Variant
Dim mata2() As Variant
Dim matfecha1() As Variant
Dim matfecha2() As Variant
Dim matfechas() As Date
Dim noesc As Integer
Dim fechaesc As Date
Dim fechax As Date
Dim matfr() As Variant
Dim indpos As Integer
Dim opcion As Integer
Dim noescestres As Integer
Dim matb() As Integer

noescestres = UBound(MatFechasEstres, 1) + 3
ReDim matEscEstres(1 To noescestres, 1 To NoFactores) As Double
ReDim matNomEscEstres(1 To noescestres) As String
For opcion = 1 To noescestres
    Select Case opcion
    Case 1  'desviaciones estandar
         noesc = 500
         htiempo = 1
         matNomEscEstres(opcion) = "3 desv est"
  'para esta parte se tienen que crear 2 vectores que guarden la
  'informacion de las volatilidades calculadas
         indpos = BuscarValorArray(fecha, MatFactRiesgo, 1)
         If indpos = 0 Then
            MsgBox "No se encuentra la fecha en el vector de tasas"
            Screen.MousePointer = 0
            Exit Sub
         End If
         ReDim matinctasa1(1 To NoFactores, 1 To 1) As Double
  'se obtiene una matriz de los factores de riesgo sin la fecha anexada
         matx = ExtraerSMatFR(indpos, noesc + htiempo, MatFactRiesgo, True, SiFactorRiesgo)
         matfechas = ConvArVtDT(ExtraeSubMatrizV(matx, 1, 1, 1 + htiempo, UBound(matx, 1)))
         matx1 = ConvArVtDbl(ExtraeSubMatrizV(matx, 2, UBound(matx, 2), 1, UBound(matx, 1)))
         Call GenRends3(matx1, htiempo, matfechas, matrends, matb)
         ReDim MatCv1(1 To NoFactores) As Double
   'se generan las desviaciones estandar
         For i = 1 To NoFactores
             MatCv1(i) = CVarianza2(ExtVecMatD(matrends, i, 0), 1, "c")
         Next i
         For i = 1 To NoFactores
              If MatCaracFRiesgo(i).tfactor = "TASA" Or MatCaracFRiesgo(i).tfactor = "SOBRETASA" Or MatCaracFRiesgo(i).tfactor = "TASA EXT" Or MatCaracFRiesgo(i).tfactor = "TCDS" Or MatCaracFRiesgo(i).tfactor = "UDI" Or MatCaracFRiesgo(i).tfactor = "TASA REF EXT" Or MatCaracFRiesgo(i).tfactor = "TASA REF" Or MatCaracFRiesgo(i).tfactor = "YIELD" Or MatCaracFRiesgo(i).tfactor = "YIELD IS" Then
                 incfact = (MatCv1(i)) ^ 0.5
              ElseIf MatCaracFRiesgo(i).tfactor = "TASA REAL" Or MatCaracFRiesgo(i).tfactor = "UMS" Then
                 incfact = (MatCv1(i)) ^ 0.5
              ElseIf MatCaracFRiesgo(i).tfactor = "IPyC" Or MatCaracFRiesgo(i).tfactor = "ACCION" Or MatCaracFRiesgo(i).tfactor = "T CAMBIO" Or MatCaracFRiesgo(i).tfactor = "INDICE" Then
                 incfact = (MatCv1(i)) ^ 0.5
              Else
                 MsgBox "No se ha definido incremento para el factor " & MatCaracFRiesgo(i).indFactor & " " & MatCaracFRiesgo(i).tfactor
              End If
              matinctasa1(i, 1) = Abs(MatFactR1(i, 1)) * 3 * incfact
              DoEvents
         Next i
    Case 2  'incrementos maximos ad hoc 1
         matNomEscEstres(opcion) = "Ad Hoc 1"
         htiempo = 1
         indpos = BuscarValorArray(fecha, MatFactRiesgo, 1)
         If indpos = 0 Then
            exito = False
            MsgBox "No se encuentra la fecha en las fechas para calculo de var"
            Exit Sub
         End If
         ReDim matinctasa1(1 To NoFactores, 1 To 1) As Double
         mata1 = LeerBaseValExtO(fecha, "max")
         mata2 = LeerBaseValExtO(fecha, "min")
         If UBound(mata1, 1) > 0 And UBound(mata2, 1) > 0 Then
            For i = 1 To NoFactores
                incfact = 0
                indice = BuscarValorArray(MatCaracFRiesgo(i).nomFactor & " " & MatCaracFRiesgo(i).plazo, mata1, 6)
                If indice <> 0 Then
                   If MatCaracFRiesgo(i).tfactor = "YIELD" Or MatCaracFRiesgo(i).tfactor = "TASA" Or MatCaracFRiesgo(i).tfactor = "SOBRETASA" Or MatCaracFRiesgo(i).tfactor = "TASA EXT" Or MatCaracFRiesgo(i).tfactor = "T CAMBIO" Or MatCaracFRiesgo(i).tfactor = "TASA N" Or MatCaracFRiesgo(i).tfactor = "UDI" Or MatCaracFRiesgo(i).tfactor = "TASA REAL" Or MatCaracFRiesgo(i).tfactor = "TASA REF EXT" Or MatCaracFRiesgo(i).tfactor = "TASA REF" Or MatCaracFRiesgo(i).tfactor = "YIELD IS" Then
                      incfact = mata1(indice, 5) 'incremento
                   ElseIf MatCaracFRiesgo(i).tfactor = "IPyC" Or MatCaracFRiesgo(i).tfactor = "ACCION" Or MatCaracFRiesgo(i).tfactor = "INDICE" Then
                      incfact = mata2(indice, 5) 'decremento
                   Else
                     MsgBox "no estoy estresando el factor " & MatCaracFRiesgo(i).indFactor & " " & MatCaracFRiesgo(i).tfactor
                   End If
                End If
                matinctasa1(i, 1) = Abs(MatFactR1(i, 1)) * incfact
                DoEvents
            Next i
         End If
    Case 3  'nuevo ad hoc 2
         matNomEscEstres(opcion) = "Ad Hoc 2"
         ReDim matinctasa1(1 To NoFactores, 1 To 1) As Double
         For i = 1 To NoFactores
             incfact = IncAplicarAdHoc2(i)
             matinctasa1(i, 1) = MatFactR1(i, 1) * incfact
             DoEvents
         Next i
    Case Is > 3
         matNomEscEstres(opcion) = MatFechasEstres(opcion - 3, 3)
         htiempo = 1
         fechaesc = MatFechasEstres(opcion - 3, 2)
         ReDim matinctasa1(1 To NoFactores, 1 To 1) As Double
         siesfv = EsFechaVaR(fechaesc)
         If siesfv <> 0 Then
      'se obtiene una matriz de los factores de riesgo sin la fecha anexada
            fechax = DetFechaFNoEsc(fechaesc, htiempo + 1)
            Call CrearMatFRiesgo2(fechax, fechaesc, matfr, "", exito)
            matx = ExtraerSMatFR(2, 2, matfr, True, SiFactorRiesgo)
            matfechas = ConvArVtDT(ExtraeSubMatrizV(matx, 1, 1, 1 + htiempo, UBound(matx, 1)))
            matx1 = ConvArVtDbl(ExtraeSubMatrizV(matx, 2, UBound(matx, 2), 1, UBound(matx, 1)))
            Call GenRends3(matx1, htiempo, matfechas, matrends, matb)
            For i = 1 To NoFactores
                If matb(1, i) = 1 Then
                   matinctasa1(i, 1) = Abs(MatFactR1(i, 1)) * matrends(1, i)
                Else
                   matinctasa1(i, 1) = matrends(1, i)
                End If
                DoEvents
            Next i
         Else
            MensajeProc = "No se cargo la fecha " & fechaesc
            exito = False
         End If
    End Select
    For i = 1 To NoFactores
        matEscEstres(opcion, i) = MatFactR1(i, 1) + matinctasa1(i, 1)
    Next i
Next opcion
End Sub

Function IncAplicarAdHoc2(ByVal i As Integer)
Dim incfact  As Double
  If MatCaracFRiesgo(i).nomFactor = "DESC IRS" Then
     incfact = 0.1
  ElseIf MatCaracFRiesgo(i).nomFactor = "CCMID" Then
     incfact = 0.05
  Else
  If MatCaracFRiesgo(i).tfactor = "TASA" Then
     incfact = 0.1
  ElseIf MatCaracFRiesgo(i).tfactor = "YIELD" Then
     incfact = 0.1
  ElseIf MatCaracFRiesgo(i).tfactor = "YIELD IS" Then
     incfact = 0.1
  ElseIf MatCaracFRiesgo(i).tfactor = "TASA REF" Then
     incfact = 0.1
  ElseIf MatCaracFRiesgo(i).tfactor = "TASA REF EXT" Then
     incfact = 0.05
  ElseIf MatCaracFRiesgo(i).tfactor = "SOBRETASA" Then
     incfact = 0.1
  ElseIf MatCaracFRiesgo(i).tfactor = "TASA EXT" Then
     incfact = 0.05
  ElseIf MatCaracFRiesgo(i).tfactor = "UMS" Then
     incfact = 0.05
  ElseIf MatCaracFRiesgo(i).tfactor = "TASA REAL" Then
     incfact = 0.1           'validado
  ElseIf MatCaracFRiesgo(i).tfactor = "ACCION" Then
     incfact = -0.1
  ElseIf MatCaracFRiesgo(i).tfactor = "IPyC" Then
     incfact = -0.1
  ElseIf MatCaracFRiesgo(i).tfactor = "T CAMBIO" Then
     incfact = 0.1            'validado
  ElseIf MatCaracFRiesgo(i).tfactor = "UDI" Then
     incfact = 0.01
  ElseIf MatCaracFRiesgo(i).tfactor = "INDICE" Then
     incfact = -0.1            'validado
  Else
    MsgBox "No estoy aplicando incremento al factor " & MatCaracFRiesgo(i).indFactor & "  " & MatCaracFRiesgo(i).tfactor
  End If
 End If
IncAplicarAdHoc2 = incfact
End Function

Function RealizarPartEsc(ByVal fecha As Date, ByVal fecha1 As Date, ByVal htiempo As Long) As Date()
Dim contar As Long
Dim año As Integer
Dim fechaa As Date
Dim fechab As Date
Dim mata() As Date


año = Year(fecha1)
Do While DateSerial(año, 1, 1) <= fecha
   contar = contar + 1
   ReDim Preserve mata(1 To 2, 1 To contar) As Date
   fechaa = DateSerial(año, 1, 1)
   mata(1, contar) = PBD1(fechaa, 1, "MX")
   fechab = Minimo(fecha, DateSerial(año, 12, 31))
   If NoLabMX(fechab) Then
      mata(2, contar) = PBD1(fechab, 1, "MX")
   Else
      mata(2, contar) = fechab
   End If
   año = año + 1
Loop
mata = MTranDt(mata)
RealizarPartEsc = mata
End Function

