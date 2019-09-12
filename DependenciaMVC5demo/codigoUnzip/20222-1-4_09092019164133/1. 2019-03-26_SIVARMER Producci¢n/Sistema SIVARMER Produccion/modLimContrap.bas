Attribute VB_Name = "modLimContrap"
Option Explicit

Sub GenerarSubpMaxExpFwds(ByVal fecha As Date, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim k As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Long
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim txtport As String
Dim esswap As Boolean
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim mata() As String
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc
    ConAdo.Execute txtborra
    contar = DeterminaMaxRegSubproc(id_tabla)
    mata = ObtenerContrapNoFinSwaps(fecha)
    noreg = UBound(mata, 1)
    txtcadena = "DELETE FROM " & TablaExpFwds & " WHERE FECHA  = " & txtfecha
    ConAdo.Execute txtcadena
    For i = 1 To noreg
        txtport = "Deriv Contrap " & mata(i, 1)
        txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE "
        txtfiltro2 = txtfiltro2 & " FECHA_PORT = " & txtfecha
        txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
        txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
        rmesa.Open txtfiltro1, ConAdo
        noreg2 = rmesa.Fields(0)
        rmesa.Close
        If noreg2 <> 0 Then
           rmesa.Open txtfiltro2, ConAdo
           For j = 1 To noreg2
               tipopos = rmesa.Fields("TIPOPOS").value
               fechareg = rmesa.Fields("FECHAREG").value
               txtnompos = rmesa.Fields("NOMPOS").value
               horareg = rmesa.Fields("HORAREG").value
               cposicion = rmesa.Fields("CPOSICION").value
               coperacion = rmesa.Fields("COPERACION").value
               esswap = DeterminaSiEsSwap(fechareg, coperacion)
               If Not esswap Then
                  contar = contar + 1
                  txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Max exp fwds", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, "", "", "", "", "", "", id_tabla)
                  ConAdo.Execute txtcadena
                  DoEvents
               End If
               rmesa.MoveNext
           Next j
           rmesa.Close
        End If
    Next i
   
End Sub

Sub GenerarLSubpLimC1(ByVal fecha0 As Date, ByVal fecha As Date, ByVal txtport As String, ByVal opcion_c As Integer, ByVal opc_fecha As Integer, ByVal fechax As Date, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim j As Integer
Dim k As Integer
Dim contar As Long
Dim noreg2 As Long
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim matf() As Date
Dim esswap As Boolean
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    matf = GenPartFechasEsc(fecha0, fecha, 100)
    txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE "
    txtfiltro2 = txtfiltro2 & " FECHA_PORT = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg2
           tipopos = rmesa.Fields("TIPOPOS").value
           fechareg = rmesa.Fields("FECHAREG").value
           txtnompos = rmesa.Fields("NOMPOS").value
           horareg = rmesa.Fields("HORAREG").value
           cposicion = rmesa.Fields("CPOSICION").value
           coperacion = rmesa.Fields("COPERACION").value
           esswap = DeterminaSiEsSwap(fechareg, coperacion)
          If esswap Then
              For k = 1 To UBound(matf, 1)
                  contar = contar + 1
                  txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Límites Contrap Swap 1", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, k, matf(k, 1), matf(k, 2), opcion_c, opc_fecha, fechax, id_tabla)
                  ConAdo.Execute txtcadena
                  DoEvents
                  Next k
           End If
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
   
End Sub

Sub GenerarLSubpLimCPosSim1(ByVal fecha0 As Date, ByVal tipopos As Integer, ByVal fecha As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal opc_max_min As Integer, ByVal opc_fecha As Integer, ByVal fecha1 As Date, ByVal id_proc As Integer, ByVal id_tabla As Integer)
    
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim k As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Long
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim matf() As Date
Dim txtport As String
Dim esswap As Boolean
Dim fechareg As Date
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc
    txtborra = txtborra & " AND PARAMETRO3 = '" & txtnompos & "'"
    txtborra = txtborra & " AND PARAMETRO5 = '" & cposicion & "'"
    txtborra = txtborra & " AND PARAMETRO6 = '" & coperacion & "'"
    
    ConAdo.Execute txtborra
    contar = DeterminaMaxRegSubproc(id_tabla)
    matf = GenPartFechasEsc(fecha0, fecha, 100)
    txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & "  WHERE "
    txtfiltro2 = txtfiltro2 & " TIPOPOS  = " & tipopos
    txtfiltro2 = txtfiltro2 & " AND FECHAREG  = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND NOMPOS = '" & txtnompos & "'"
    txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion
    txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & coperacion & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg2
           tipopos = rmesa.Fields("TIPOPOS").value
           fechareg = rmesa.Fields("FECHAREG").value
           txtnompos = rmesa.Fields("NOMPOS").value
           horareg = rmesa.Fields("HORAREG").value
           cposicion = rmesa.Fields("CPOSICION").value
           coperacion = rmesa.Fields("COPERACION").value
           txtcadena = "DELETE FROM " & TablaLimContrap1 & " WHERE FECHA  = " & txtfecha & " AND COPERACION = '" & coperacion & "'"
           ConAdo.Execute txtcadena
           txtcadena = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc & " AND PARAMETRO6 = '" & coperacion & "'"
           ConAdo.Execute txtcadena
           For k = 1 To UBound(matf, 1)
               contar = contar + 1
               txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Límites Contrap Swap 1", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, k, matf(k, 1), matf(k, 2), opc_max_min, opc_fecha, fecha1, id_tabla)
               ConAdo.Execute txtcadena
               DoEvents
           Next k
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
  
End Sub

Sub GenerarLSubpLimContrap1(ByVal fecha0 As Date, ByVal fecha As Date, ByVal id_contrap As Integer, ByVal opcion_c As Integer, ByVal id_proc As Integer, ByVal id_tabla As Integer)
    
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim k As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Long
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim matf() As Date
Dim txtport As String
Dim esswap As Boolean
Dim tipopos As Integer
Dim fechareg As Date
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim txtnompos As String
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc
    ConAdo.Execute txtborra
    matf = GenPartFechasEsc(fecha0, fecha, 100)
    contar = DeterminaMaxRegSubproc(id_tabla)
    txtport = "Deriv Contrap " & id_contrap
    txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE "
    txtfiltro2 = txtfiltro2 & " FECHA_PORT = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg2
           tipopos = rmesa.Fields("TIPOPOS").value
           fechareg = rmesa.Fields("FECHAREG").value
           txtnompos = rmesa.Fields("NOMPOS").value
           horareg = rmesa.Fields("HORAREG").value
           cposicion = rmesa.Fields("CPOSICION").value
           coperacion = rmesa.Fields("COPERACION").value
           txtcadena = "DELETE FROM " & TablaLimContrap1 & " WHERE FECHA  = " & txtfecha & " AND COPERACION = '" & coperacion & "'"
           ConAdo.Execute txtcadena
           For k = 1 To UBound(matf, 1)
               contar = contar + 1
               txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Límites Contrap Swap 1", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, k, matf(k, 1), matf(k, 2), opcion_c, "", "", id_tabla)
               ConAdo.Execute txtcadena
               DoEvents
           Next k
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
  
End Sub

Sub GenerarLSubpLimContrap2(ByVal fecha0 As Date, ByVal fecha As Date, ByVal id_contrap As Integer, ByVal opcionc As Integer, ByVal id_proc As Integer, ByVal id_tabla As Integer)
    
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim k As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Long
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim matf() As Date
Dim txtport As String
Dim esswap As Boolean
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    'txtborra = "DELETE FROM " & TablaSubProcesos & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc
    'conAdo.Execute txtborra
    contar = DeterminaMaxRegSubproc(id_tabla)
    matf = GenPartFechasEsc(fecha0, fecha, 100)
    txtport = "Deriv Contrap " & id_contrap
    txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE "
    txtfiltro2 = txtfiltro2 & " FECHA_PORT = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg2
           tipopos = rmesa.Fields("TIPOPOS").value
           fechareg = rmesa.Fields("FECHAREG").value
           txtnompos = rmesa.Fields("NOMPOS").value
           horareg = rmesa.Fields("HORAREG").value
           cposicion = rmesa.Fields("CPOSICION").value
           coperacion = rmesa.Fields("COPERACION").value
           txtcadena = "DELETE FROM " & TablaLimContrap1 & " WHERE FECHA  = " & txtfecha & " AND COPERACION = '" & coperacion & "'"
           ConAdo.Execute txtcadena
           For k = 1 To UBound(matf, 1)
               contar = contar + 1
               txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Límites Contrap Swap 2", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, k, matf(k, 1), matf(k, 2), opcionc, "", "", id_tabla)
               ConAdo.Execute txtcadena
               DoEvents
           Next k
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
  
End Sub

Sub GenerarLSubpLimCPosSim2(ByVal fecha0 As Date, ByVal tipopos As Integer, ByVal fecha As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal opcionc As Integer, ByVal opc_fecha As Integer, ByVal fecha1 As Date, ByVal id_proc As Integer, ByVal id_tabla As Integer)
    
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim k As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Long
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim matf() As Date
Dim txtport As String
Dim esswap As Boolean
Dim fechareg As Date
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    matf = GenPartFechasEsc(fecha0, fecha, 100)
    txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & "  WHERE "
    txtfiltro2 = txtfiltro2 & " TIPOPOS = " & tipopos
    txtfiltro2 = txtfiltro2 & " AND FECHAREG = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND NOMPOS  = '" & txtnompos & "'"
    txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion
    txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & coperacion & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg2
           tipopos = rmesa.Fields("TIPOPOS").value
           fechareg = rmesa.Fields("FECHAREG").value
           txtnompos = rmesa.Fields("NOMPOS").value
           horareg = rmesa.Fields("HORAREG").value
           cposicion = rmesa.Fields("CPOSICION").value
           coperacion = rmesa.Fields("COPERACION").value
           txtcadena = "DELETE FROM " & TablaLimContrap2 & " WHERE FECHA  = " & txtfecha & " AND COPERACION = '" & coperacion & "'"
           ConAdo.Execute txtcadena
           txtcadena = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc & " AND PARAMETRO3 = '" & txtnompos & "'"
           txtcadena = txtcadena & " AND PARAMETRO5 = '" & cposicion & "'"
           txtcadena = txtcadena & " AND PARAMETRO6 = '" & coperacion & "'"
           ConAdo.Execute txtcadena
           For k = 1 To UBound(matf, 1)
               contar = contar + 1
               txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Límites Contrap Swap 2", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, k, matf(k, 1), matf(k, 2), opcionc, opc_fecha, fecha1, id_tabla)
               ConAdo.Execute txtcadena
               DoEvents
           Next k
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
    
End Sub



Function DeterminaSiEsSwap(ByVal fecha As Date, ByVal coperacion As String)
Dim txtfiltro As String
Dim txtfecha As String
Dim noreg As Long
Dim rmesa2 As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtfiltro = "SELECT COUNT(*) FROM " & TablaPosSwaps & " WHERE "
    txtfiltro = txtfiltro & " FECHAREG = " & txtfecha
    txtfiltro = txtfiltro & " AND COPERACION = '" & coperacion & "'"
    txtfiltro = txtfiltro & " AND TIPOPOS = 1"
    rmesa2.Open txtfiltro, ConAdo
    noreg = rmesa2.Fields(0)
    rmesa2.Close
    If noreg <> 0 Then
       DeterminaSiEsSwap = True
    Else
       DeterminaSiEsSwap = False
    End If
End Function

Sub GenLSubConsolLimC1(ByVal fecha As Date, ByVal txtport As String, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim j As Integer
Dim contar As Long
Dim noreg2 As Integer
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE "
    txtfiltro2 = txtfiltro2 & " FECHA_PORT = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       ReDim mata(1 To noreg2, 1 To 3) As Variant
       For j = 1 To noreg2
           tipopos = rmesa.Fields("TIPOPOS").value
           fechareg = rmesa.Fields("FECHAREG").value
           txtnompos = rmesa.Fields("NOMPOS").value
           horareg = rmesa.Fields("HORAREG").value
           cposicion = rmesa.Fields("CPOSICION").value
           coperacion = rmesa.Fields("COPERACION").value
           contar = contar + 1
           txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Consol Calc Lim Contrap 1", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, "", "", "", "", "", "", id_tabla)
           ConAdo.Execute txtcadena
           rmesa.MoveNext
       Next j
       rmesa.Close
   End If
   
End Sub

Function DetermTablaSubproc(ByVal opcion As Integer)
If opcion = 1 Then
   DetermTablaSubproc = TablaSubProcesos1
ElseIf opcion = 2 Then
   DetermTablaSubproc = TablaSubProcesos2
ElseIf opcion = 3 Then
   DetermTablaSubproc = TablaSubProcesos3
End If
End Function

Function DetermTablaProc(ByVal opcion As Integer)
If opcion = 1 Then
   DetermTablaProc = TablaProcesos1
ElseIf opcion = 2 Then
   DetermTablaProc = TablaProcesos2
End If
End Function


Sub GenLSubConsolLimContrap1(ByVal fecha As Date, ByVal id_contrap As String, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Integer
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim txtport As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim txttabla As String
Dim mata() As Variant
Dim rmesa As New ADODB.recordset

    txttabla = DetermTablaSubproc(id_tabla)
    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtborra = "DELETE FROM " & txttabla & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc
    ConAdo.Execute txtborra
    contar = DeterminaMaxRegSubproc(id_tabla)
        txtport = "Swap Contrap " & id_contrap
        txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE "
        txtfiltro2 = txtfiltro2 & " FECHA_PORT = " & txtfecha
        txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
        txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
        rmesa.Open txtfiltro1, ConAdo
        noreg2 = rmesa.Fields(0)
        rmesa.Close
        If noreg2 <> 0 Then
           rmesa.Open txtfiltro2, ConAdo
           ReDim mata(1 To noreg2, 1 To 3) As Variant
           For j = 1 To noreg2
               tipopos = rmesa.Fields("TIPOPOS").value
               fechareg = rmesa.Fields("FECHAREG").value
               txtnompos = rmesa.Fields("NOMPOS").value
               horareg = rmesa.Fields("HORAREG").value
               cposicion = rmesa.Fields("CPOSICION").value
               coperacion = rmesa.Fields("COPERACION").value
               contar = contar + 1
               txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Consol Calc Lim Contrap 1", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, "", "", "", "", "", "", id_tabla)
               ConAdo.Execute txtcadena
               rmesa.MoveNext
           Next j
           rmesa.Close
       End If
   
End Sub

Sub GenLSubConsolLimContrap2(ByVal fecha As Date, ByVal id_contrap As String, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Integer
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim txtport As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc
    ConAdo.Execute txtborra
    contar = DeterminaMaxRegSubproc(id_tabla)
        txtport = "Swap Contrap " & id_contrap
        txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE "
        txtfiltro2 = txtfiltro2 & " FECHA_PORT = " & txtfecha
        txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
        txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
        rmesa.Open txtfiltro1, ConAdo
        noreg2 = rmesa.Fields(0)
        rmesa.Close
        If noreg2 <> 0 Then
           rmesa.Open txtfiltro2, ConAdo
           ReDim mata(1 To noreg2, 1 To 3) As Variant
           For j = 1 To noreg2
               tipopos = rmesa.Fields("TIPOPOS").value
               fechareg = rmesa.Fields("FECHAREG").value
               txtnompos = rmesa.Fields("NOMPOS").value
               horareg = rmesa.Fields("HORAREG").value
               cposicion = rmesa.Fields("CPOSICION").value
               coperacion = rmesa.Fields("COPERACION").value
               contar = contar + 1
               txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Consol Calc Lim Contrap 2", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, "", "", "", "", "", "", id_tabla)
               ConAdo.Execute txtcadena
               rmesa.MoveNext
           Next j
           rmesa.Close
       End If
   
End Sub

Sub GenLSubConsolLimCPosSim1(ByVal tipopos As Integer, ByVal fecha As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal opcionc As Integer, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Integer
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim txtport As String
Dim fechareg As Date

Dim rmesa As New ADODB.recordset

    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & " WHERE "
    txtfiltro2 = txtfiltro2 & " TIPOPOS = " & tipopos
    txtfiltro2 = txtfiltro2 & " AND FECHAREG = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND NOMPOS  = '" & txtnompos & "'"
    txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion
    txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & coperacion & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg2
           tipopos = rmesa.Fields("TIPOPOS").value
           fechareg = rmesa.Fields("FECHAREG").value
           txtnompos = rmesa.Fields("NOMPOS").value
           horareg = rmesa.Fields("HORAREG").value
           cposicion = rmesa.Fields("CPOSICION").value
           coperacion = rmesa.Fields("COPERACION").value
           txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc & " AND PARAMETRO3 = '" & txtnompos & "'"
           txtborra = txtborra & " AND PARAMETRO5 = '" & cposicion & "'"
           txtborra = txtborra & " AND PARAMETRO6 = '" & coperacion & "'"
           ConAdo.Execute txtborra
           contar = contar + 1
           txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Consol Calc Lim Contrap 1", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, opcionc, "", "", "", "", "", id_tabla)
           ConAdo.Execute txtcadena
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
 
End Sub

Sub GenLSubConsolLimCPosSim2(ByVal tipopos As Integer, ByVal fecha As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal opcionc As Integer, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Integer
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim txtport As String
Dim fechareg As Date
Dim txttabla As String
Dim rmesa As New ADODB.recordset

txttabla = DetermTablaSubproc(id_tabla)
    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND ID_SUBPROCESO = " & id_proc
    txtborra = txtborra & " AND PARAMETRO3 = '" & txtnompos & "'"
    txtborra = txtborra & " AND PARAMETRO5 = '" & cposicion & "'"
    txtborra = txtborra & " AND PARAMETRO6 = '" & coperacion & "'"
    ConAdo.Execute txtborra
    contar = DeterminaMaxRegSubproc(id_tabla)
    txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & "  WHERE "
    txtfiltro2 = txtfiltro2 & " TIPOPOS = " & tipopos
    txtfiltro2 = txtfiltro2 & " AND FECHAREG  = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND NOMPOS = '" & txtnompos & "'"
    txtfiltro2 = txtfiltro2 & " AND CPOSICION  = " & cposicion
    txtfiltro2 = txtfiltro2 & " AND COPERACION  = '" & coperacion & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg2
           tipopos = rmesa.Fields("TIPOPOS").value
           fechareg = rmesa.Fields("FECHAREG").value
           txtnompos = rmesa.Fields("NOMPOS").value
           horareg = rmesa.Fields("HORAREG").value
           cposicion = rmesa.Fields("CPOSICION").value
           coperacion = rmesa.Fields("COPERACION").value
           contar = contar + 1
           txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Consol Calc Lim Contrap 2", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, opcionc, "", "", "", "", "", id_tabla)
           ConAdo.Execute txtcadena
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
 
End Sub

Sub GenLSubConsolLimC2(ByVal fecha As Date, ByVal txtport As String, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim contar As Long
Dim noreg As Integer
Dim noreg2 As Integer
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset
        
    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE "
    txtfiltro2 = txtfiltro2 & " FECHA_PORT = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
               ReDim mata(1 To noreg2, 1 To 3) As Variant
       For j = 1 To noreg2
           tipopos = rmesa.Fields("TIPOPOS").value
           fechareg = rmesa.Fields("FECHAREG").value
           txtnompos = rmesa.Fields("NOMPOS").value
           horareg = rmesa.Fields("HORAREG").value
           cposicion = rmesa.Fields("CPOSICION").value
           coperacion = rmesa.Fields("COPERACION").value
           contar = contar + 1
           txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Consol Calc Lim Contrap 2", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, "", "", "", "", "", "", id_tabla)
           ConAdo.Execute txtcadena
           rmesa.MoveNext
       Next j
       rmesa.Close
    End If
End Sub

Sub GenerarLSubpLimC(ByVal fecha0 As Date, ByVal fecha As Date, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim i As Integer, j As Integer
Dim contar As Integer
Dim noreg As Integer
Dim hinicio As String
Dim hfinal As String
Dim txtfiltro As String
Dim txtborra As String
Dim txtcadena As String
Dim matf() As Date
Dim mata() As String

        frmProgreso.Show
        mata = ObtenerContrapNoFinSwaps(fecha)
        noreg = UBound(mata, 1)
        txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
        hinicio = Format$(0, "HH:MM:SS")
        hfinal = Format$(0, "HH:MM:SS")
        txtborra = "DELETE FROM " & DetermTablaSubproc(id_tabla) & " WHERE FECHAP = " & txtfecha & " AND IDTAREA = 100"
        ConAdo.Execute txtborra
        contar = DeterminaMaxRegSubproc(id_tabla)
        matf = GenPartFechasEsc(fecha0, fecha, 100)
        For i = 1 To noreg
            For j = 1 To UBound(matf, 1)
                contar = contar + 1
                txtcadena = CrearCadInsSub(fecha, 100, contar, "Límites Contraparte", mata(i, 1), j, matf(j, 1), matf(j, 2), "", "", "", "", "", "", "", "", id_tabla)
                ConAdo.Execute txtcadena
            Next j
        Next i

        Unload frmProgreso
   
End Sub

Function GenPartFechasEsc(ByVal fecha0 As Date, ByVal fecha As Date, ByVal ndias As Integer)
Dim matf() As Date
Dim contar As Integer
Dim fecha1 As Date
Dim fecha2 As Date
Dim fechaa As Date
Dim fechab As Date
Dim indice As Integer

ReDim matf(1 To 2, 1 To 1) As Date
contar = 0
ReDim Preserve matf(1 To 2, 1 To 1) As Date
fecha1 = fecha0
indice = BuscarValorArray(fecha1, MatFechasVaR, 1)
If indice <> 0 Then
Do While True
   contar = contar + 1
   indice = indice + ndias
   fecha2 = Minimo(MatFechasVaR(Minimo(indice, UBound(MatFechasVaR)), 1), fecha)
   If fecha >= fecha1 And fecha <= fecha2 Then
      ReDim Preserve matf(1 To 2, 1 To contar) As Date
      matf(1, contar) = fecha1
      matf(2, contar) = fecha2
      GenPartFechasEsc = MTranDt(matf)
      Exit Function
   ElseIf fecha > fecha2 Then
      ReDim Preserve matf(1 To 2, 1 To contar) As Date
      matf(1, contar) = fecha1
      matf(2, contar) = fecha2
   ElseIf fecha < fecha0 - 1 Then
      ReDim matf(0 To 0, 0 To 0) As Date
      GenPartFechasEsc = matf
      Exit Function
   End If
   fecha1 = MatFechasVaR(indice + 1, 1)
Loop
Else
 MsgBox "No es una fecha valida"
End If
End Function


Function GenFechasEscEstres(ByVal dtfecha As Date) As Date()
Dim matf() As Date
Dim contar As Integer
Dim año As Integer
Dim fecha As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim fechaa As Date
Dim fechab As Date

ReDim matf(1 To 2, 1 To 1) As Date
año = 1996
contar = 0
Do While True
   contar = contar + 1
   fecha1 = DateSerial(año, 12, 31)
   fecha2 = DateSerial(año + 1, 12, 31)
   ReDim Preserve matf(1 To 2, 1 To contar) As Date
   If dtfecha > fecha1 And dtfecha <= fecha2 Then
      fechaa = PBD1(fecha1, 1, "MX")
      fechab = PBD1(fecha, 1, "MX")
      matf(1, contar) = fechaa
      matf(2, contar) = fechab
      GenFechasEscEstres = MTranDt(matf)
      Exit Function
   ElseIf dtfecha > fecha2 Then
      fechaa = PBD1(fecha1, 1, "MX")
      fechab = PBD1(fecha2, 1, "MX")
      matf(1, contar) = fechaa
      matf(2, contar) = fechab
   ElseIf dtfecha < #12/31/1996# Then
      ReDim matf(0 To 0, 0 To 0) As Date
      GenFechasEscEstres = matf
      Exit Function
   End If
   año = año + 1
Loop

End Function

Sub GenerarLSubpLimC2(ByVal fecha0 As Date, ByVal fecha As Date, ByVal txtport As String, ByVal opcion_c As Integer, ByVal opc_fecha As Integer, ByVal fechax As Date, ByVal id_proc As Integer, ByVal id_tabla As Integer)
Dim txtfecha As String
Dim j As Integer
Dim k As Integer
Dim contar As Long
Dim noreg2 As Long
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim matf() As Date
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim esswap As Boolean
Dim txttabla As String
Dim rmesa As New ADODB.recordset

    txttabla = DetermTablaSubproc(id_tabla)
    txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    contar = DeterminaMaxRegSubproc(id_tabla)
    matf = GenPartFechasEsc(fecha0, fecha, 100)
    txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE "
    txtfiltro2 = txtfiltro2 & " FECHA_PORT = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    If noreg2 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       ReDim mata(1 To noreg2, 1 To 3) As Variant
       For j = 1 To noreg2
           tipopos = rmesa.Fields("TIPOPOS").value
           fechareg = rmesa.Fields("FECHAREG").value
           txtnompos = rmesa.Fields("NOMPOS").value
           horareg = rmesa.Fields("HORAREG").value
           cposicion = rmesa.Fields("CPOSICION").value
           coperacion = rmesa.Fields("COPERACION").value
           esswap = DeterminaSiEsSwap(fechareg, coperacion)
           If esswap Then
              For k = 1 To UBound(matf, 1)
                  contar = contar + 1
                  txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Límites Contraparte 2", tipopos, fechareg, txtnompos, horareg, cposicion, coperacion, k, matf(k, 1), matf(k, 2), opcion_c, opc_fecha, fechax, id_tabla)
                  ConAdo.Execute txtcadena
              Next k
           End If
           rmesa.MoveNext
       Next j
       rmesa.Close
     End If
  End Sub

Sub CalcLimContrapSwap1(ByVal dtfecha As Date, ByVal tipopos As Integer, ByVal fechareg As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal id_grupoc As Integer, ByVal dtfecha1 As Date, ByVal dtfecha2 As Date, ByVal opc_max_min As Integer, ByVal opc_fecha As Integer, ByVal fecha0 As Date, ByRef txtmsg As String, ByRef final As Boolean, ByRef bl_exito As Boolean)
If ActivarControlErrores Then
   On Error GoTo hayerror
End If
    Dim matpos() As New propPosRiesgo
    Dim matposmd() As New propPosMD
    Dim matposdiv() As New propPosDiv
    Dim matposswaps() As New propPosSwaps
    Dim matposfwd() As New propPosFwd
    Dim matposdeuda() As New propPosDeuda
    Dim matflswap() As New estFlujosDeuda
    Dim matfldeuda() As New estFlujosDeuda
    Dim parval As New ParamValPos
    Dim mat_fr() As Variant
    Dim mat_f_fr() As Date
    Dim nofechasfr As Integer
    Dim nofechasval As Integer
    Dim i As Integer, j As Integer
    Dim ll As Integer
    Dim matx() As Variant
    Dim suma As Double
    Dim sumaact As Double
    Dim sumapas As Double
    Dim m_fechas_val() As Date
    Dim mprecio() As resValIns
    Dim valmax As Double
    Dim valactmax As Double
    Dim valpasmax As Double
    Dim ffrvalmax As Date
    Dim f_val_max As Date
    Dim mattxt() As String
    Dim mFactR() As Double
    Dim mrvalflujo() As resValFlujo
    Dim txtmsg2 As String
    Dim exito As Boolean
    Dim exito2 As Boolean
    Dim exito3 As Boolean
    Dim txtmsg0 As String
    Dim txtmsg3 As String
    bl_exito = False
    mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
    Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
   'se crea la posicion en funcion de la contraparte seleccionada
    If UBound(matpos, 1) > 0 Then
       Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
       If exito2 Then
          SiIncTasaCVig = False
          mat_f_fr = LeerFechasVaR(dtfecha1, dtfecha2)
          m_fechas_val = DetFechasCalculo(dtfecha, fecha0, 1, matpos, matposswaps, matposfwd, opc_fecha)
          nofechasfr = UBound(mat_f_fr, 1)
          nofechasval = UBound(m_fechas_val, 1)
          Call CrearMatFRiesgo2(mat_f_fr(1, 1), mat_f_fr(nofechasfr, 1), MatFactRiesgo, txtmsg2, exito)
        'mat_fr tiene fechas
          mat_fr = ExtraeSubMatV(MatFactRiesgo, 1, UBound(MatFactRiesgo, 2), 1, nofechasfr)
        'se anexan las caracteristicas adicionales desde la tabla valuacion
          Set parval = DeterminaPerfilVal("LCONTRAPARTE")
        'Se carga la estructura de tasas para ese día de la matriz vector tasas
          ReDim matval(1 To nofechasfr, 1 To nofechasval) As Double
          ReDim matvalact(1 To nofechasfr, 1 To nofechasval) As Double
          ReDim matvalpas(1 To nofechasfr, 1 To nofechasval) As Double
          For i = 1 To nofechasval
              For j = 1 To nofechasfr
                  mFactR = ExtVecFactRiesgo(j, mat_fr)
                  mprecio = CalcValuacion(m_fechas_val(i), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, mFactR, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
                  suma = 0
                  sumaact = 0
                  sumapas = 0
                  For ll = 1 To UBound(mprecio, 1)
                      suma = suma + mprecio(ll).mtm_sucio
                      sumaact = sumaact + mprecio(ll).ps_activa
                      sumapas = sumapas + mprecio(ll).ps_pasiva
                  Next ll
                  matval(j, i) = suma
                  matvalact(j, i) = sumaact
                  matvalpas(j, i) = sumapas
                  If i = 1 And j = 1 Then
                     ffrvalmax = mat_f_fr(j, 1)
                     f_val_max = m_fechas_val(i)
                     valmax = suma
                  Else
                     If opc_max_min = 0 Then
                        If valmax < suma Then
                           valmax = suma
                           valactmax = sumaact
                           valpasmax = sumapas
                           f_val_max = m_fechas_val(i)
                           ffrvalmax = mat_f_fr(j, 1)
                        End If
                     ElseIf opc_max_min = 1 Then
                        If valmax > suma Then
                           valmax = suma
                           valactmax = sumaact
                           valpasmax = sumapas
                           f_val_max = m_fechas_val(i)
                           ffrvalmax = mat_f_fr(j, 1)
                        End If
                     End If
                  End If
              Next j
              AvanceProc = i / nofechasval
              MensajeProc = "Operacion " & coperacion & " Grupo " & id_grupoc & " Etapa 1 " & Format$(AvanceProc, "##0.00 %")
              DoEvents
          Next i
          Call GuardarResCLimContrap1(dtfecha, coperacion, id_grupoc, opc_max_min, dtfecha1, dtfecha2, ffrvalmax, f_val_max, valmax, valactmax, valpasmax, nofechasval, nofechasfr, m_fechas_val, mat_f_fr, matval, matvalact, matvalpas)
          txtmsg = "El proceso finalizo correctamente"
          final = True
          bl_exito = True
          SiIncTasaCVig = True
       Else
          txtmsg = txtmsg2
          final = True
          bl_exito = False
       End If
    Else
        txtmsg = "No se encontro la operacion " & coperacion & " a esta fecha"
        final = True
        bl_exito = False
    End If
Exit Sub
hayerror:
 MsgBox error(Err())
bl_exito = False
End Sub

Sub GuardarResCLimContrap1(ByVal fecha As Date, ByVal coperacion As String, ByVal id_grupoc As Integer, ByVal opcion As Integer, ByVal dtfecha1 As Date, ByVal dtfecha2 As Date, ffrvalmax, f_val_max, valmax, valactmax, valpasmax, nofechasval, nofechasfr, matfechasval, mat_f_fr, matval, matvalact, matvalpas)
Dim i As Long
Dim j As Long
Dim txtcadena1 As String
Dim txtcadena2 As String
Dim txtcadena3 As String
Dim txtcadena4 As String
Dim txtcadena5 As String
          txtcadena1 = ""
          txtcadena2 = ""
          txtcadena3 = ""
          txtcadena4 = ""
          txtcadena5 = ""
          For i = 1 To nofechasval
              txtcadena1 = txtcadena1 & matfechasval(i) & ","
          Next i
          For i = 1 To nofechasfr
              txtcadena2 = txtcadena2 & mat_f_fr(i, 1) & ","
          Next i
          For i = 1 To nofechasval
              For j = 1 To nofechasfr
                  txtcadena3 = txtcadena3 & matval(j, i) & ","
                  txtcadena4 = txtcadena4 & matvalact(j, i) & ","
                  txtcadena5 = txtcadena5 & matvalpas(j, i) & ","
              Next j
          Next i
          RegResLimC1.AddNew
          RegResLimC1.Fields(0).value = CLng(fecha)
          RegResLimC1.Fields("COPERACION").value = coperacion
          RegResLimC1.Fields("ID_GRUPOC").value = id_grupoc
          RegResLimC1.Fields("ID_CALCULO").value = opcion
          RegResLimC1.Fields("FECHA1").value = dtfecha1
          RegResLimC1.Fields("FECHA2").value = dtfecha2
          RegResLimC1.Fields("FESCMAX").value = ffrvalmax
          RegResLimC1.Fields("FVALMAX").value = f_val_max
          RegResLimC1.Fields("VALMAX").value = valmax
          RegResLimC1.Fields("VALACTMAX").value = valactmax
          RegResLimC1.Fields("VALPASMAX").value = valpasmax
          Call GuardarElementoClob(txtcadena1, RegResLimC1, "H_FECHAS1")
          Call GuardarElementoClob(txtcadena2, RegResLimC1, "H_FECHAS2")
          Call GuardarElementoClob(txtcadena3, RegResLimC1, "H_VALMAX")
          Call GuardarElementoClob(txtcadena4, RegResLimC1, "H_VALMAXACT")
          Call GuardarElementoClob(txtcadena5, RegResLimC1, "H_VALMAXPAS")
          RegResLimC1.Update
End Sub


Sub GuardarElementoClob(ByVal txtcadena As String, ByRef regis As ADODB.recordset, ByVal campo As String)
Dim largo As Long
Dim numbloques As Long
Dim leftover As Long
Dim txttexto As String
Dim i As Long
    largo = Len(txtcadena)
    numbloques = Int(largo / BlockSize)
    leftover = largo Mod BlockSize
    For i = 1 To numbloques
        txttexto = Mid(txtcadena, (i - 1) * BlockSize + 1, BlockSize)
        regis(campo).AppendChunk txttexto
    Next i
    If leftover <> 0 Then
       txttexto = Mid(txtcadena, numbloques * BlockSize + 1, leftover)
       regis(campo).AppendChunk txttexto
    End If
End Sub

Sub CalcLimContrapSwap2(ByVal dtfecha As Date, ByVal tipopos As Integer, ByVal fechareg As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal idgrupoc As Integer, ByVal dtfecha1 As Date, ByVal dtfecha2 As Date, ByVal opc_calc As Integer, ByVal opc_fecha As Integer, ByVal fecha0 As Date, ByVal id_procp As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef final As Boolean, ByRef bl_exito As Boolean)
    Dim matpos() As New propPosRiesgo
    Dim matposmd() As New propPosMD
    Dim matposdiv() As New propPosDiv
    Dim matposswaps() As New propPosSwaps
    Dim matposfwd() As New propPosFwd
    Dim matposdeuda() As New propPosDeuda
    Dim matflswap() As New estFlujosDeuda
    Dim matfldeuda() As New estFlujosDeuda
    Dim parval As New ParamValPos
    Dim matfr() As Variant
    Dim MatFactR1() As Double
    Dim matrends()  As Double
    Dim noreg1 As Integer
    Dim nofechasval As Integer
    Dim i As Integer, j As Integer
    Dim indice1 As Integer
    Dim indice2 As Integer
    Dim matx() As Variant
    Dim matx1() As Double
    Dim txtpossim As String
    Dim suma As Double
    Dim sumaact As Double
    Dim sumapas As Double
    Dim matfechas1() As Date
    Dim matfechassh() As Date
    Dim mprecio() As resValIns
    Dim ll As Integer
    Dim noesc As Integer
    Dim fcurvavalmax As Date
    Dim ffutvalmax1 As Date
    Dim valmax1 As Double
    Dim valmax As Double
    Dim valactmax As Double
    Dim valpasmax As Double
    Dim fescvalmax As Date
    Dim f_val_max As Date
    Dim txtfecha1 As String
    Dim txtcadena As String
    Dim txtcadena1 As String
    Dim txtcadena2 As String
    Dim txtcadena3 As String
    Dim txtcadena4 As String
    Dim txtcadena5 As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim noreg12 As Integer
    Dim noreg3 As Integer
    Dim fcurvavalmax1 As Date
    Dim MatFactoresR() As Double
    Dim mrvalflujo() As resValFlujo
    Dim mattxt() As String
    Dim exito As Boolean
    Dim exito1 As Boolean
    Dim exito2 As Boolean
    Dim txtmsg2 As String
    Dim matb() As Integer
    Dim txttabla As String
    Dim exito3 As Boolean
    Dim txtmsg0 As String
    Dim txtmsg3 As String
    Dim rmesa As New ADODB.recordset
    txttabla = DetermTablaSubproc(id_tabla)
   'se crea la posicion en funcion de la contraparte seleccionada
    txtfecha1 = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    txtfiltro2 = "SELECT * FROM " & TablaResLimContrap & " WHERE FECHA = " & txtfecha1 & " AND COPERACION = '" & coperacion & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM  (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg12 = rmesa.Fields(0)
    rmesa.Close
    If noreg12 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       fcurvavalmax1 = rmesa.Fields("FESCMAX1")
       ffutvalmax1 = rmesa.Fields("FVALMAX1")
       valmax1 = rmesa.Fields("VALMAX1")
       rmesa.Close
       txtfiltro1 = "SELECT COUNT(*) FROM " & txttabla & " WHERE FECHAP = " & txtfecha1
       txtfiltro1 = txtfiltro1 & " AND ID_SUBPROCESO = " & id_procp & " AND PARAMETRO6 = '" & coperacion & "' AND FINALIZADO = 'N'"
       rmesa.Open txtfiltro1, ConAdo
       noreg3 = rmesa.Fields(0)
       rmesa.Close
       If noreg3 = 0 Then
          If noreg12 <> 0 Then
             mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
             Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
             If UBound(matpos, 1) > 0 Then
                Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
                If exito2 Then
                   SiIncTasaCVig = False
                   Call VerifCargaFR(dtfecha1, dtfecha2)
                   indice1 = 0
                   indice2 = 0
                   Do While indice1 = 0
                      indice1 = BuscarValorArray(dtfecha1, MatFactRiesgo, 1)
                      dtfecha1 = dtfecha1 + 1
                   Loop
                   Do While indice2 = 0
                      indice2 = BuscarValorArray(dtfecha2, MatFactRiesgo, 1)
                      dtfecha2 = dtfecha2 - 1
                   Loop
                   matfr = ExtraeSubMatV(MatFactRiesgo, 1, UBound(MatFactRiesgo, 2), indice1, indice2)
                   matfechas1 = DetFechasCalculo(dtfecha, fecha0, 1, matpos, matposswaps, matposfwd, opc_fecha)
                   nofechasval = UBound(matfechas1, 1)
                 'se anexan las caracteristicas adicionales desde la tabla valuacion
                   bl_exito = True
                   Set parval = DeterminaPerfilVal("LCONTRAPARTE")
                   noreg1 = UBound(MatFactRiesgo, 1)
                   matx = ExtraerSMatFR(noreg1, noreg1, matfr, True, SiFactorRiesgo)
                   matx1 = ConvArVtDbl(ExtraeSubMatV(matx, 2, UBound(matx, 2), 1, UBound(matx, 1)))
                   matfechassh = ConvArVtDT(ExtraeSubMatV(matx, 1, 1, 2, UBound(matx, 1)))
                   Call GenRends3(matx1, 1, matfechassh, matrends, matb)
                   noesc = UBound(matrends, 1)
                   MatFactoresR = CargaFR1Dia(fcurvavalmax1, exito1)
          'Se carga la estructura de tasas para ese día de la matriz vector tasas
                   ReDim matval(1 To noesc, 1 To nofechasval) As Variant
                   ReDim matvalact(1 To noesc, 1 To nofechasval) As Variant
                   ReDim matvalpas(1 To noesc, 1 To nofechasval) As Variant
                   For i = 1 To nofechasval
                       For j = 1 To noesc
                           MatFactR1 = GenEscHist2(MatFactoresR, matrends, matb, j)
                           mprecio = CalcValuacion(matfechas1(i), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
                           suma = 0
                           sumaact = 0
                           sumapas = 0
                           For ll = 1 To UBound(mprecio, 1)
                               suma = suma + mprecio(ll).mtm_sucio
                                sumaact = sumaact + mprecio(ll).ps_activa
                               sumapas = sumapas + mprecio(ll).ps_pasiva
                           Next ll
                           matval(j, i) = suma
                           matvalact(j, i) = sumaact
                           matvalpas(j, i) = sumapas
                           If i = 1 And j = 1 Then
                              valmax = suma
                              valactmax = sumaact
                              valpasmax = sumapas
                              f_val_max = matfechas1(i)
                              fescvalmax = matfechassh(j, 1)
                           Else
                              If opc_calc = 0 Then
                                 If valmax < suma Then
                                    valmax = suma
                                    valactmax = sumaact
                                    valpasmax = sumapas
                                    f_val_max = matfechas1(i)
                                    fescvalmax = matfechassh(j, 1)
                                 End If
                              ElseIf opc_calc = 1 Then
                                 If valmax > suma Then
                                    valmax = suma
                                    valactmax = sumaact
                                    valpasmax = sumapas
                                    f_val_max = matfechas1(i)
                                    fescvalmax = matfechassh(j, 1)
                                 End If
                              End If
                           End If
                       Next j
                       AvanceProc = i / nofechasval
                       MensajeProc = "Swap " & coperacion & " grupo " & idgrupoc & " Paso 2 " & Format$(AvanceProc, "##0.00 %")
                       DoEvents
                   Next i
                   txtcadena1 = ""
                   txtcadena2 = ""
                   txtcadena3 = ""
                   txtcadena4 = ""
                   txtcadena5 = ""
                   For i = 1 To nofechasval
                       txtcadena1 = txtcadena1 & matfechas1(i) & ","
                   Next i
                   For i = 1 To noesc
                       txtcadena2 = txtcadena2 & matfechassh(i, 1) & ","
                   Next i
                   For i = 1 To nofechasval
                       For j = 1 To noesc
                           txtcadena3 = txtcadena3 & matval(j, i) & ","
                           txtcadena4 = txtcadena4 & matvalact(j, i) & ","
                           txtcadena5 = txtcadena5 & matvalpas(j, i) & ","
                       Next j
                   Next i
                   RegResLimC2.AddNew
                   RegResLimC2.Fields(0) = dtfecha
                   RegResLimC2.Fields(1) = coperacion
                   RegResLimC2.Fields(2) = idgrupoc
                   RegResLimC2.Fields(3) = opc_calc
                   RegResLimC2.Fields(4) = dtfecha1
                   RegResLimC2.Fields(5) = dtfecha2
                   RegResLimC2.Fields(6) = fescvalmax
                   RegResLimC2.Fields(7) = f_val_max
                   RegResLimC2.Fields(8) = valmax
                   RegResLimC2.Fields(9) = valactmax
                   RegResLimC2.Fields(10) = valpasmax
                   Call GuardarElementoClob(txtcadena1, RegResLimC2, "H_FECHAS1")
                   Call GuardarElementoClob(txtcadena2, RegResLimC2, "H_FECHAS2")
                   Call GuardarElementoClob(txtcadena3, RegResLimC2, "H_VALMAX")
                   Call GuardarElementoClob(txtcadena4, RegResLimC2, "H_VALMAXACT")
                   Call GuardarElementoClob(txtcadena5, RegResLimC2, "H_VALMAXPAS")
                   RegResLimC2.Update
                   final = True
                   bl_exito = True
                   SiIncTasaCVig = True
                   txtmsg = "Proceso finalizado correctamente"
                Else
                   txtmsg = txtmsg2
                   final = False
                   bl_exito = False
                End If
             Else
                txtmsg = "No ha posicion"
                final = True
                bl_exito = False
             End If
          Else
             txtmsg = "No se puede ejecutar este proceso"
             final = False
             bl_exito = False
          End If
       Else
          final = False
          bl_exito = False
       End If
    Else
      txtmsg = "No se encontraron datos de la operacion " & coperacion
      final = False
      bl_exito = False
    End If
End Sub

Sub DetCurvaValMaxOper(ByVal dtfecha As Date, ByVal coperacion As String, ByVal opcionc As Integer, ByVal id_procp As Integer, ByVal opcion As Integer, ByRef txtmsg As String, ByRef final As Boolean, ByRef bl_exito As Boolean)
If ActivarControlErrores Then
 On Error GoTo hayerror
End If
Dim i As Integer, j As Integer
Dim l As Integer
Dim noreg As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim valmax As Double
Dim fmax As Date
Dim ffutmax As Date
Dim txtborra As String
Dim txtinserta As String
Dim rmesa As New ADODB.recordset

bl_exito = False
'se determina si ya corrieron todos los subprocesos requeridos para realizar este calculo
txtfecha1 = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro1 = "SELECT COUNT(*) FROM " & DetermTablaSubproc(opcion) & " WHERE FECHAP = " & txtfecha1 & " AND ID_SUBPROCESO = " & id_procp & " AND PARAMETRO6 = '" & coperacion & "'  AND EXITO <> 'S'"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg = 0 Then
    txtfiltro2 = "SELECT * FROM " & TablaLimContrap1 & " WHERE FECHA = " & txtfecha1 & " AND COPERACION = '" & coperacion & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       ReDim mata(1 To noreg, 1 To 3)
       For i = 1 To noreg
           mata(i, 1) = rmesa.Fields("FESCMAX")  'fecha escenarios
           mata(i, 2) = rmesa.Fields("FVALMAX")  'fecha valuacion
           mata(i, 3) = rmesa.Fields("VALMAX")   'valuacion
           rmesa.MoveNext
           AvanceProc = i / noreg
           MensajeProc = "Calculando curva val max " & Format(AvanceProc, "##0.00 %")
       Next i
       rmesa.Close
       valmax = 0
       fmax = 0
       ffutmax = 0
       For i = 1 To noreg
           If i = 1 Then
              fmax = mata(i, 1)
              ffutmax = mata(i, 2)
              valmax = mata(i, 3)
           Else
              If opcionc = 0 Then
                 If mata(i, 3) > valmax Then
                    fmax = mata(i, 1)
                    ffutmax = mata(i, 2)
                    valmax = mata(i, 3)
                 End If
              ElseIf opcionc = 1 Then
                 If mata(i, 3) < valmax Then
                    fmax = mata(i, 1)
                    ffutmax = mata(i, 2)
                    valmax = mata(i, 3)
                 End If
              End If
           End If
       Next i
       txtfecha2 = "TO_DATE('" & Format$(fmax, "DD/MM/YYYY") & "','DD/MM/YYYY')"
       txtfecha3 = "TO_DATE('" & Format$(ffutmax, "DD/MM/YYYY") & "','DD/MM/YYYY')"
       txtborra = "DELETE FROM " & TablaResLimContrap & " WHERE FECHA = " & txtfecha1 & " AND coperacion = '" & coperacion & "'"
       ConAdo.Execute txtborra
       txtinserta = "INSERT INTO " & TablaResLimContrap & " VALUES("
       txtinserta = txtinserta & txtfecha1 & ","
       txtinserta = txtinserta & "'" & coperacion & "',"
       txtinserta = txtinserta & valmax & ","
       txtinserta = txtinserta & txtfecha2 & ","
       txtinserta = txtinserta & txtfecha3 & ","
       txtinserta = txtinserta & "null,"
       txtinserta = txtinserta & "null,"
       txtinserta = txtinserta & "null)"
       ConAdo.Execute txtinserta
       txtmsg = "El proceso finalizo correctamente"
       final = True
       bl_exito = True
    Else
       txtmsg = "No hay datos en la base"
       final = True
       bl_exito = False
    End If
Else
    txtmsg = "No se han terminado las operaciones previas"
    final = False
    bl_exito = False
End If
Exit Sub
hayerror:
MsgBox error(Err())
End Sub

Sub DeterminaEscValMax(ByVal dtfecha As Date, ByVal coperacion As String, ByVal opcionc As Integer, ByVal id_procp As Integer, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef final As Boolean, ByRef bl_exito As Boolean)
If ActivarControlErrores Then
 On Error GoTo hayerror
End If
Dim i As Integer, j As Integer
Dim noreg As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtact As String
Dim valmax As Double
Dim fmax As Date
Dim ffutmax As Date
Dim txttabla As String
Dim rmesa As New ADODB.recordset

txttabla = DetermTablaSubproc(id_tabla)
bl_exito = False
txtfecha1 = "TO_DATE('" & Format$(dtfecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro1 = "SELECT COUNT(*) FROM " & txttabla & " WHERE FECHAP = " & txtfecha1 & " AND ID_SUBPROCESO = " & id_procp & " AND PARAMETRO6 = '" & coperacion & "' AND EXITO <> 'S'"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg = 0 Then
    txtfiltro2 = "SELECT * FROM " & TablaLimContrap2 & " WHERE FECHA = " & txtfecha1 & " AND COPERACION = '" & coperacion & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       ReDim mata(1 To noreg, 1 To 3)
       For i = 1 To noreg
           mata(i, 1) = rmesa.Fields("FESCMAX")
           mata(i, 2) = rmesa.Fields("FVALMAX2")
           mata(i, 3) = rmesa.Fields("VALMAX2")
           rmesa.MoveNext
       Next i
       rmesa.Close
       For i = 1 To noreg
           If i = 1 Then
              fmax = mata(i, 1)
              ffutmax = mata(i, 2)
              valmax = mata(i, 3)
           Else
              If opcionc = 0 Then
                 If mata(i, 3) > valmax Then
                    fmax = mata(i, 1)
                    ffutmax = mata(i, 2)
                    valmax = mata(i, 3)
                 End If
              ElseIf opcionc = 1 Then
                 If mata(i, 3) < valmax Then
                    fmax = mata(i, 1)
                    ffutmax = mata(i, 2)
                    valmax = mata(i, 3)
                 End If
              End If
           End If
       Next i
       txtfecha2 = "TO_DATE('" & Format$(fmax, "DD/MM/YYYY") & "','DD/MM/YYYY')"
       txtfecha3 = "TO_DATE('" & Format$(ffutmax, "DD/MM/YYYY") & "','DD/MM/YYYY')"
       txtact = "UPDATE " & TablaResLimContrap & " SET "
       txtact = txtact & " FESCMAX2 = " & txtfecha2 & ","
       txtact = txtact & " FVALMAX2 = " & txtfecha3 & ","
       txtact = txtact & " VALMAX2 = " & valmax
       txtact = txtact & " WHERE FECHA = " & txtfecha1
       txtact = txtact & " AND COPERACION = '" & coperacion & "'"
       ConAdo.Execute txtact
       final = True
       bl_exito = True
       txtmsg = "El proceso finalizo correctamente"
    Else
       final = True
       bl_exito = True
    End If
Else
    txtmsg = "No se han terminado todos los calculos"
    final = False
    bl_exito = False
End If
Exit Sub
hayerror:
MsgBox error(Err())
End Sub

Sub CalcExpMaxFwd(ByVal fecha As Date, ByVal tipopos As Integer, ByVal fechareg As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal nconf As Double, ByRef txtmsg As String, ByRef exito As Boolean)
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim txtcurva As String
Dim txttc As String
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim strike As Double
Dim montonoc As Double
Dim sigma As Double
Dim dxv As Long
Dim tasa As Double
Dim contar As Long
Dim htiempo As Long
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim indice As Long
Dim indice1 As Long
Dim indice2 As Long
Dim nofr As Long
Dim matplz() As Long
Dim matplz1() As Long
Dim matc() As Variant
Dim mattc() As Variant
Dim tc As Double
Dim tcf As Double
Dim exposicion As Double
Dim txtborra As String
Dim txtcadena As String
Dim txtfecha As String
Dim txtmsg0 As String
Dim txtmsg2 As String

    mattxt = CrearFiltroPosOperPort(tipopos, fechareg, txtnompos, horareg, cposicion, coperacion)
    Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
    If Not EsArrayVacio(matpos) Then
       Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
       If exito2 Then
          Call VerifCargaFR3(#1/1/2008#, fecha)
          noreg = UBound(MatFactRiesgo, 1)
          indice1 = BuscarValorArray(#2/1/2008#, MatFactRiesgo, 1)
          indice2 = BuscarValorArray(fecha, MatFactRiesgo, 1)
          nofr = indice2 - indice1 + 1
          ReDim matfr(1 To nofr, 1 To 1) As Double
          If matposfwd(1).ClaveProdFwd = "FWD MXN/USD" Then
            txtcurva = "TC FWD FIX"
             txttc = "DOLAR PIP FIX"
          Else
             txtcurva = "TC FWD EURO"
             txttc = "EURO BM D"
          End If
          contar = 0
          For i = 1 To NoFactores
              If MatCaracFRiesgo(i).nomFactor = txtcurva Then
                 contar = contar + 1
                 ReDim Preserve matfr(1 To nofr, 1 To contar) As Double
                 ReDim Preserve matplz(1 To contar) As Long
                 ReDim Preserve matplz1(1 To contar) As Long
                 matplz(contar) = MatCaracFRiesgo(i).plazo
                 For j = 1 To nofr
                     matfr(j, contar) = MatFactRiesgo(indice1 - 1 + j, i)
                 Next j
               End If
          Next i
          ReDim matcov(1 To contar) As New propCurva
          For i = 1 To contar
              If matplz(i) <> 1 Then
                 htiempo = Int(matplz(i) * 2 / 3)
                 matplz1(i) = Int(matplz(i) * 2 / 3)
              Else
                 htiempo = 1
                 matplz1(i) = 1
              End If
              ReDim matrends(1 To nofr - htiempo, 1 To 1) As Double
              For j = 1 To nofr - htiempo
                  matrends(j, 1) = CalcRend2(matfr(j, i), matfr(j + htiempo, i), MatCaracFRiesgo(i).tfactor)
              Next j
              matcov(i).valor = Sqr(CVarianzap(matrends, 1, "c"))
              matcov(i).plazo = matplz(i)
          Next i
          montonoc = matposfwd(1).MontoNocFwd
          dxv = matposfwd(1).FVencFwd - fecha
          strike = matposfwd(1).PAsignadoFwd
          sigma = CalculaTasa(matcov, dxv, 1)
          matc = LeerCurvaCompleta(fecha, exito2)
          For i = 1 To UBound(matc, 2)
              If matc(1, i) = 21 Then
                tasa = matc(dxv + 1, i)
                Exit For
              End If
          Next i
          mattc = Leer1FactorR(fecha, fecha, txttc, 0)
          indice = BuscarValorArray(fecha, mattc, 1)
          tc = mattc(indice, 2)
          tcf = tc * Exp(sigma * NormalInv(1 - nconf))
          exposicion = montonoc * (strike - tcf) / (1 + tasa * dxv / 360)
          txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtborra = "DELETE FROM " & TablaExpFwds & " WHERE FECHA = " & txtfecha
          txtborra = txtborra & " AND CPOSICION = " & cposicion
          txtborra = txtborra & " AND COPERACION = '" & coperacion & "'"
          ConAdo.Execute txtborra
          txtcadena = "INSERT INTO " & TablaExpFwds & " VALUES("
          txtcadena = txtcadena & txtfecha & ","
          txtcadena = txtcadena & cposicion & ","
          txtcadena = txtcadena & txtfecha & ","
          txtcadena = txtcadena & "'" & coperacion & "',"
          txtcadena = txtcadena & "'" & matposfwd(1).ClaveProdFwd & "',"
          txtcadena = txtcadena & sigma & ","
          txtcadena = txtcadena & exposicion & ")"
          ConAdo.Execute txtcadena
          txtmsg = "Proceso finalizado correctamente"
          exito = True
       Else
          txtmsg = txtmsg2
          exito = False
       End If
    Else
       exito = False
    End If
End Sub
