Attribute VB_Name = "modMarcoOp"
Option Explicit

Sub Carga_VAR_MO(ByVal fecha1 As Date, ByVal fecha2 As Date)

Dim TablaDetalleMo As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfiltro3 As String
Dim txtfecha As String
Dim txtejecuta As String
Dim txtinsert As String
Dim Sector As String
Dim moneda As String
Dim row As Integer
Dim i As Long
Dim j As Long
Dim k As Long
Dim indice As Long
Dim indice2 As Long
Dim c_emision As String
Dim fechas() As Date
Dim posicion() As Variant
Dim val_posicion() As Variant
Dim cat2() As Variant
Dim cat4() As Variant
Dim cat5() As Variant
Dim cat6() As Variant
Dim cat7() As Variant
Dim tipoContraparte As String
Dim tipoEmision As String
Dim calif_sp As String
Dim calif_moodys As String
Dim calif_fitch As String
Dim calif_hr As String
Dim calif() As Variant
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim monto_cir As Double
Dim tasa_ref As String
Dim v_nominal As Long
Dim f_venc As Date
Dim tipo_md As String
Dim rmesa As New ADODB.recordset

fechas = LeerFechasVaR(fecha1, fecha2)

'Carga catálogos en memoria
'cat4 = CargaCatalogo(35)
'cat5 = CargaCatalogo(38)
'cat6 = CargaCatalogo(40)
'cat7 = CargaCatalogo(44)

'Carga posicion de mesa en memoria
For i = 1 To UBound(fechas)
    txtfiltro1 = "SELECT * FROM " & TablaPosMD
    txtfiltro1 = txtfiltro1 & " WHERE FECHAREG = TO_DATE ( '" & fechas(i, 1) & "','dd/mm/yyyy' )"
    txtfiltro1 = txtfiltro1 & " AND TIPOPOS=1"
    txtfiltro1 = txtfiltro1 & " AND ( TOPERACION = '1' OR TOPERACION = '4' )"
    
    txtfiltro2 = "SELECT * FROM " & TablaValPos & " "
    txtfiltro2 = txtfiltro2 & "WHERE FECHAP = TO_DATE('" & fechas(i, 1) & "','dd/mm/yyyy') "
    txtfiltro2 = txtfiltro2 & "AND ID_VALUACION = 1 "
    txtfiltro2 = txtfiltro2 & "AND ESC_FR='Normal' "
    txtfiltro2 = txtfiltro2 & "AND (CPOSICION = 1 OR CPOSICION = 2 OR CPOSICION= 8 OR CPOSICION= 9)"
    
    txtfiltro = "SELECT TV, EMISION, SERIE, t1.CPOSICION, t1.TOPERACION, "
    txtfiltro = txtfiltro & "t1.COPERACION, t1.NO_TITULOS, P_SUCIO, VAL_PIP_S, DUR_ACT "
    txtfiltro = txtfiltro & "FROM (" & txtfiltro1 & ") t1 "
    txtfiltro = txtfiltro & "INNER JOIN (" & txtfiltro2 & ") t2 "
    txtfiltro = txtfiltro & "ON t1.COPERACION = t2.COPERACION "
    txtfiltro = txtfiltro & "ORDER BY TV, EMISION, SERIE"
    txtfiltro3 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
    rmesa.Open txtfiltro3, ConAdo
    row = rmesa.Fields(0)
    rmesa.Close
    If row <> 0 Then
       rmesa.Open txtfiltro, ConAdo
    ReDim posicion(1 To row, 1 To 10) As Variant
       For j = 1 To row
           For k = 1 To 10
               posicion(j, k) = rmesa.Fields(k - 1).value
           Next k
           rmesa.MoveNext
       Next j
       rmesa.Close
    
       Call DepuraVentasFV(posicion)
       Call DepuraEmision(posicion, MatMOGub)
       txtfecha = "TO_DATE('" & fechas(i, 1) & "','dd/mm/yyyy')"
       txtejecuta = "DELETE " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
       'ConAdo.Execute txtejecuta
       matvp = LeerVPrecios(fechas(i, 1), mindvp)
       For j = 1 To row
        'Control tabla VAR_TD_VEC_PRECIOS
           c_emision = GeneraClaveEmision(posicion(j, 1), posicion(j, 2), posicion(j, 3))
           indice = BuscarValorArray(c_emision, matvp, 22)
           If indice = 0 Then
              MsgBox "No se encuentra el matvp " & posicion(j, 1) & "_" & _
                     posicion(j, 2) & "_" & posicion(j, 3) & " en " & TablaVecPrecios & " " & _
                    " para la fecha " & fechas(i, 1)
            GoTo SiguienteInstr
           End If
           monto_cir = matvp(indice, 12)
           tasa_ref = matvp(indice, 21)
           v_nominal = matvp(indice, 21)
           f_venc = matvp(indice, 12)
           calif_sp = matvp(indice, 18)
           calif_fitch = matvp(indice, 19)
           calif_moodys = matvp(indice, 20)
           calif_hr = matvp(indice, 21)
           'Asigna calificación
           calif = AsignaCalif(calif_sp, calif_moodys, calif_fitch, calif_hr)
           indice2 = BuscarValorArray(posicion(j, 1) & "_" & posicion(j, 2), MatMOSectorMD, 4)
           If indice2 <> 0 Then
              tipo_md = MatMOSectorMD(indice2, 4)
              Sector = MatMOSectorMD(indice2, 5)
           Else
              tipo_md = ""
              Sector = ""
           End If
        
        'Control catálogo 1
     
        
        'Asigna tasa fija o variable
           If (matvp(indice, 23) = "Tasa Fija" Or matvp(indice, 23) = "NA") Then
              tasa_ref = "TF"
           Else
              tasa_ref = "TV"
           End If
        
        'Asigna tipo contraparte
           indice2 = BuscarValorArray(posicion(j, 1), MatMOContrap, 1)
           If indice2 = 0 Then
              tipoContraparte = "OTROS"
           Else
              tipoContraparte = MatMOContrap(indice2, 2)
           End If
        
        'Asigna tipo emisión
           indice2 = BuscarValorArray(c_emision, MatMOEmPriv, 1)
           If indice2 = 0 Then
              tipoEmision = "Pública"
           Else
              tipoEmision = "Privada"
           End If
        'Asigna moneda
           indice2 = BuscarValorArray(matvp(indice, 14), MatMOMon, 1)
           If indice2 <> 0 Then
              moneda = MatMOMon(indice2, 2)
           Else
              moneda = ""
           End If
        
        '************************************
           txtinsert = "Insert into " & TablaDetalleMo & " values ("
           txtinsert = txtinsert & "to_date('" & fechas(i, 1) & "','dd/mm/yyyy'), "
           txtinsert = txtinsert & "'" & posicion(j, 1) & "', "                        'tv
           txtinsert = txtinsert & "'" & posicion(j, 2) & "', "                        'emision
           txtinsert = txtinsert & "'" & posicion(j, 3) & "', "                        'serie
           txtinsert = txtinsert & posicion(j, 7) & ", "                               'no titulos
           txtinsert = txtinsert & "'" & posicion(j, 4) & "', "                        'posicion
           txtinsert = txtinsert & "'" & calif_moodys & "', "                          'calif moodys
           txtinsert = txtinsert & "'" & calif_sp & "', "                              'calif sp
           txtinsert = txtinsert & "'" & calif_fitch & "', "                           'calif fitch
           txtinsert = txtinsert & "'" & calif_hr & "', "                                   'calif hr
           txtinsert = txtinsert & "'" & Sector & "', "      'Sector
           txtinsert = txtinsert & "'" & tipo_md & "', "      'Tipo
           txtinsert = txtinsert & (f_venc - fechas(i, 1)) / 365 & ", "     'Plazo en años
           txtinsert = txtinsert & "'" & tasa_ref & "', "                     'Tasa
           txtinsert = txtinsert & v_nominal & ", "                                 'Valor nominal
           txtinsert = txtinsert & posicion(j, 7) & ", "                               'PrecioS_SIVARMER
           txtinsert = txtinsert & posicion(j, 8) & ", "                               'PrecioS_PIP
           txtinsert = txtinsert & posicion(j, 9) & ", "                               'Duración
           txtinsert = txtinsert & monto_cir & ", "                                 'monto_circulación
           txtinsert = txtinsert & "'" & calif(0, 0) & "', "                           'calif_min
           txtinsert = txtinsert & "'" & calif(0, 1) & "', "                           'escala
           txtinsert = txtinsert & "'" & tipoContraparte & "', "                       'Tipo Contraparte
           txtinsert = txtinsert & "'" & tipoEmision & "', "                           'Pública/Privada
           txtinsert = txtinsert & "'" & moneda & "') "                                'Moneda
           'ConAdo.Execute txtinsert

SiguienteInstr:
    Next j
End If
SiguienteFecha:
Next i

MsgBox "Carga " & TablaDetalleMo & " de los días seleccionados terminada correctamente"

End Sub

Function AsignaCalif2(ByVal calif_sp As String, ByVal calif_mdys As String, ByVal calif_fitch As String, ByVal calif_hr As String, ByRef mat() As Variant) As Variant

Dim min As Integer
Dim num2 As Integer
Dim num3 As Integer
Dim num4 As Integer
Dim matres(0, 1) As Variant

min = BuscarIndice(calif_sp, mat, 1)
num2 = BuscarIndice(calif_mdys, mat, 2)
If min > num2 Then min = min Else min = num2
num3 = BuscarIndice(calif_fitch, mat, 3)
If min > num3 Then min = min Else min = num3
num4 = BuscarIndice(calif_hr, mat, 4)
If min > num4 Then min = min Else min = num4

If min <> -1 Then
   matres(0, 0) = mat(min, 7)
   matres(0, 1) = mat(min, 5)
End If

AsignaCalif2 = matres
End Function

Function BuscarIndice(ByVal valor As Variant, ByRef mat() As Variant, col As Integer) As Long

Dim i As Integer, n As Integer

i = 1
n = UBound(mat, 1)
While valor <> mat(i, col)
    If i < n Then
        i = i + 1
    Else
        i = -1
        GoTo final
    End If
Wend

final: BuscarIndice = i

End Function

Sub DepuraVentasFV(ByRef mat1() As Variant)

Dim i As Long

For i = 1 To UBound(mat1(), 1)
    If mat1(i, 5) = "4" Then
        mat1(i, 7) = mat1(i, 7) * (-1)
    End If
Next i

End Sub

Sub DepuraEmision(mat1() As Variant, mat2() As Variant)
Dim i As Long
Dim emision As String
Dim indice As Long
For i = 1 To UBound(mat1())
    If mat1(i, 2) = "GOBFED" Or Left(mat1(i, 2), 3) = "BPA" Then
        indice = BuscarValorArray(mat1(i, 1), mat2, 1)
        If indice = 0 Then
           MsgBox "El Tipo Valor" & mat1(i, 1) & " no se encuentra en el catálogo 2"
           emision = ""
        Else
           emision = mat2(indice, 2)
        End If
        mat1(i, 2) = emision
    End If

Next i

End Sub

Function ObtenerLimMarcoOp(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim nogrupos As Integer
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim rmesa As New ADODB.recordset

nogrupos = 18
txtfecha1 = "TO_DATE('" & Format$(fecha1, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfecha2 = "TO_DATE('" & Format$(fecha2, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaResMO & " WHERE FECHA > " & txtfecha1 & " AND FECHA <= " & txtfecha2
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matres(1 To nogrupos, 1 To 6) As Variant
   For i = 1 To nogrupos
       matres(i, 1) = i
   Next i
   matres(1, 2) = "Posicion en tasa fija"
   matres(2, 2) = "Posicion en tasa fija>1 año"
   matres(3, 2) = "Cetes"
   matres(4, 2) = "Bonos M y S"
   matres(5, 2) = "Otros"
   matres(6, 2) = "PRLV BD"
   matres(7, 2) = "CEBURES BD"
   matres(8, 2) = "PRLV BP"
   matres(9, 2) = "CEBURES BP"
   matres(10, 2) = "Plazo> 1 año"
   matres(11, 2) = "Plazo < 1 año"
   matres(12, 2) = "Posición en tasa variable"
   matres(13, 2) = "Gubernamental"
   matres(14, 2) = "Banca de desarrollo"
   matres(15, 2) = "Banca privada"
   matres(16, 2) = "Entidades paraestatales"
   matres(17, 2) = "Posicion en directo"
   matres(18, 2) = "Entidades paraestatales"
  
   For i = 1 To noreg
       For j = 1 To nogrupos
           If matres(j, 1) = rmesa.Fields("ID_REPORTE") Then
              If matres(j, 3) <> 0 Then
                 matres(j, 3) = Maximo(rmesa.Fields("CONSUMO_LIM"), matres(j, 3))
              Else
                 matres(j, 3) = rmesa.Fields("CONSUMO_LIM")
              End If
              If matres(j, 4) <> 0 Then
                 matres(j, 4) = Minimo(rmesa.Fields("CONSUMO_LIM"), matres(j, 4))
              Else
                 matres(j, 4) = rmesa.Fields("CONSUMO_LIM")
              End If
              matres(j, 5) = matres(j, 5) + rmesa.Fields("CONSUMO_LIM")
              matres(j, 6) = matres(j, 6) + 1
           End If
       Next j
       rmesa.MoveNext
   Next i
   For i = 1 To nogrupos
       If matres(i, 6) <> 0 Then matres(i, 5) = matres(i, 5) / matres(i, 6)
   Next i
   rmesa.Close
End If
ObtenerLimMarcoOp = matres

End Function

Function DeterminaPorcCalif(ByVal fecha As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim txtfecha As String
Dim valor As Double
Dim mata(1 To 9, 1 To 2) As Variant
Dim rmesa As New ADODB.recordset

mata(1, 1) = "AAA"
mata(2, 1) = "AA+"
mata(3, 1) = "AA"
mata(4, 1) = "AA-"
mata(5, 1) = "A+"
mata(6, 1) = "A"
mata(7, 1) = "BBB+"
mata(8, 1) = "GR 2"
mata(9, 1) = "BBB-"

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT SUM(N_TITULOS*PSUCIO_SIVARMER) FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND (MESA =" & ClavePosPICV & " OR MESA =" & ClavePosPIDV & ")"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"

rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   valor = rmesa.Fields(0)
   rmesa.Close
End If
txtfiltro2 = "SELECT SUM(N_TITULOS*PSUCIO_SIVARMER) FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND (MESA =" & ClavePosPICV & " OR MESA =" & ClavePosPIDV & ")"
txtfiltro2 = txtfiltro2 & " AND (CALIF_MIN ='AAA' OR CALIF_MIN ='GR 1' OR TIPO_CONTRAPARTE = 'GUBER')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(1, 2) = rmesa.Fields(0)
   rmesa.Close
End If

txtfiltro2 = "SELECT SUM(N_TITULOS*PSUCIO_SIVARMER) FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND (MESA = " & ClavePosPICV & " OR MESA = " & ClavePosPIDV & ")"
txtfiltro2 = txtfiltro2 & " AND (CALIF_MIN = 'AA+')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(2, 2) = rmesa.Fields(0)
   rmesa.Close
End If

txtfiltro2 = "SELECT SUM(N_TITULOS*PSUCIO_SIVARMER) FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND (MESA = " & ClavePosPICV & " OR MESA = " & ClavePosPIDV & ")"
txtfiltro2 = txtfiltro2 & " AND (CALIF_MIN = 'AA')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(3, 2) = ReemplazaVacioValor(rmesa.Fields(0), 0)
   rmesa.Close
End If

txtfiltro2 = "SELECT SUM(N_TITULOS*PSUCIO_SIVARMER) FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND (MESA = " & ClavePosPICV & " OR MESA = " & ClavePosPIDV & ")"
txtfiltro2 = txtfiltro2 & " AND (CALIF_MIN = 'AA-')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(4, 2) = rmesa.Fields(0)
   rmesa.Close
End If
txtfiltro2 = "SELECT SUM(N_TITULOS*PSUCIO_SIVARMER) FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND (MESA = " & ClavePosPICV & " OR MESA = " & ClavePosPIDV & ")"
txtfiltro2 = txtfiltro2 & " AND (CALIF_MIN = 'A+')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(5, 2) = rmesa.Fields(0)
   rmesa.Close
End If

txtfiltro2 = "SELECT SUM(N_TITULOS*PSUCIO_SIVARMER) FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND (MESA = " & ClavePosPICV & " OR MESA = " & ClavePosPIDV & ")"
txtfiltro2 = txtfiltro2 & " AND (CALIF_MIN = 'A')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(6, 2) = rmesa.Fields(0)
   rmesa.Close
End If

txtfiltro2 = "SELECT SUM(N_TITULOS*PSUCIO_SIVARMER) FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND (MESA = " & ClavePosPICV & " OR MESA = " & ClavePosPIDV & ")"
txtfiltro2 = txtfiltro2 & " AND CALIF_MIN = 'BBB+' AND ESCALA_CALIF_MIN = 'Global' AND TIPO_CONTRAPARTE <> 'GUBER'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(7, 2) = rmesa.Fields(0)
   rmesa.Close
End If

txtfiltro2 = "SELECT SUM(N_TITULOS*PSUCIO_SIVARMER) FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND (MESA = " & ClavePosPICV & " OR MESA = " & ClavePosPIDV & ")"
txtfiltro2 = txtfiltro2 & " AND (CALIF_MIN = 'GR 2')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(8, 2) = rmesa.Fields(0)
   rmesa.Close
End If

txtfiltro2 = "SELECT SUM(N_TITULOS*PSUCIO_SIVARMER) FROM " & TablaDetalleMo & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND (MESA = " & ClavePosPICV & " OR MESA = " & ClavePosPIDV & ")"
txtfiltro2 = txtfiltro2 & " AND (CALIF_MIN = 'BBB-')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(9, 2) = rmesa.Fields(0)
   rmesa.Close
End If


For i = 1 To 9
   mata(i, 2) = mata(i, 2) / valor
Next i

DeterminaPorcCalif = mata

End Function

Function DetermPlazoPromMO(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim nosector As Integer
Dim noreg As Long
Dim i As Integer
Dim indice As Long
Dim contar As Long
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim rmesa As New ADODB.recordset

nosector = 5
ReDim matres(1 To nosector, 1 To 4) As Variant

matres(1, 1) = "G"
matres(2, 1) = "BD"
matres(3, 1) = "BP"
matres(4, 1) = "PE"
matres(5, 1) = "Otros"
txtfecha1 = "TO_DATE('" & Format$(fecha1, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfecha2 = "TO_DATE('" & Format$(fecha2, "DD/MM/YYYY") & "','DD/MM/YYYY')"

For i = 1 To nosector
    txtfiltro2 = "SELECT max(PLAZO),min(PLAZO),SUM(N_TITULOS*PSUCIO_SIVARMER),sum(plazo*n_titulos*psucio_sivarmer)"
    txtfiltro2 = txtfiltro2 & " FROM " & TablaDetalleMo & " WHERE (FECHA > " & txtfecha1 & " AND FECHA <= " & txtfecha2 & ")"
    txtfiltro2 = txtfiltro2 & " AND SECTOR = '" & matres(i, 1) & "' AND MESA = 1"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       matres(i, 2) = rmesa.Fields(0)
       matres(i, 3) = rmesa.Fields(1)
       If rmesa.Fields(2) <> 0 Then
       matres(i, 4) = rmesa.Fields(3) / rmesa.Fields(2)
       Else
       matres(i, 4) = 0
       End If
       rmesa.Close
    End If
Next i
DetermPlazoPromMO = matres
End Function

Function AsignaCalif(ByVal calif_sp As String, ByVal calif_mdys As String, ByVal calif_fitch As String, ByVal calif_hr As String)
Dim indice As Integer
Dim indice1 As Integer
Dim indice2 As Integer
Dim indice3 As Integer
Dim indice4 As Integer
Dim escala As String
Dim escala1 As String
Dim escala2 As String
Dim escala3 As String
Dim escala4 As String
Dim plazo As String
Dim plazo1 As String
Dim plazo2 As String
Dim plazo3 As String
Dim plazo4 As String
Dim i As Integer


For i = 1 To UBound(MatMOCalif)
    If calif_sp = MatMOCalif(i, 1) Then
       indice1 = MatMOCalif(i, 5)
       escala1 = MatMOCalif(i, 6)
       plazo1 = MatMOCalif(i, 7)
       Exit For
    End If
Next i
For i = 1 To UBound(MatMOCalif)
    If calif_mdys = MatMOCalif(i, 2) Then
       indice2 = MatMOCalif(i, 5)
       escala2 = MatMOCalif(i, 6)
       plazo2 = MatMOCalif(i, 7)
       Exit For
    End If
Next i
For i = 1 To UBound(MatMOCalif)
    If calif_fitch = MatMOCalif(i, 3) Then
       indice3 = MatMOCalif(i, 5)
       escala3 = MatMOCalif(i, 6)
       plazo3 = MatMOCalif(i, 7)
       Exit For
    End If
Next i
For i = 1 To UBound(MatMOCalif)
    If calif_hr = MatMOCalif(i, 4) Then
       indice4 = MatMOCalif(i, 5)
       escala4 = MatMOCalif(i, 6)
       plazo4 = MatMOCalif(i, 7)
       Exit For
    End If
Next i
indice = indice1
If indice <> 0 Then
escala = escala1
plazo = plazo1
End If
If indice2 > indice Then
   indice = indice2
   escala = escala2
   plazo = plazo2
End If

If indice3 > indice Then
   indice = indice3
   escala = escala3
   plazo = plazo3
End If
If indice4 > indice Then
   indice = indice4
   escala = escala4
   plazo = plazo4
End If

If indice <> 0 Then
For i = 1 To UBound(MatMOCalif, 1)
   If indice = MatMOCalif(i, 5) And escala = MatMOCalif(i, 6) And plazo = MatMOCalif(i, 7) Then
      AsignaCalif = MatMOCalif(i, 8)
      Exit Function
   End If
Next i
Else
   AsignaCalif = ""
End If

End Function


Function DetermMOSector(ByVal tv As String)
Dim i As Integer

'For i = 1 To UBound(MatMOSectorMD, 1)
'    If tv = MatMOSector(i, 1) Then
'       DetermMOSector = MatMOSectorMD(i, 4)
'       Exit Function
'    End If
'Next i
DetermMOSector = ""
End Function

Function DetermMOTipo(ByVal clave As String)
Dim i As Integer
For i = 1 To UBound(MatMOSectorMD, 1)
    If clave = MatMOSectorMD(i, 1) Then
       DetermMOTipo = MatMOSectorMD(i, 5)
       Exit Function
    End If
Next i
DetermMOTipo = ""
End Function

