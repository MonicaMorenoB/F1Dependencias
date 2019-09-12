Attribute VB_Name = "modActSWRIESGOS3"
Option Explicit

Sub GuardaIkosDerivados(mat_data2, ByVal fecha As Date)
 Dim txtfecha As String
 Dim fecha1 As String
 Dim fecha2 As String
 Dim fecha4 As String
 Dim fecha5 As String
 Dim fecha6 As String
 Dim i As Integer
 Dim sql_Del As String
 Dim nombase_ As String
 Dim totregoral As Long
 Dim pas_tipo As String
 Dim sql_Mat As String
 
 txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 sql_Del = "delete sw_riesgos3_respaldo where fecha_pos= " & txtfecha
 ConAdo.Execute sql_Del
           
 Dim totregora1, TotCamOra1 As Integer
 
 nombase_ = "sw_riesgos3_respaldo"
 totregora1 = UBound(mat_data2, 1)
  
 For i = 1 To totregora1
    If Val(mat_data2(i, 27)) > 0 Then
       pas_tipo = "Fija"
    Else
       pas_tipo = "Variable"
    End If
    'funtion que convierte fechas normal a fechas oracle
    fecha1 = fecha_oracle(mat_data2(i, 1))
    fecha2 = fecha_oracle(mat_data2(i, 7))
    txtfecha = fecha_oracle(mat_data2(i, 8))
    fecha4 = fecha_oracle(mat_data2(i, 9))
    fecha5 = fecha_oracle(mat_data2(i, 12))
    fecha6 = fecha_oracle(mat_data2(i, 20))
    'inicia
    ' correccion de numsec
    sql_Mat = " insert into " & nombase_ & " values "
    sql_Mat = sql_Mat + "(" & fecha1
    sql_Mat = sql_Mat + "," & "'" & mat_data2(i, 2) & "'"
    sql_Mat = sql_Mat + "," & "'" & mat_data2(i, 3) & "'"
    sql_Mat = sql_Mat + "," & "'" & mat_data2(i, 4) & "'"
    sql_Mat = sql_Mat + "," & mat_data2(i, 5)
    sql_Mat = sql_Mat + "," & mat_data2(i, 6)
    sql_Mat = sql_Mat + "," & fecha2
    sql_Mat = sql_Mat + "," & txtfecha
    sql_Mat = sql_Mat + "," & fecha4
    sql_Mat = sql_Mat + "," & "'" & mat_data2(i, 10) & "'"
    sql_Mat = sql_Mat + "," & "'" & mat_data2(i, 11) & "'"
    sql_Mat = sql_Mat + "," & fecha5
    sql_Mat = sql_Mat + "," & mat_data2(i, 13)
    sql_Mat = sql_Mat + "," & mat_data2(i, 14)
    sql_Mat = sql_Mat + "," & mat_data2(i, 15)
    sql_Mat = sql_Mat + "," & mat_data2(i, 16)
    sql_Mat = sql_Mat + "," & mat_data2(i, 17)
    sql_Mat = sql_Mat + "," & "'" & mat_data2(i, 18) & "'"
    sql_Mat = sql_Mat + "," & "'" & mat_data2(i, 26) & "'"
    sql_Mat = sql_Mat + "," & mat_data2(i, 19)
    sql_Mat = sql_Mat + "," & fecha6
    sql_Mat = sql_Mat + "," & mat_data2(i, 21)
    sql_Mat = sql_Mat + "," & mat_data2(i, 22)
    sql_Mat = sql_Mat + "," & mat_data2(i, 23)
    sql_Mat = sql_Mat + "," & mat_data2(i, 24)
    sql_Mat = sql_Mat + "," & mat_data2(i, 25)
    sql_Mat = sql_Mat + "," & "'" & pas_tipo & "'"
    sql_Mat = sql_Mat + "," & "'" & mat_data2(i, 27) & "'"
    sql_Mat = sql_Mat + "," & mat_data2(i, 28)
    sql_Mat = sql_Mat + "," & mat_data2(i, 29)
    sql_Mat = sql_Mat + "," & mat_data2(i, 30)
    sql_Mat = sql_Mat + "," & mat_data2(i, 31)
    sql_Mat = sql_Mat + "," & "'" & mat_data2(i, 32) & "'"
    sql_Mat = sql_Mat + "," & mat_data2(i, 33)
    sql_Mat = sql_Mat + "," & mat_data2(i, 34)
    sql_Mat = sql_Mat + "," & mat_data2(i, 35)
    sql_Mat = sql_Mat + "," & mat_data2(i, 36) & ")"
    
  If mat_data2(i, 2) <> 0 Then
     ConAdo.Execute sql_Mat
  End If
  DoEvents
Next i
  
  
End Sub

Sub Lee_Ikos_Derivados(ByVal fecha As Date)
Dim txtfecha As String
Dim sql_deriv As String
Dim sql_num_deriv As String
Dim i As Long
Dim TotReg_deriv As Long
Dim Acum As Double
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
sql_num_deriv = "select count(*) from " & TablaInterfCarac & " where fecha_pos = " & txtfecha     ' where fecharegistro = " & txtfecha
rmesa.Open sql_num_deriv, ConAdo
TotReg_deriv = rmesa.Fields(0)
rmesa.Close
If TotReg_deriv <> 0 Then
   sql_deriv = "select * from " & TablaInterfCarac & " where fecha_pos = " & txtfecha
   rmesa.Open sql_deriv, ConAdo
   ReDim Mat_derivDinero(1 To TotReg_deriv, 1 To 36)
   rmesa.MoveFirst
   Acum = 0
   For i = 1 To TotReg_deriv
       Mat_derivDinero(i, 1) = ReemplazaVacioValor(rmesa.Fields("fecha_pos"), 0)
       Mat_derivDinero(i, 2) = ReemplazaVacioValor(rmesa.Fields("PERIODO"), 0)
       Mat_derivDinero(i, 3) = ReemplazaVacioValor(rmesa.Fields("CVE_INST"), 0)
       Mat_derivDinero(i, 4) = ReemplazaVacioValor(rmesa.Fields("CVE_PROV"), 0)
       Mat_derivDinero(i, 5) = ReemplazaVacioValor(rmesa.Fields("NUMSEC"), 0)
       Mat_derivDinero(i, 6) = ReemplazaVacioValor(rmesa.Fields("CVECONTPP"), 0)
       Mat_derivDinero(i, 7) = ReemplazaVacioValor(rmesa.Fields("FEC_OPERAC"), 0)
       Mat_derivDinero(i, 8) = ReemplazaVacioValor(rmesa.Fields("FEC_INI"), 0)
       Mat_derivDinero(i, 9) = ReemplazaVacioValor(rmesa.Fields("FEC_VENC"), 0)
       Mat_derivDinero(i, 10) = ReemplazaVacioValor(rmesa.Fields("TIPO_VALOR"), 0)
       Mat_derivDinero(i, 11) = ReemplazaVacioValor(rmesa.Fields("TIPO_OPER"), 0)
       Mat_derivDinero(i, 12) = ReemplazaVacioValor(rmesa.Fields("FEC_PROX_FLUJO"), 0)
       Mat_derivDinero(i, 13) = ReemplazaVacioValor(rmesa.Fields("PERIODO_FLUJO"), 0)
       Mat_derivDinero(i, 14) = ReemplazaVacioValor(rmesa.Fields("BASE_CAL_TASA"), 0)
       Mat_derivDinero(i, 15) = ReemplazaVacioValor(rmesa.Fields("CVEMONEDA"), 0)
       Mat_derivDinero(i, 16) = ReemplazaVacioValor(rmesa.Fields("TIPO_CAMBIO_REC"), 0)
       Mat_derivDinero(i, 17) = ReemplazaVacioValor(rmesa.Fields("M_NOC_FLUJO"), 0)
       Mat_derivDinero(i, 18) = ReemplazaVacioValor(rmesa.Fields("FORMULA_FLUJO"), 0)
       Mat_derivDinero(i, 19) = ReemplazaVacioValor(rmesa.Fields("TASA_RECIBE"), 0)
       Mat_derivDinero(i, 20) = ReemplazaVacioValor(rmesa.Fields("F_PROX_FLUJO_ENT"), 0)
       Mat_derivDinero(i, 21) = ReemplazaVacioValor(rmesa.Fields("PERIODICIDAD"), 0)
       Mat_derivDinero(i, 22) = ReemplazaVacioValor(rmesa.Fields("BASE_CALCULO"), 0)
       Mat_derivDinero(i, 23) = ReemplazaVacioValor(rmesa.Fields("TIPO_MONEDA"), 0)
       Mat_derivDinero(i, 24) = ReemplazaVacioValor(rmesa.Fields("TIPO_CAMBIO_ENT"), 0)
       Mat_derivDinero(i, 25) = ReemplazaVacioValor(rmesa.Fields("M_NOC_FLUJO_ENT"), 0)
       Mat_derivDinero(i, 26) = ReemplazaVacioValor(rmesa.Fields("TASA_REF_ACTIVA"), 0)
       Mat_derivDinero(i, 27) = ReemplazaVacioValor(rmesa.Fields("TASA_REF_PASIVA"), 0)
       Mat_derivDinero(i, 28) = ReemplazaVacioValor(rmesa.Fields("SOBRETASA_ACTIVA"), 0)
       Mat_derivDinero(i, 29) = ReemplazaVacioValor(rmesa.Fields("SOBRETASA_PASIVA"), 0)
       Mat_derivDinero(i, 30) = ReemplazaVacioValor(rmesa.Fields("TASA_ENTREGA"), 0)
       Mat_derivDinero(i, 31) = ReemplazaVacioValor(rmesa.Fields("PRIMA"), 0)
       Mat_derivDinero(i, 32) = ReemplazaVacioValor(rmesa.Fields("OBJETIVO_OPER"), 0)
       Mat_derivDinero(i, 33) = ReemplazaVacioValor(rmesa.Fields("VAL_ACTIVA"), 0)
       Mat_derivDinero(i, 34) = ReemplazaVacioValor(rmesa.Fields("VAL_PASIVA"), 0)
       Mat_derivDinero(i, 35) = ReemplazaVacioValor(rmesa.Fields("VALOR_NETO"), 0)
       Mat_derivDinero(i, 36) = ReemplazaVacioValor(rmesa.Fields("MARCA_MERCADO"), 0)
       Acum = Acum + Mat_derivDinero(i, 36)
       rmesa.MoveNext
   Next i
   rmesa.Close
   Call GuardaIkosDerivados(Mat_derivDinero, fecha)
Else
    MsgBox "Posición derivados al " & fecha & " aún no Disponible"
End If
If Acum = 0 Then
    MsgBox "Posición derivados al " & fecha & " so se encuentra valuado "
End If
  Screen.MousePointer = 0
End Sub

Function fecha_oracle(fecha_i)
Dim dd As String
Dim mm As String
Dim aa As String
Dim fechax As String
Dim fecha1 As String

    dd = Mid$(CDate(fecha_i), 1, 2)
    mm = Mid$(CDate(fecha_i), 4, 2)
    aa = Mid$(CDate(fecha_i), 7, 4)
    fechax = dd + "/" + mm + "/" + aa
    fecha1 = "to_date('" & fechax & "','dd/mm/yyyy') "
    fecha_oracle = fecha1
End Function

