Attribute VB_Name = "ModTablas"
Option Explicit

Sub TCuadroAnalisisTasas(ByRef rejilla1 As MSFlexGrid, ByRef rejilla2 As MSFlexGrid)
Dim i As Integer
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 rejilla1.Cols = 7
 rejilla1.Rows = NoGruposPort + 3
 rejilla1.TextMatrix(0, 1) = "Efecto tequila"
 rejilla1.TextMatrix(0, 2) = "efecto 11 sep"
 rejilla1.TextMatrix(0, 3) = "Ad Hoc 1"
 rejilla1.TextMatrix(0, 4) = "Ad Hoc 2"
 rejilla1.TextMatrix(0, 5) = "3 desviaciones estandar"
 rejilla1.TextMatrix(0, 6) = "6 desviaciones estandar"

 rejilla1.TextMatrix(NoGruposPort + 2, 0) = MatPortafolios(1, 2)
 rejilla1.RowHeight(0) = 600
 rejilla1.ColWidth(0) = 2000
 rejilla1.ColWidth(1) = 1500
 rejilla1.ColWidth(2) = 1500
 rejilla1.ColWidth(3) = 1500
 rejilla1.ColWidth(4) = 1500
 rejilla1.ColWidth(5) = 1500
 
 rejilla2.Rows = NoFactores + 1
 rejilla2.Cols = 5
 rejilla2.TextMatrix(0, 1) = "Tasa/Valor observado"
 rejilla2.TextMatrix(0, 2) = "Variación observada"
 rejilla2.TextMatrix(0, 3) = "Fecha variación observada"
 rejilla2.TextMatrix(0, 4) = "Tasa/Valor con incremento"
 
 For i = 1 To NoFactores
 rejilla2.TextMatrix(i, 0) = MatCaracFRiesgo(i).indFactor
 Next i
 rejilla2.RowHeight(0) = 600
 rejilla2.ColWidth(0) = 1800
 rejilla2.ColWidth(1) = 1400
 rejilla2.ColWidth(2) = 1400
 rejilla2.ColWidth(3) = 1400
 rejilla2.ColWidth(4) = 1400
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub TablaVARMarkDesg(rejilla As MSFlexGrid)
Dim i As Integer
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
For i = 1 To NoGruposPort
rejilla.TextMatrix(i, 1) = Format(MatVARMarkowitz(i, 1), "###,###,###,###,###,##0.00")
rejilla.TextMatrix(i, 2) = Format(MatVARMarkowitz(i, 2), "###,###,###,###,###,##0.00")
If Val(MatVARMarkowitz(i, 3)) <> 0 Then rejilla.TextMatrix(i, 3) = Format(MatVARMarkowitz(i, 3), "###,###,###,###,###,##0.0000")
Next i

rejilla.TextMatrix(NoGruposPort + 3, 1) = Format(MatTotalMarkowitz(1), "###,###,###,###,###,###,##0.00")


On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub TitulosResumenPosicion(rejilla2 As MSFlexGrid, rejilla3 As MSFlexGrid, rejilla4 As MSFlexGrid)
Dim nott As Integer
Dim i As Integer
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'son todos los cuadros resumen de var por los diferentes metodos

'este es el cuadro resumen de var
nott = UBound(MatGruposPortPos, 1)
rejilla2.Rows = 3
rejilla2.Cols = 24
rejilla2.TextMatrix(0, 1) = "Titulo"
rejilla2.TextMatrix(0, 1) = "Titulos Compra Directo"
rejilla2.TextMatrix(0, 2) = "Titulos Compra Reporto"
rejilla2.TextMatrix(0, 3) = "Titulos Venta Reporto"
rejilla2.TextMatrix(0, 4) = "Titulos Venta Directo"
rejilla2.TextMatrix(0, 5) = "Total Titulos"
rejilla2.TextMatrix(0, 6) = "Compras Directo/Activa"
rejilla2.TextMatrix(0, 7) = "Compras Reporto"
rejilla2.TextMatrix(0, 8) = "Ventas Reporto"
rejilla2.TextMatrix(0, 9) = "Ventas Directo/Pasiva"
rejilla2.TextMatrix(0, 10) = "Marca Mercado"
rejilla2.TextMatrix(0, 11) = "plazo compras directo"
rejilla2.TextMatrix(0, 12) = "plazo compras Reporto"
rejilla2.TextMatrix(0, 13) = "plazo ventas Reporto"
rejilla2.TextMatrix(0, 14) = "plazo ventas Directo"
rejilla2.TextMatrix(0, 15) = "VaR Markowitz"
rejilla2.TextMatrix(0, 16) = "VaR Montecarlo"
rejilla2.TextMatrix(0, 17) = "CVAR"
rejilla2.TextMatrix(0, 18) = "3 Desv. Est."
rejilla2.TextMatrix(0, 19) = "6 Desv. Est."
rejilla2.TextMatrix(0, 20) = "Crisis Tequila"
rejilla2.TextMatrix(0, 21) = "Crisis 11 sep."
rejilla2.TextMatrix(0, 22) = "Ad Hoc 1"
rejilla2.TextMatrix(0, 23) = "Ad Hoc 2"


rejilla2.RowHeight(0) = 500
rejilla2.ColWidth(0) = 2500
For i = 1 To 23
rejilla2.ColWidth(i) = 1500
Next i


On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Sub TitulosResumenBacktesting(rejilla1 As MSFlexGrid)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
rejilla1.Cols = 7
rejilla1.TextMatrix(0, 0) = "Fecha"
rejilla1.TextMatrix(0, 1) = "M Mercado"
rejilla1.TextMatrix(0, 2) = "M Mercado dia siguiente"
rejilla1.TextMatrix(0, 3) = "Variacion Observada"
rejilla1.TextMatrix(0, 4) = "Límite Inf. VaR"
rejilla1.TextMatrix(0, 5) = "Límite Sup. VaR"
rejilla1.TextMatrix(0, 6) = "Acierto"
rejilla1.ColWidth(0) = 1500
rejilla1.ColWidth(1) = 1800
rejilla1.ColWidth(2) = 1800
rejilla1.ColWidth(3) = 1800
rejilla1.ColWidth(4) = 1800
rejilla1.ColWidth(5) = 1800
rejilla1.ColWidth(6) = 1800
rejilla1.RowHeight(0) = 700
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub TitulosPosicion(a As MSFlexGrid)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
a.WordWrap = True
a.Cols = 12
a.RowHeight(0) = 700
a.ColWidth(0) = 1400
a.ColWidth(1) = 1400
a.ColWidth(2) = 1400
a.ColWidth(3) = 1400
a.ColWidth(4) = 1400
a.ColWidth(5) = 1400
a.ColWidth(6) = 1400
a.ColWidth(7) = 1400
a.ColWidth(8) = 1400

a.TextMatrix(0, 1) = "Clave de la posicion"
a.TextMatrix(0, 2) = "Clave de operacion"
a.TextMatrix(0, 3) = "Valor unitario"
a.TextMatrix(0, 4) = "MTM sucio"
a.TextMatrix(0, 5) = "Paste activa sucia"
a.TextMatrix(0, 6) = "parte pasiva sucia"
a.TextMatrix(0, 7) = "valor unitario limpio"
a.TextMatrix(0, 8) = "MTM limpio"
a.TextMatrix(0, 9) = "Parte activa limpia"
a.TextMatrix(0, 10) = "Parte pasiva limpia"
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Sub VerDatosTabla(ByRef mat() As Variant, ByRef mat2() As Variant, ByRef tabla As MSFlexGrid)
Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se muestran los valores de las tasas en pantalla
If IsArray(mat) Then
n = UBound(mat, 1)
m = UBound(mat, 2)
tabla.Rows = 2
tabla.Cols = 2
tabla.Rows = n + 1
tabla.Cols = m
For i = 1 To m - 1
tabla.TextMatrix(0, i) = mat2(i, 1)
Next i
For i = 1 To n
For j = 1 To m
If Not IsNull(mat(i, j)) Then
If mat(i, j) < 1 Then
tabla.TextMatrix(i, j - 1) = Format(100 * mat(i, j), "###,###,###,##0.0000")
Else
tabla.TextMatrix(i, j - 1) = mat(i, j)
End If
End If
Next j
Next i
tabla.RowHeight(0) = 1000
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub VerFactoresPos(ByRef mat() As Variant, ByRef obj As MSFlexGrid)
Dim noreg As Integer
Dim nocol As Integer
Dim i As Integer
Dim j As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
noreg = UBound(mat, 1)
nocol = UBound(mat, 2)
obj.Rows = noreg + 1
obj.Cols = nocol + 1
For i = 1 To noreg
For j = 1 To nocol
If Not EsVariableVacia(mat(i, j)) Then
 obj.TextMatrix(i, j) = mat(i, j)
End If
Next j
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub


Sub VerDetallesPosicion(ByRef a As MSFlexGrid, ByRef matpos() As propPosRiesgo)
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Call TitulosPosicion(frmDesglosePos.MSFlexGrid1)
' se muestra el desglose de la posicion una tabla del sistema
a.Rows = UBound(matpos, 1) + 1
For i = 1 To UBound(matpos, 1)
    a.TextMatrix(i, 1) = matpos(i).C_Posicion
    a.TextMatrix(i, 2) = matpos(i).c_operacion
    a.TextMatrix(i, 3) = MatPrecios(i, 1)
    a.TextMatrix(i, 4) = MatPrecios(i, 2)
    a.TextMatrix(i, 5) = MatPrecios(i, 3)
    a.TextMatrix(i, 6) = MatPrecios(i, 4)
    a.TextMatrix(i, 7) = MatPrecios(i, 5)
    a.TextMatrix(i, 8) = MatPrecios(i, 6)
    a.TextMatrix(i, 9) = MatPrecios(i, 7)
    a.TextMatrix(i, 10) = MatPrecios(i, 8)
    
Next i


On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ShowTablasVMark(ByRef matfriesgo() As Variant, ByRef matrend() As Double, ByRef matcov() As Double, ByRef delt() As Double, ByVal ndatos As Integer, ByVal ndias As Integer, ByRef rejilla1 As MSFlexGrid, ByRef rejilla2 As MSFlexGrid, ByRef rejilla3 As MSFlexGrid)
Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se muestran algunos resultados estadisticos
'se listan las series con los factores de riesgo
n = UBound(matrend, 1)
m = UBound(matrend, 2)
rejilla1.WordWrap = True
rejilla1.Rows = n + 2
rejilla1.Cols = m + 1

For j = 1 To m
 rejilla1.TextMatrix(0, j) = matfriesgo(1, j)
For i = 1 To n
 If matrend(i, j) < 1 Then
  rejilla1.TextMatrix(i, j) = Format(matrend(i, j) * 100, "###,###,##0.00000")
 Else
  rejilla1.TextMatrix(i, j) = Format(matrend(i, j), "###,###,###,##0.00")
 End If
Next i
rejilla1.ColWidth(j) = 1500
Next j
rejilla1.RowHeight(0) = 800

'la matriz de covarianzas
rejilla2.WordWrap = True
rejilla2.Cols = ndatos + 1
rejilla2.Rows = ndatos + 1
rejilla2.ColWidth(0) = 2200
For j = 1 To ndatos
 rejilla2.TextMatrix(0, j) = matfriesgo(1, j)
 rejilla2.TextMatrix(j, 0) = matfriesgo(1, j)
For i = 1 To ndatos
 rejilla2.TextMatrix(i, j) = Format(matcov(i, j) * 100, "###,##0.00000")
Next i
rejilla2.ColWidth(j) = 1800
Next j
rejilla2.RowHeight(0) = 800

'se muestran Sensibilidades y valores observados de los
'parametros

rejilla3.Rows = ndatos + 1
rejilla3.Cols = 4
rejilla3.TextMatrix(0, 1) = "Medias del portafolio"
rejilla3.TextMatrix(0, 2) = "Sens del portafolio"
rejilla3.TextMatrix(0, 2) = "Sens del portafolio"
For i = 1 To ndatos
 rejilla3.TextMatrix(i, 0) = matfriesgo(1, i)
'If Abs(mmedias(i, 1)) < 1 Then
'rejilla3.TextMatrix(1, i) = Format(mmedias(i, 1) * 100, "###,###,##0.00000")
'Else
'rejilla3.TextMatrix(1, i) = Format(mmedias(i, 1), "###,##0.0000")
'End If

rejilla3.TextMatrix(i, 2) = Format(delt(1, i), "###,###,###,##0.0000")
rejilla3.TextMatrix(i, 3) = Format((matcov(i, i)) ^ 0.5, "###,###,###,##0.0000")
rejilla3.RowHeight(i) = 500
Next i
rejilla3.ColWidth(0) = 1200
rejilla3.ColWidth(1) = 1200
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub VerRLimInfVHist(ByRef rejilla1 As MSFlexGrid, ByRef mat2() As Variant, ByRef rejilla2 As MSFlexGrid)
Dim n As Integer
Dim m As Integer
Dim i As Integer
Dim j As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

For i = 1 To NoGruposPort + 3
  
'   rejilla1.TextMatrix(i, 2) = Format(MatResumenVART(i, 6), "###,###,###,###,###,###,##0.00")

Next i
n = UBound(mat2, 1)
m = UBound(mat2, 2)
rejilla2.Rows = n + 1
rejilla2.Cols = m + 1
For i = 1 To n
For j = 1 To m
rejilla2.TextMatrix(i, j) = Format(mat2(i, j) * 100, "##0.000000")
Next j
Next i

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub MuestraPortRejilla(ByRef rejilla1 As MSFlexGrid)
Dim noreg As Integer
Dim i As Integer
Dim txtport As String
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Set base1 = OpenDatabase("", dbDriverNoPrompt, False, ";Pwd=" & ContraseñaCatalogos)
Set registros1 = base1.OpenRecordset("select * from [] where PORTAFOLIO= '" & txtport & "'", dbOpenDynaset, dbReadOnly)

If registros1.RecordCount <> 0 Then
 registros1.MoveLast
 noreg = registros1.RecordCount
 registros1.MoveFirst
 For i = 1 To noreg
 
 
 registros1.MoveNext
 Next i

End If

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub
