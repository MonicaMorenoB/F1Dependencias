Attribute VB_Name = "modVaRHistorico"
Option Explicit

Sub CalcEscHist(ByVal f_val As Date, _
                ByVal f_factor As Date, _
                ByVal htiempo As Integer, _
                ByVal noesc As Integer, _
                ByVal dfwd As Integer, _
                ByRef matpos() As propPosRiesgo, _
                ByRef matposmd() As propPosMD, _
                ByRef matposdiv() As propPosDiv, _
                ByRef matposswaps() As propPosSwaps, _
                ByRef matposfwd() As propPosFwd, _
                ByRef matflswap() As estFlujosDeuda, _
                ByRef matposdeuda() As propPosDeuda, _
                ByRef matfldeuda() As estFlujosDeuda, _
                ByRef matprecios0() As resValIns, ByRef matpyg() As Double)
                
Dim parval As New ParamValPos
Dim mrvalflujo() As resValFlujo
Dim matfactr() As Double
Dim matfrsim() As Double
Dim matx() As Variant
Dim matx1() As Double
Dim i As Integer
Dim j As Integer
Dim nprecios As Integer

Dim fechaf As Date
Dim exito1 As Boolean
Dim txtmsg As String
Dim exito As Boolean
If noescSH <> noesc Or htiempoSH <> htiempo Or fechaSH <> f_factor Then
   Call PrevCVaRHistorico(f_factor, noesc, htiempo, MatFactRiesgo, matrendsSH, matBndSH)
   noescSH = noesc
   htiempoSH = htiempo
   fechaSH = f_factor
End If
'se calculan los precios en t0 y con los fr originales
Set parval = DeterminaPerfilVal("HISTORICO")
parval.perfwd = dfwd
If f_factor <> FechaMatFactR1 Then
   MatFactR1 = CargaFR1Dia(f_factor, exito1)
   FechaMatFactR1 = f_factor
End If
'la valuacion en el escenario tabla
matprecios0 = CalcValuacion(f_val + dfwd, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg, exito)
If exito Then
   nprecios = UBound(matprecios0, 1)
ReDim matpyg(1 To noesc, 1 To nprecios) As Double
   For i = 1 To noesc
   'ahora se aplican los incrementos generados historicamente a la serie del dia de hoy
       matfrsim = GenEscHist2(MatFactR1, matrendsSH, matBndSH, i)
   'se valua la posicion con estos nuevos precios
       fechaf = f_val
       MatPrecios = CalcValuacion(fechaf + dfwd, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, matfrsim, MatCurvasT, parval, mrvalflujo, txtmsg, exito)
       For j = 1 To nprecios
           matpyg(i, j) = MatPrecios(j).mtm_sucio - matprecios0(j).mtm_sucio
       Next j
       AvanceProc = i / noesc
       MensajeProc = "Avance de la simulación Histórica: " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
End If
End Sub

Sub CalcEscHistVR(ByVal f_val As Date, _
                  ByVal f_factor As Date, _
                  ByVal htiempo As Integer, _
                  ByVal noesc As Integer, _
                  ByVal dfwd As Integer, _
                  ByRef matpos() As propPosRiesgo, _
                  ByRef matposmd() As propPosMD, _
                  ByRef matposdiv() As propPosDiv, _
                  ByRef matposswaps() As propPosSwaps, _
                  ByRef matposfwd() As propPosFwd, _
                  ByRef matflswap() As estFlujosDeuda, _
                  ByRef matposdeuda() As propPosDeuda, _
                  ByRef matfldeuda() As estFlujosDeuda, _
                  ByRef mprecios0() As resValIns, _
                  ByRef mprecios1() As resValIns, _
                  ByRef matpyg() As Double)
'calcula los escenarios historicos proyectados dfwd para una posicion de derivados
'todas las variables de posicion que carga son cargadas por un proceso previo
'datos de entrada
'f_val       - fecha de valuacion
'f_factor    - fecha de los factores de riesgo
'htiempo     - horizonte de tiempo para realizar el calculo de rendimientos
'noesc       - numero de escenarios para realizar el calculo
'las demas variables de posicion son necesarias para encapular los valores de la posicion cargada
'el resultados de esta rutina es la matriz matpyg()
                  
Dim parval As New ParamValPos
Dim mrvalflujo() As resValFlujo
Dim matfactr() As Double
Dim matfrsim() As Double
Dim matx() As Variant
Dim matx1() As Double
Dim i As Integer
Dim j As Integer
Dim nprecios As Integer
Dim exito1 As Boolean
Dim exito As Boolean
Dim bl_exito As Boolean
Dim txtmsg As String

If noescSH <> noesc Or htiempoSH <> htiempo Or fechaSH <> f_factor Then
   Call PrevCVaRHistorico(f_factor, noesc, htiempo, MatFactRiesgo, matrendsSH, matBndSH)
   noescSH = noesc
   htiempoSH = htiempo
   fechaSH = f_factor
End If
If f_factor <> FechaArchCurvas Or EsArrayVacio(MatCurvasT) Then
   FechaArchCurvas = f_factor
   MatCurvasT = LeerCurvaCompleta(f_factor, bl_exito)
End If
'se calculan los precios en t0 y con los fr originales
ValExacta = True
Set parval = DeterminaPerfilVal("VALUACION")
parval.perfwd = 0
If f_factor <> FechaMatFactR1 Then
   MatFactR1 = CargaFR1Dia(f_factor, exito1)
   FechaMatFactR1 = f_factor
End If
'la valuacion en el escenario futuro f_val + dfwd
mprecios0 = CalcValuacion(f_val, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg, exito)
ValExacta = False
Set parval = DeterminaPerfilVal("HISTORICO")
parval.perfwd = dfwd
nprecios = UBound(mprecios0, 1)
mprecios1 = CalcValuacion(f_val + dfwd, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, mrvalflujo, txtmsg, exito)

ReDim matpyg(1 To noesc, 1 To nprecios) As Double
For i = 1 To noesc
'ahora se aplican los incrementos generados historicamente a la serie del dia de hoy
    matfrsim = GenEscHist2(MatFactR1, matrendsSH, matBndSH, i)
'se valua la posicion con estos nuevos precios
    MatPrecios = CalcValuacion(f_val + dfwd, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, matfrsim, MatCurvasT, parval, mrvalflujo, txtmsg, exito)
    For j = 1 To nprecios
        matpyg(i, j) = MatPrecios(j).mtm_sucio - mprecios1(j).mtm_sucio
    Next j
    AvanceProc = i / noesc
    MensajeProc = "Avance de la simulación Histórica: " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
End Sub


Sub PrevCVaRHistorico(ByVal fecha As Date, ByVal noesc As Long, ByVal htiempo As Integer, ByRef mfriesgo() As Variant, ByRef matrends() As Double, ByRef matb() As Integer)
Dim indice0 As Long
Dim matx1() As Double
Dim matx() As Variant
Dim matfechas() As Date
Dim i As Integer
Dim j As Integer
    indice0 = BuscarValorArray(fecha, mfriesgo, 1)
    matx = ExtraerSMatFR(indice0, noesc + htiempo, mfriesgo, True, SiFactorRiesgo)
    matfechas = ConvArVtDT(ExtraeSubMatrizV(matx, 1, 1, 1 + htiempo, UBound(matx, 1)))
    matx1 = ConvArVtDbl(ExtraeSubMatrizV(matx, 2, UBound(matx, 2), 1, UBound(matx, 1)))
    Call GenRends3(matx1, htiempo, matfechas, matrends, matb)
End Sub


Function CalcVaRMark(ByRef matx() As Double, ByVal tvar As Integer, ByVal nconf As Double, ByVal lambda As Double)
Dim valfa As Double
Dim valor As Double
Dim val1() As Double
Dim matb() As Double
'tvar=0 var por percentil
'tvar=1 var por desviacion estandar
'valfa  valor correspondiente al nivel de confianza
'lambda valor de lambda
valfa = NormalInv(nconf)

ReDim mata(1 To 2) As Double
If tvar = 0 Then
   valor = CMedia2(matx, 1, "c") + (CVarianza2(matx, 1, "c")) ^ 0.5 * valfa
Else
   'promedios moviles ponderados exponencialmente
   val1 = GenMedias(matx, 1, lambda)
   matb = GenCovar(matx, matx, 1, lambda)
   valor = val1(1, 1) + (matb(1, 1)) ^ 0.5 * valfa
End If
CalcVaRMark = valor
End Function

Function GenEscHist2(ByRef mata() As Double, ByRef matr() As Double, ByRef matb() As Integer, ByVal ind1 As Integer) As Double()
Dim j As Integer
Dim valor As Double
Dim estaln As Boolean
ReDim mattas(1 To NoFactores, 1 To 1) As Double
For j = 1 To NoFactores
    If matb(ind1, j) = 1 Then
       mattas(j, 1) = mata(j, 1) + Abs(mata(j, 1)) * matr(ind1, j)
    Else
       mattas(j, 1) = mata(j, 1) + matr(ind1, j)
    End If
Next j
GenEscHist2 = mattas
End Function

Function Esblacklistfr(ByVal fecha As Date, ByVal txtfactor As String) As Boolean
'en funcion de la fecha y el nombre del factor indica si se encuentra en el catalogo
'cargada en memoria en MatBlEsc, si se encuentra, devuelve un valor true,
' en caso contrario devuelve un valor false
Dim i As Integer
For i = 1 To UBound(MatBlEsc, 1)
    If fecha = MatBlEsc(i, 1) And txtfactor = MatBlEsc(i, 2) Then
       Esblacklistfr = True
       Exit Function
    End If
Next i
Esblacklistfr = False
End Function



