Attribute VB_Name = "BaseDatos"
Option Explicit

Function LeerValDerivIKOS(ByVal fecha As Date) As Variant()
'objetivo: lee la valuacion de la tabla "TablaVDerIKOS" para la fecha de entrada fecha
'y la guarda en un array de 2 dimensiones
'los datos que hay en el array son
'la fecha de valuacion
'la clave de operacion
'valuacion de la pata activa
'valuacion de la pata pasiva
'marca a mercado de la operacion

'la dimension minima es 1 y la dimension maxima es el no de registros leidos
'si no hay datos devuelve un array de 2 dimensiones de 0 x 0

Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset


txtfecha = "TO_DATE('" & Format(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaVDerIKOS & " WHERE FECHA = " & txtfecha & " ORDER BY CLAVE"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 rmesa.Open txtfiltro2, ConAdo
 ReDim mata(1 To noreg, 1 To 5) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("FECHA")
     mata(i, 2) = rmesa.Fields("CLAVE")
     mata(i, 3) = rmesa.Fields("VAL_ACTIVAIK")
     mata(i, 4) = rmesa.Fields("VAL_PASIVAIK")
     mata(i, 5) = rmesa.Fields("MTMIK")
     rmesa.MoveNext
     AvanceProc = i / noreg
     MensajeProc = "Leyendo las valuaciones del sistema IKOS Derivados " & Format(AvanceProc, "##0.00 %")
 Next i
 rmesa.Close
 mata = RutinaOrden(mata, 2, SRutOrden)
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerValDerivIKOS = mata
End Function

Sub CalculaValPos(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal id_val As Integer, ByRef exito As Boolean)
'esta rutina junta los 2 procesos de lectura de posicion y valuacion
'datos de entrada
'f_pos    - fecha de la posicion
'f_factor - fecha de los factores de riesgo
'f_val    - fecha de valuacion
'txtport  - nombre del portafolio de posicion
'id_val   - modelo de valuacion a aplicar a derivados
'exito    - indica cuando el proceso tiene exito o no

Dim exito2 As Boolean
Dim exito1 As Boolean
Dim exito3 As Boolean
Dim exito4 As Boolean
Dim sicontinuart As Boolean
Dim mattxt() As String
Dim parval As ParamValPos
Dim txtmsg As String
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim txtmsg4 As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
    
    'se crea el array con los filtros de posicion
    mattxt = CrearFiltroPosPort(f_pos, txtport)
    Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
    'si el numero de registros leidos es mayor que cero se procede a la valuacion
    If UBound(matpos, 1) > 0 Then
       Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, id_val, txtmsg2, exito2)
       If exito2 Then
          'carga de factores de riesgo
          Call RutinaCargaFR(f_factor, exito3)
          'carga de curvas completas
          MatCurvasT = LeerCurvaCompleta(f_factor, exito)
          If exito3 Then
             FechaVPrecios = 0
             Call AnexarDatosVPrecios(f_val, matposmd)
             Set parval = DeterminaPerfilVal("VALUACION")
             MatPrecios = CalcValuacion(f_val, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactR1, MatCurvasT, parval, MatResValFlujo, txtmsg4, exito4)
             exito = True
             txtmsg = "El proceso finalizo correctamente"
          Else
             txtmsg = "No hay factores para esta fecha"
             exito3 = False
          End If
       End If
    Else
      txtmsg = "No hay registros para la posicion definida"
      exito = False
    End If
End Sub

Function LeerTAccess(objeto, ByVal campo As Integer, reg) As Variant
Dim var As Variant
On Error GoTo error1
var = objeto.Fields(campo)
LeerTAccess = var
Exit Function
error1:
'MsgBox "ha ocurrido un error en la lectura de datos en el campo " & campo & " en el registro" & reg
LeerTAccess = 0
End Function

Sub GrabarTAccess(objeto, campo, ByVal valor As Variant, reg)
On Error GoTo error1
'MsgBox objeto.Fields(campo).Type
objeto.Fields(campo) = valor
Exit Sub
error1:
MsgBox "Ha ocurrido un error en la grabacion de datos en el campo " & campo & " en el registro " & reg
End Sub

Sub CalculaVolatilidad(ByVal fvol As Integer, ByVal ndias1 As Integer, ByVal nconf As Double, ByVal novolatil As Integer, ByVal indice1 As Integer, ByVal indice2 As Integer, ByVal opvol As Integer, ByVal lambda As Double)
Dim matvolatil1() As Double
Dim matvolatil2() As Double
Dim matrends1() As Double
Dim matrends2() As Double
Dim i As Integer
Dim j As Integer
Dim NivelCritico As Double
Dim rejilla As MSFlexGrid
Dim rejilla3 As MSFlexGrid
Dim ivol As Long
Dim valor() As Double
Dim limi As Double
Dim lims As Double

' se realiza el calculo de las volatilidades
'esa rutina se debe de dejar aparte como un modulo
'para esto se debe de eliminar toda referencia de
'cualquier objeto
NivelCritico = NormalInv(nconf)

If indice1 = 0 Or indice2 = 0 Then
MensajeProc = "No se puede realizar el calculo de las volatilidades"
 Exit Sub
End If

Set rejilla = frmVolatilidades.MSFlexGrid4
ivol = BuscarValorArray(fvol, MatFactRiesgo, 1)
'ahora se revisa si se tiene suficiente datos
'hacia atras para realizar los calculos
If ivol = 0 Then
   MensajeProc = "Falta la fecha en la tabla de datos, se hara con la ultima fecha de la tabla"
   ivol = UBound(MatFactRiesgo, 1)
End If

frmVolatilidades.Combo4.Text = MatFactRiesgo(ivol, 1)
If ivol < ndias1 + novolatil + 1 Then
 MensajeProc = "No ha suficientes datos para realizar los calculos"
 Exit Sub
End If
'como si hay suficientes datos se leen los datos para el calculo
'de volatilidades
 matvolatil1 = ExtSerieFR(MatFactRiesgo, indice1, ivol, ndias1 + novolatil)
 matvolatil2 = ExtSerieFR(MatFactRiesgo, indice2, ivol, ndias1 + novolatil)
'SE CALCULAN los rendimientos, ya sean logaritmicos o
'aritmeticos

 matrends1 = CalculaRendimientoColumna(matvolatil1, 1)
 matrends2 = CalculaRendimientoColumna(matvolatil2, 1)

'Ahora si se procede a calcular medias y volatilidades
'aqui se usan 2 tecnicas: una rutina que obtiene la
'submatriz de la cual se va a obtener la media y la desviacion estandar
'y las funciones para obtener medias y deviaciones estandar
'de un vector que ya estan bien definidas, estos resultados
'a su vez se ponen en un vector llamado MatA
ErrorVarianza = 0
ReDim mata(1 To ndias1, 1 To 2) As Variant
For i = 1 To UBound(matrends1, 1) - novolatil + 1
 
 mata(i, 1) = GenMedias(ExtSerieAD(matrends1, 1, i + novolatil - 1, novolatil), opvol, lambda)
 valor = GenCovar(ExtSerieAD(matrends1, 1, i + novolatil - 1, novolatil), ExtSerieAD(matrends2, 1, i + novolatil - 1, novolatil), opvol, lambda)
 mata(i, 2) = valor(1, 1)
If i > 1 Then ErrorVarianza = ErrorVarianza + (matrends1(i, 1) * matrends2(i, 1) - mata(i - 1, 2) ^ 2) ^ 2
AvanceProc = i / ndias1
MensajeProc = "Calculando Medias y Volatilidades: " & Format(i / ndias1, "###,##0.00")

Next i

ErrorVarianza = (ErrorVarianza / ndias1) ^ 0.5
'Se muestran los rendimientos mas las volatilidades en pantalla
rejilla.Rows = 1
rejilla.Cols = 5
rejilla.Rows = ndias1 + novolatil + 1
rejilla.TextMatrix(0, 1) = "Valor"
rejilla.TextMatrix(0, 2) = "Rendimiento"
rejilla.TextMatrix(0, 3) = "Promedio movil"
rejilla.TextMatrix(0, 4) = "Varianza/Covarianza"
For i = 1 To ndias1 + novolatil
'la fecha y el precio
For j = 1 To 1
 rejilla.TextMatrix(i, j - 1) = matvolatil1(i, j)
Next j
'luego el rendimiento
If i > 1 Then rejilla.TextMatrix(i, 2) = Format(matrends1(i - 1, 1) * 100, "###,##0.0000")
'por ultimo las medias y volatilidades calculadas
If i <= UBound(mata, 1) Then
rejilla.TextMatrix(i + novolatil, 3) = Format(mata(i, 1) * 100, "###,##0.0000")
rejilla.TextMatrix(i + novolatil, 4) = Format(mata(i, 2) * 100, "###,##0.0000")
End If
Next i

'en funcion de las volatilidades se
'calculan los limites
frmVolatilidades.Text3.Text = Format(ErrorVarianza * 100, "###,###,###,###,##0.00000")
Set rejilla3 = frmVolatilidades.MSFlexGrid5


rejilla3.Cols = 5
rejilla3.Rows = ndias1
rejilla3.RowHeight(0) = 800
rejilla3.TextMatrix(0, 1) = "Precio"
rejilla3.TextMatrix(0, 2) = "Limite Inferior"
rejilla3.TextMatrix(0, 3) = "Limite Superior"
NoAciertos = 0
For i = 1 To ndias1 - 1
'se pone primero la fecha y el precio
rejilla3.TextMatrix(i, 0) = matvolatil1(i + novolatil, 1)
rejilla3.TextMatrix(i, 1) = matvolatil1(i + novolatil, 2)
If mata(i, 2) >= 0 Then
'el limite inferior
limi = matvolatil1(i + novolatil - 1, 2) * Exponen(mata(i, 1) - NivelCritico * (mata(i, 2) ^ 0.5))
'el limite superior
lims = matvolatil1(i + novolatil - 1, 2) * Exponen(mata(i, 1) + NivelCritico * (mata(i, 2) ^ 0.5))
Else
limi = 0: lims = 0
End If
rejilla3.TextMatrix(i, 2) = Format(limi, "###,###,##0.0000")
rejilla3.TextMatrix(i, 3) = Format(lims, "###,###,##0.0000")
 If matvolatil1(i + novolatil, 2) >= limi And matvolatil1(i + novolatil, 2) <= lims Then
  frmVolatilidades.MSFlexGrid5.TextMatrix(i, 4) = 1
  NoAciertos = NoAciertos + 1
 Else
  rejilla3.TextMatrix(i, 4) = 0
 End If
Next i
'Call MostrarMensajeSistema( NoAciertos / (ndias1 - 1) * 100
'se grafican las volatilidades en un gráfico de columnas
Set grafico = frmVolatilidades.MSChart1
grafico.chartType = 0
grafico.chartType = VtChChartType2dBar
grafico.ColumnCount = 1
grafico.RowCount = ndias1 - 1
For i = 2 To ndias1
grafico.row = i - 1
grafico.Column = 1
grafico.RowLabel = matvolatil1(i + novolatil, 1)
Next i

grafico.Title = "EFICIENCIA DEL CALCULO DE VOLATILIDADES"
grafico.Refresh

'se grafica las bandas de limites
Set grafico = frmVolatilidades.MSChart4
grafico.chartType = 0
grafico.chartType = VtChChartType2dLine
grafico.ColumnCount = 3
grafico.RowCount = ndias1 - 1
For i = 2 To ndias1
grafico.row = i - 1
grafico.Column = 1
grafico.Data = matvolatil1(i + novolatil, 2)
grafico.Column = 2
grafico.Data = matvolatil1(i + novolatil - 1, 2) * Exponen(mata(i, 1) - NivelCritico * Sqr(Abs(mata(i, 2))))
grafico.Column = 3
grafico.Data = matvolatil1(i + novolatil - 1, 2) * Exponen(mata(i, 1) + NivelCritico * Sqr(Abs(mata(i, 2))))
grafico.RowLabel = matvolatil1(i + novolatil, 1)
Next i
Call MostrarMensajeSistema(Format(NoAciertos / ndias1 * 100, "###,##0.00"), frmProgreso.Label2, 1, Date, Time, NomUsuario)
grafico.Title = "EFICIENCIA DEL CALCULO DE VOLATILIDADES"
grafico.Refresh
End Sub

Function LeerBitacoraOp(ByVal fecha As Date, ByVal idevento As Integer) As Variant()
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
'====================================================
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "select * from " & TablaBitacora
txtfiltro2 = txtfiltro2 & " WHERE FINICIO = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND ID_T_EVENTO = " & idevento
txtfiltro2 = txtfiltro2 & " ORDER BY HINICIO"
txtfiltro1 = "SELECT COUNT(*) from (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mata(1 To noreg, 1 To 11) As Variant
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("T_EVENTO")
       mata(i, 2) = rmesa.Fields("ID_PROCESO")
       mata(i, 3) = rmesa.Fields("DESCRIPCION")
       mata(i, 4) = rmesa.Fields("USUARIO")
       mata(i, 5) = rmesa.Fields("DIRECCION_IP")
       mata(i, 6) = rmesa.Fields("FECHAP")
       mata(i, 7) = rmesa.Fields("FINICIO")
       mata(i, 8) = rmesa.Fields("HINICIO")
       mata(i, 9) = rmesa.Fields("FFINAL")
       mata(i, 10) = rmesa.Fields("HFINAL")
       mata(i, 11) = rmesa.Fields("MENSAJE")
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerBitacoraOp = mata
End Function

Function LeerBitacoraIF(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Integer
Dim rmesa As New ADODB.recordset

'====================================================
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro1 = "select * from " & TablaBitacoraIF & " WHERE FECHA = " & txtfecha & " order by hora"
txtfiltro2 = "select count(*) from (" & txtfiltro1 & ")"
rmesa.Open txtfiltro2, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 4) As Variant
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
For i = 1 To noreg
    mata(i, 1) = rmesa.Fields("FECHA")
    mata(i, 2) = rmesa.Fields("HORA")
    mata(i, 3) = rmesa.Fields("USUARIO")
    mata(i, 4) = rmesa.Fields("NOINTENTOS")
    rmesa.MoveNext
Next i
rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerBitacoraIF = mata
End Function


Function LeerOpMesaO(ByVal txtport As String, ByVal txtmensaje As String)
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim nocampos As Long
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle

sql_num_mesa = "select count(*) from " & txtport
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close

If noreg <> 0 Then
sql_mesa = "select * from " & txtport
rmesa.Open sql_mesa, ConAdo
rmesa.MoveFirst
nocampos = rmesa.Fields.Count
ReDim mata(1 To noreg, 1 To nocampos) As Variant
For i = 1 To noreg
For j = 1 To nocampos
If IsNull(rmesa.Fields(j - 1)) Then          'tipo de operacion
    If rmesa.Fields(j - 1).Type = 131 Then
     mata(i, 1) = 0
    ElseIf rmesa.Fields(j - 1).Type = 129 Then
     mata(i, j) = ""
    ElseIf rmesa.Fields(j - 1).Type = 135 Then
     mata(i, j) = 0
    ElseIf rmesa.Fields(j - 1).Type = 200 Then
     mata(i, j) = 0
    ElseIf rmesa.Fields(j - 1).Type = 5 Then
     mata(i, j) = 0
    Else
     MsgBox "clasificame"
    End If
Else
    If rmesa.Fields(j - 1).Type = 131 Then
     mata(i, j) = Val(rmesa.Fields(j - 1))
    ElseIf rmesa.Fields(j - 1).Type = 129 Then
     mata(i, j) = Trim(rmesa.Fields(j - 1))
    ElseIf rmesa.Fields(j - 1).Type = 135 Then
     mata(i, j) = CDate(rmesa.Fields(j - 1))
    ElseIf rmesa.Fields(j - 1).Type = 5 Then
     mata(i, j) = Val(rmesa.Fields(j - 1))
    ElseIf rmesa.Fields(j - 1).Type = 200 Then
     mata(i, j) = Trim(rmesa.Fields(j - 1))
    Else
     MsgBox "clasificame"
    End If
End If
Next j
rmesa.MoveNext  'se lee el siguiente registro
 AvanceProc = i / noreg
 MensajeProc = txtmensaje & " " & Format(AvanceProc, "#,##0.00 %")
 DoEvents
Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerOpMesaO = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerBaseMD(ByVal txtfiltro As String) As propPosMD()
If ActivarControlErrores Then
On Error GoTo hayerror
End If
'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
Dim txtfiltro1 As String
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim mata() As New propPosMD
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset

If Not EsVariableVacia(txtfiltro) Then
txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg) As New propPosMD
rmesa.Open txtfiltro, ConAdo
rmesa.MoveFirst
For i = 1 To noreg
    mata(i).tipopos = rmesa.Fields("TIPOPOS")
    mata(i).nompos = rmesa.Fields("NOMPOS")
    mata(i).C_Posicion = Val(rmesa.Fields("CPOSICION"))
    mata(i).fechareg = rmesa.Fields("FECHAREG")
    mata(i).HoraRegOp = rmesa.Fields("HORAREG")
    mata(i).c_operacion = rmesa.Fields("COPERACION")
    mata(i).intencion = rmesa.Fields("INTENCION")
    mata(i).Tipo_Mov = ReemplazaVacioValor(rmesa.Fields("TOPERACION"), 0)
    mata(i).Signo_Op = TraducirTMov(mata(i).Tipo_Mov)
    mata(i).noTitulosMD = ReemplazaVacioValor(rmesa.Fields("NO_TITULOS"), 0)
    mata(i).tValorMD = ReemplazaVacioValor(rmesa.Fields("TV"), "")
    mata(i).emisionMD = ReemplazaVacioValor(rmesa.Fields("EMISION"), "")
    mata(i).serieMD = ReemplazaVacioValor(rmesa.Fields("SERIE"), "")
    mata(i).cEmisionMD = ReemplazaVacioValor(rmesa.Fields("C_EMISION"), "")
    mata(i).fCompraMD = ReemplazaVacioValor(rmesa.Fields("F_COMPRA"), 0)
    mata(i).pAsignadoMD = ReemplazaVacioValor(rmesa.Fields("P_COMPRA"), 0)
    mata(i).tReporto = ReemplazaVacioValor(rmesa.Fields("T_REPORTO"), 0)
    mata(i).fVencMD = ReemplazaVacioValor(rmesa.Fields("f_venc_oper"), 0)
    mata(i).subport1MD = ReemplazaVacioValor(rmesa.Fields("subport_1"), "")
    mata(i).SiFlujosMD = ReemplazaVacioValor(rmesa.Fields("SI_FLUJOS"), "N")
    mata(i).CalifLP = ReemplazaVacioValor(rmesa.Fields("CALIF_N_L"), 0)
    mata(i).escala = ReemplazaVacioValor(rmesa.Fields("ESCALA"), "")
    mata(i).Sector = ReemplazaVacioValor(rmesa.Fields("SECTOR"), "")
    mata(i).recupera = ReemplazaVacioValor(rmesa.Fields("RECUPERA"), 0)
    rmesa.MoveNext  'se lee el siguiente registro
    AvanceProc = i / noreg
    MensajeProc = "Leyendo la posicion de mercado de dinero: " & Format(AvanceProc, "##0.00 %")
Next i
 rmesa.Close
Else
 ReDim mata(0 To 0) As New propPosMD
End If
Else
 ReDim mata(0 To 0) As New propPosMD
End If
LeerBaseMD = mata
On Error GoTo 0
Exit Function
hayerror:
MsgBox "leerbasemd " & error(Err())
End Function

Function TraducirTMov(ByVal clave As Integer)
           If clave = 1 Or clave = 2 Or clave = 6 Then
              TraducirTMov = 1
           ElseIf clave = 3 Or clave = 4 Or clave = 7 Then
              TraducirTMov = -1
           Else
              TraducirTMov = 1
           End If
End Function

Function LeerTablaSwaps(ByVal txtfiltro As String)
On Error GoTo hayerror

Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset
'leer toda la tabla de datos de swaps sin filtrar
If Not EsVariableVacia(txtfiltro) Then
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   ReDim mata(1 To noreg) As New propPosSwaps
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i).tipopos = rmesa.Fields("TIPOPOS")                  'tipo de posicion
       mata(i).nompos = rmesa.Fields("NOMPOS")
       mata(i).C_Posicion = rmesa.Fields("CPOSICION")              'clave de la posicion
       mata(i).fechareg = rmesa.Fields("FECHAREG")                'fecha de registro
       mata(i).HoraRegOp = rmesa.Fields("HORAREG")                'HORA de registro de la op ikos
       mata(i).intencion = rmesa.Fields("INTENCION")              'intencion
       mata(i).EstructuralSwap = rmesa.Fields("ESTRUCTURAL")      'swap estructural
       mata(i).c_operacion = rmesa.Fields("COPERACION")            'clave de operacion
       mata(i).Tipo_Mov = rmesa.Fields("TOPERACION")
       mata(i).Signo_Op = TraducirTMov(mata(i).Tipo_Mov)
       mata(i).FCompraSwap = rmesa.Fields("FINICIO")               'fecha de compra
       mata(i).FvencSwap = rmesa.Fields("FVENCIMIENTO")            'fecha de vencimiento
       mata(i).IntercIFSwap = rmesa.Fields("INTER_I")              'intercambio inicial de flujos
       mata(i).IntercFFSwap = rmesa.Fields("INTER_F")              'intercambio intermedio y final de flujos
       mata(i).RIntAct = rmesa.Fields("AC_INT_ACT")                'acumula intereses activa
       mata(i).RIntPas = rmesa.Fields("AC_INT_PAS")                'acumula intereses pasiva
       mata(i).TCActivaSwap = rmesa.Fields("TC_ACTIVA")            'tasa referencia activa
       mata(i).TCPasivaSwap = rmesa.Fields("TC_PASIVA")            'tasa referencia pasiva
       mata(i).STActiva = rmesa.Fields("ST_ACTIVA")                'sobretasa cupon activa
       mata(i).STPasiva = rmesa.Fields("ST_PASIVA")                'sobretasa cupon pasiva
       mata(i).ConvIntAct = rmesa.Fields("CONV_INT_ACT")           'dias comerciales activa
       mata(i).ConvIntPas = rmesa.Fields("CONV_INT_PAS")           'dias comerciales pasiva
       mata(i).ClaveProdSwap = ReemplazaVacioValor(rmesa.Fields("CPRODUCTO"), "")          'clave de producto
       mata(i).cProdSwapGen = rmesa.Fields("FVALUACION")           'llave de valuacion
       mata(i).c_em_pidv = ReemplazaVacioValor(rmesa.Fields("C_EM_PIDV"), "")              'clave de emision cubierta
       mata(i).ID_ContrapSwap = rmesa.Fields("ID_CONTRAP")         'clave de contraparte
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo la posición de swaps: " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   rmesa.Close
   MensajeProc = "Se leyeron " & noreg & " de la posición de swaps "
Else
   ReDim mata(0 To 0) As New propPosSwaps
End If
Else
   ReDim mata(0 To 0) As New propPosSwaps
End If
LeerTablaSwaps = mata
On Error GoTo 0
Exit Function
hayerror:
MsgBox "LeerTablaSwaps" & error(Err())
End Function

Function LeerTablaDeuda(ByVal txtfiltro As String)
On Error GoTo hayerror

Dim txtfiltro1 As String
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset

If Not EsVariableVacia(txtfiltro) Then
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   nocampos = 46
ReDim mata(1 To noreg) As New propPosDeuda
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i).tipopos = rmesa.Fields("TIPOPOS")             'tipo de posicion
       mata(i).nompos = ReemplazaVacioValor(rmesa.Fields("NOMPOS"), "")
       mata(i).C_Posicion = rmesa.Fields("CPOSICION")         'clave de posicion
       mata(i).fechareg = rmesa.Fields("FECHAREG")           'fecha de registro
       mata(i).HoraRegOp = rmesa.Fields("HORAREG")           'HORA DE registro de la op ikos
       mata(i).c_operacion = rmesa.Fields("COPERACION")       'clave de operacion
       mata(i).Tipo_Mov = rmesa.Fields("TOPERACION")          'tipo de operacion
       mata(i).Signo_Op = TraducirTMov(mata(i).Tipo_Mov)
       mata(i).FinicioDeuda = rmesa.Fields("FINICIO")
       mata(i).FVencDeuda = rmesa.Fields("FVENCIMIENTO")
       mata(i).InteriDeuda = rmesa.Fields("INTER_I")
       mata(i).InterfDeuda = rmesa.Fields("INTER_F")
       mata(i).RintDeuda = rmesa.Fields("AC_INT")
       mata(i).TRefDeuda = rmesa.Fields("TCUPON")
       mata(i).SpreadDeuda = rmesa.Fields("SOBRETASA")
       mata(i).ConvIntDeuda = rmesa.Fields("CONV_INT")
       mata(i).ProductoDeuda = ReemplazaVacioValor(rmesa.Fields("CPRODUCTO"), "")
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo las operaciones de credito y emisiones: " & Format(AvanceProc, "##0.00 %")
       DoEvents
  Next i
  rmesa.Close
  MensajeProc = "Se leyeron " & noreg & " de la posición de deuda "
 'ordenar por clave de emision
 Else
 ReDim mata(0 To 0) As New propPosDeuda
End If
Else
 ReDim mata(0 To 0) As New propPosDeuda
End If
LeerTablaDeuda = mata
On Error GoTo 0
Exit Function
hayerror:
MsgBox "LeerTabladeuda " & error(Err())
End Function

Function LeerTablaPosFwd(ByVal txtfiltro As String)
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset
'1 tipo de posicion
'2 clave de posicion
'3 fecha registro
'4 clave de operacion
'5 intencion
If Not EsVariableVacia(txtfiltro) Then
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
rmesa.Open txtfiltro, ConAdo
ReDim mata(1 To noreg) As New propPosFwd
rmesa.MoveFirst
For i = 1 To noreg
    mata(i).tipopos = rmesa.Fields("tipopos")
    mata(i).fechareg = CDate(rmesa.Fields("FECHAREG"))
    mata(i).nompos = rmesa.Fields("NOMPOS")
    mata(i).HoraRegOp = rmesa.Fields(3)                       'fecha de registro de la op ikos
    mata(i).intencion = rmesa.Fields("intencion")
    mata(i).EstructuralFwd = rmesa.Fields("ESTRUCTURAL")          'DERIVADO ESTRUCTURAL
    mata(i).ReclasificaFwd = rmesa.Fields("RECLASIFICA")          'OPERACION DEL PORT DE RECLASIFICACION
    mata(i).C_Posicion = Val(rmesa.Fields("cposicion"))
    mata(i).c_operacion = rmesa.Fields("coperacion")           'Clave de operación
    mata(i).Tipo_Mov = rmesa.Fields("TOPERACION")                    'tipo de operacion 1 larga 4 corta
    mata(i).Signo_Op = TraducirTMov(mata(i).Tipo_Mov)
    mata(i).MontoNocFwd = rmesa.Fields("m_nocional")
    mata(i).FCompraFwd = rmesa.Fields("finicio")                 'fecha de inicio del
    If Not IsNull(rmesa.Fields("fvencimiento")) Then
       mata(i).FVencFwd = rmesa.Fields("fvencimiento")       'fecha de vencimiento
    Else
       mata(i).FVencFwd = 0
    End If
    mata(i).PAsignadoFwd = rmesa.Fields("ppactado")               'TIPO CAMBIO PACTADO
    mata(i).ClaveProdFwd = rmesa.Fields("CPRODUCTO")               'clave de tipo de operacion
    mata(i).ID_ContrapFwd = rmesa.Fields("ID_CONTRAP")            'OPERACION DEL PORT DE RECLASIFICACION
 rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Leyendo la tabla de forwards : " & Format(AvanceProc, "##0.00 %")
 DoEvents
 Next i
 rmesa.Close
 MensajeProc = "Se leyeron " & noreg & " de la posición de forwards "
Else
 ReDim mata(0 To 0) As New propPosFwd
End If
Else
 ReDim mata(0 To 0) As New propPosFwd
End If
 LeerTablaPosFwd = mata
End Function

Function LeerPosDiv(ByVal txtfiltro As String) As propPosDiv()
On Error GoTo hayerror

'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
Dim txtfiltro1 As String
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

If Not EsVariableVacia(txtfiltro) Then
txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg) As New propPosDiv
rmesa.Open txtfiltro, ConAdo
rmesa.MoveFirst
For i = 1 To noreg
    mata(i).tipopos = rmesa.Fields("TIPOPOS")
    mata(i).nompos = rmesa.Fields("NOMPOS")
    mata(i).C_Posicion = Val(rmesa.Fields("CPOSICION"))
    mata(i).fechareg = rmesa.Fields("FECHAREG")
    mata(i).HoraRegOp = rmesa.Fields("HORAREG")          'fecha de registro de la op ikos
    mata(i).c_operacion = rmesa.Fields("COPERACION")
    mata(i).intencion = rmesa.Fields("INTENCION")
    If IsNull(rmesa.Fields("toperacion")) Then
       mata(i).Tipo_Mov = 0
    Else
       mata(i).Tipo_Mov = rmesa.Fields("TOPERACION")
    End If
    mata(i).Signo_Op = TraducirTMov(mata(i).Tipo_Mov)
    If IsNull(rmesa.Fields("TV")) Then
       mata(i).TValorDiv = 0
    Else
       mata(i).TValorDiv = rmesa.Fields("TV")
    End If
If IsNull(rmesa.Fields("EMISION")) Then
   mata(i).EmisionDiv = 0
Else
 mata(i).EmisionDiv = rmesa.Fields("EMISION")
End If
If IsNull(rmesa.Fields("SERIE")) Then
   mata(i).SerieDiv = 0
Else
 mata(i).SerieDiv = rmesa.Fields("SERIE")
End If
If IsNull(rmesa.Fields("C_EMISION")) Then
   mata(i).CEmisionDiv = 0
Else
 mata(i).CEmisionDiv = rmesa.Fields("C_EMISION")
End If
If IsNull(rmesa.Fields("NO_titulos")) Then
  mata(i).MontoNocDiv = 0
Else
  mata(i).MontoNocDiv = Val(rmesa.Fields("NO_titulos"))
End If
If IsNull(rmesa.Fields("f_compra")) Then
 mata(i).FCompraDiv = 0
Else
 mata(i).FCompraDiv = CDate(rmesa.Fields("f_compra"))
End If
   rmesa.MoveNext  'se lee el siguiente registro
   AvanceProc = i / noreg
   MensajeProc = "Leyendo la posicion de divisas: " & Format(AvanceProc, "##0.00 %")
Next i
 rmesa.Close
Else
 ReDim mata(0 To 0) As New propPosDiv
End If
Else
ReDim mata(0 To 0) As New propPosDiv
End If
LeerPosDiv = mata
Exit Function
hayerror:
MsgBox "Leerposdiv " & error(Err())
End Function



Sub CargaLimites()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Dim txtfiltro As String
Dim txtcadena As String
Dim i As Integer
Dim rprecios As New ADODB.recordset

'====================================================
txtfiltro = "select * from " & PrefijoBD & TablaLimites
txtcadena = "select count(*) from (" & txtfiltro & ")"
rprecios.Open txtcadena, ConAdo
NoCapitalBase = rprecios.Fields(0)
rprecios.Close
If NoCapitalBase <> 0 Then
txtcadena = txtfiltro
rprecios.Open txtcadena, ConAdo
ReDim MatCapitalSist(1 To NoCapitalBase, 1 To 5) As Variant
rprecios.MoveFirst
For i = 1 To NoCapitalBase
 MatCapitalSist(i, 1) = rprecios.Fields("FECHA")
 MatCapitalSist(i, 2) = rprecios.Fields("INI_VIGENCIA")
 MatCapitalSist(i, 3) = rprecios.Fields("FIN_VIGENCIA")
 MatCapitalSist(i, 4) = rprecios.Fields("CONCEPTO")
 MatCapitalSist(i, 5) = rprecios.Fields("VALOR")
 rprecios.MoveNext
 AvanceProc = i / NoCapitalBase
 MensajeProc = "Leyendo los limites de VaR : " & Format(AvanceProc, "###,##0.00 %")
Next i
rprecios.Close
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function CargaTablaD(ByVal txttabla As String, ByVal texto As String, ByVal ncol As Integer)
Dim txtcadena As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim contar As Integer
Dim nocampos As Integer
Dim rprecios As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'====================================================
txtcadena = "select count(*) from " & txttabla
rprecios.Open txtcadena, ConAdo
noreg = rprecios.Fields(0)
rprecios.Close
If noreg <> 0 Then
   txtcadena = "select * from " & txttabla
   rprecios.Open txtcadena, ConAdo
   nocampos = rprecios.Fields.Count
   ReDim mata(1 To nocampos, 1 To 1) As Variant
   rprecios.MoveFirst
   contar = 0
   For i = 1 To noreg
       If Not EsVariableVacia(rprecios.Fields(ncol - 1)) Then
          contar = contar + 1
          ReDim Preserve mata(1 To nocampos, 1 To contar) As Variant
          For j = 1 To nocampos
              mata(j, contar) = rprecios.Fields(j - 1)
          Next j
       End If
       rprecios.MoveNext
       AvanceProc = i / noreg
       MensajeProc = texto & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   rprecios.Close
   mata = RutinaOrden(MTranV(mata), ncol, SRutOrden)
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaTablaD = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CargaTablaD2(ByVal txttabla As String, ByVal texto As String, ByVal ncol As Integer) As Variant()
'====================================================
Dim txtcadena As String
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim contar As Long
Dim nocampos As Long
Dim rprecios As New ADODB.recordset

txtcadena = "select count(*) from " & txttabla
rprecios.Open txtcadena, conAdo2
noreg = rprecios.Fields(0)
rprecios.Close
If noreg <> 0 Then
   txtcadena = "select * from " & txttabla
   rprecios.Open txtcadena, conAdo2
   nocampos = rprecios.Fields.Count
   ReDim mata(1 To nocampos, 1 To 1) As Variant
   rprecios.MoveFirst
   contar = 0
   For i = 1 To noreg
       If Not EsVariableVacia(rprecios.Fields(ncol - 1)) Then
          contar = contar + 1
          ReDim Preserve mata(1 To nocampos, 1 To contar) As Variant
          For j = 1 To nocampos
              mata(j, contar) = rprecios.Fields(j - 1)
          Next j
       End If
       rprecios.MoveNext
       AvanceProc = i / noreg
       MensajeProc = texto & " " & Format(AvanceProc, "##0.00 %")
   Next i
   rprecios.Close
   mata = RutinaOrden(MTranV(mata), ncol, SRutOrden)
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaTablaD2 = mata
End Function


Sub CrearPosMDCSV(ByVal fecha As Date, ByVal siarch As Boolean, ByVal direc As String, ByRef nrmesa As Long, ByRef exito As Boolean)
Dim siposmesad As Boolean
Dim noreg As Long
Dim siesfv As Boolean
Dim i As Long
Dim j As Long
Dim nomarch As String
Dim matpos() As Variant
Dim txtcadena As String
Dim exitoarch As Boolean

'se importa la posicion total para una fecha
'de la mesa de dinero desde una tabla remota en oracle
exito = False
siposmesad = False
nrmesa = 0
'se lee la posicion desde la red
 siesfv = EsFechaVaR(fecha)
 If siesfv Then
   matpos = LeerPosMDOrigen(fecha, direc, siarch)
   noreg = UBound(matpos, 1)
   If noreg > 0 Then
       nomarch = direc & "\TESORERIA_ACTIVA" & Format(fecha, "YYYYMMDD") & ".CSV"
       Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
       If exitoarch Then
       For i = 1 To noreg
          txtcadena = ""
          For j = 1 To UBound(matpos, 2)
            txtcadena = txtcadena & matpos(i, j)
            If j < UBound(matpos, 2) Then
               txtcadena = txtcadena & ","
            End If
          Next j
          Print #1, txtcadena
       Next i
       Close #1
       MensajeProc = "Se importaron " & noreg & " registros de la posicion de mesa de dinero de la fecha " & Format(fecha, "dd/mm/yyyy")
       exito = True
       End If
   Else
        MsgBox "No se encontraron registros de la posicion de Mercado de Dinero para la fecha especificada"
   End If
   nrmesa = nrmesa + noreg
 Else
  MsgBox "No se han includo la fecha en las fechas de calculo de VaR"
 End If
End Sub

Sub ImportarPosMDinero(ByVal fecha As Date, ByVal noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
'se importa la posicion total para una fecha
'de la mesa de dinero desde una tabla remota en oracle
'se le agrega la calificacion segun las reglas del marco de operacion y el fondo de pensiones
'se verifica que los instrumentos del array se puedan valuar
'finalmente se procede a guardar la información en la tabla de la posicion de
'instrumentos de deuda


'datos de entrada:
'fecha  - fecha de la posicion
'noreg  - numeros de registros
'txtmsg - mensajes emitidos durante la ejecución del proceso
'exito  - estado del proceso al finalizar

Dim siposmesad As Boolean
Dim noreg1 As Long
Dim txtfecha As String
Dim txtfiltro As String
Dim matpos() As New propPosMD
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim sierrores As Boolean
Dim i As Long
Dim indice As Long
Dim indice1 As Long
Dim txtcalifica As String
Dim resp As Integer
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim txtmsg1 As String
Dim txtmsg2 As String
Dim txtmsg3 As String
Dim matr() As Variant
Dim mrecuperaI() As Variant
Dim mrecuperaN() As Variant
Dim escala As String
Dim Sector As String

exito = False
siposmesad = False
noreg = 0
'se lee la posicion desde la red
   exito1 = True
   exito2 = True
   txtmsg = "El proceso finalizo correctamente"
   matpos = LeerPosMDIKOS(fecha, conAdoBD)
   matvp = LeerVPrecios(fecha, mindvp)
   If UBound(matpos, 1) > 0 And UBound(matvp, 1) <> 0 Then
      
      For i = 1 To UBound(matpos, 1)
          indice = BuscarValorArray(matpos(i).cEmisionMD, mindvp, 1)
          If indice <> 0 Then
             indice1 = mindvp(indice, 2)
             matpos(i).CalifFP = DefinirCalifEmFP(fecha, matpos(i).cEmisionMD, indice1, matvp)
             matpos(i).CalifTMD = AsignaCalif(matvp(indice1).calif_sp, matvp(indice1).calif_moodys, matvp(indice1).calif_fitch, matvp(indice1).calif_hr)
             matpos(i).SiFlujosMD = ConvBolStr(DetermSiEmFlujos(matpos(i).tValorMD, matpos(i).emisionMD, matpos(i).serieMD))
             matpos(i).CalifLP = DetCalificacionLPMD(matvp(indice1).calif_sp, matvp(indice1).calif_fitch, matvp(indice1).calif_moodys, matvp(indice1).calif_hr)
          Else
             txtmsg1 = "No se encontro el instrumento " & matpos(i).cEmisionMD & " en el vector de precios"
             exito1 = False
          End If
      Next i
      Call ClasTProdMD2(matpos, txtmsg2, exito2)
      matr = LeerTablaSectorEscEm(fecha)
      mrecuperaI = CargaRecuperacion(fecha, PrefijoBD & TablaRecInt)
      mrecuperaN = CargaRecuperacion(fecha, PrefijoBD & TablaRecNacional)
      txtmsg3 = ""
      For i = 1 To UBound(matpos, 1)
          Call AnexarSectorEsc(fecha, matpos(i).tValorMD, matpos(i).emisionMD, escala, Sector, matr, exito3)
          If exito3 Then
             matpos(i).escala = escala
             matpos(i).Sector = Sector
             If escala = "N" Then
                matpos(i).recupera = Recuperacion(matpos(i).CalifLP, mrecuperaN, Sector)
             Else
                matpos(i).recupera = Recuperacion(matpos(i).CalifLP, mrecuperaI, Sector)
             End If
          Else
            txtmsg3 = txtmsg3 & "No se encontro " & matpos(i).tValorMD & matpos(i).emisionMD & " en la tabla " & PrefijoBD & TablaSectorEscEm & ","
            Exit For
          End If
      Next i
      If exito1 And exito2 And exito3 Then
         txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
         txtfiltro = TablaPosMD & " WHERE FECHAREG = " & txtfecha & " AND TIPOPOS = 1"
         Call GuardaPosMD(txtfiltro, matpos, 1, "Real", "000000", exito, noreg)
         MensajeProc = "Se importaron " & noreg & " registros de la posicion de mesa de dinero de la fecha " & Format(fecha, "dd/mm/yyyy")
         txtmsg = "El proceso finalizo correctamente"
         exito = True
      Else
         txtmsg = txtmsg1 & "," & txtmsg2 & "," & txtmsg3
         exito = False
      End If
   Else
     exito = False
     MensajeProc = "No se encontraron registros de la posicion de mesa de dinero de la fecha " & Format(fecha, "dd/mm/yyyy")
     txtmsg = MensajeProc
   End If
End Sub

Function LeerTablaSectorEscEm(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset
Dim noreg As Integer
Dim i As Integer
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"  'fecha de registro
txtfiltro2 = "SELECT * FROM " & PrefijoBD & TablaSectorEscEm & " WHERE (F_REGISTRO,TV,EMISORA) IN"
txtfiltro2 = txtfiltro2 & " (SELECT MAX(F_REGISTRO),TV,EMISORA FROM " & PrefijoBD & TablaSectorEscEm
txtfiltro2 = txtfiltro2 & " WHERE F_REGISTRO <= " & txtfecha & " GROUP BY TV,EMISORA)"
txtfiltro2 = txtfiltro2 & " ORDER BY TV,EMISORA"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 4) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("TV")
       mata(i, 2) = rmesa.Fields("EMISORA")
       If mata(i, 2) = "GASN" Then mata(i, 2) = "NM"
       mata(i, 3) = rmesa.Fields("ESCALA")
       mata(i, 4) = rmesa.Fields("SECTOR")
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerTablaSectorEscEm = mata
End Function

Sub AnexarSectorEsc(ByVal fecha As Date, ByVal tv As String, ByVal emision As String, ByRef escala As String, ByRef Sector As String, ByRef mata() As Variant, ByRef exito As Boolean)
Dim i As Integer
escala = ""
Sector = ""
For i = 1 To UBound(mata, 1)
    If mata(i, 1) = tv And mata(i, 2) = emision Then
       If mata(i, 3) = "Local" Then
          escala = "N"
       Else
          escala = "I"
       End If
       Sector = mata(i, 4)
       exito = True
       Exit Sub
    End If
Next i
If EsVariableVacia(escala) Or EsVariableVacia(Sector) Then
  exito = False
End If
End Sub


Sub AnalisisPosMD(ByVal fecha As Date)
  Dim matvp() As New propVecPrecios
  Dim mindvp() As Variant
  matvp = LeerVPrecios(fecha, mindvp)

End Sub


Function LeerPosMDOrigen(ByVal fecha As Date, ByVal direc As String, ByVal siarch As Boolean)
Dim nomarch As String
Dim sihayarch As Boolean

If siarch Then
    If Right(direc, 1) <> "\" Then
       nomarch = direc & "\pos" & Format(fecha, "yyyymmdd") & ".csv"
    Else
       nomarch = direc & "pos" & Format(fecha, "yyyymmdd") & ".csv"
    End If
    sihayarch = VerifAccesoArch(nomarch)
    If sihayarch Then
     LeerPosMDOrigen = LeerArchTexto(nomarch, ",", "Leyendo la posicion de la Mesa de Dinero")
    Else
     ReDim mata(0 To 0, 0 To 0) As Variant
     LeerPosMDOrigen = mata
    End If
Else
    LeerPosMDOrigen = LeerPosMDIKOS(fecha, conAdoBD)
End If
End Function

Function LeerPosMDIKOS(ByVal fecha As Date, ByRef conex As ADODB.Connection) As propPosMD()
'lee la posicion de de instrumentos de deuda de Banobras
'Datos de entrada: la fecha del proceso
'la conexion a la base de datos, para poder leer de diferentes cadenas de conexion a la base de datos
'condiciones de datos de entrada
'fecha: tiene que se una fecha habil en el calendario mexicano
'conex: conexion a la base de datos de DRIESGOS o RIESGOSD
'Resultado un array con los datos de la posición de instrumentos de deuda de Banobras

Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim tiempo As Date
Dim noreg1 As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim toper As Integer
Dim cpos As Integer
Dim port As String
Dim mata() As New propPosMD
Dim contar As Integer
Dim rmesa As New ADODB.recordset
'intenta 2 conexiones consecutivas a la vista de datos
'si noreg<>noreg1, se considera que la posición no esta disponible

txtfecha = Format(fecha, "yyyymmdd")
txtfiltro2 = "SELECT * FROM " & PrefijoBD & TablaPosMesaIKOS & " WHERE F_POSICION = '" & txtfecha & "'"
txtfiltro2 = txtfiltro2 & " AND (NOMPORTAFOLIO = 'TITULOS A NEGOCIAR'"
txtfiltro2 = txtfiltro2 & " OR NOMPORTAFOLIO = 'A VENCIMIENTO'"
txtfiltro2 = txtfiltro2 & " OR NOMPORTAFOLIO = 'DISPONIBLES PARA LA VENTA'"
txtfiltro2 = txtfiltro2 & " OR NOMPORTAFOLIO = 'PI DISPONIBLE PARA LA VENTA'"
txtfiltro2 = txtfiltro2 & " OR NOMPORTAFOLIO = 'PI CONSERVADOS A VENCIMIENTO'"
txtfiltro2 = txtfiltro2 & " OR NOMPORTAFOLIO = 'PI DERIVADOS'"
txtfiltro2 = txtfiltro2 & " OR NOMPORTAFOLIO = 'FONDOS DE PENSIONES')"
txtfiltro2 = txtfiltro2 & " ORDER BY M_POSICION,T_OPERACION,T_VALOR,EMISION,SERIE"
txtfiltro1 = "SELECT count(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, conex
noreg = rmesa.Fields(0)
rmesa.Close
tiempo = Time
Do While Time < (tiempo + 0.001) And noreg <> 0
Loop
rmesa.Open txtfiltro1, conex
noreg1 = rmesa.Fields(0)
rmesa.Close
If noreg <> noreg1 Then
   ReDim mata(0 To 0)
   LeerPosMDIKOS = mata
   Exit Function
End If
If noreg <> 0 Then
   rmesa.Open txtfiltro2, conex
   rmesa.MoveFirst
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To 1)
   For i = 1 To noreg
       If Trim(rmesa.Fields("T_VALOR")) <> "CM" Then
          contar = contar + 1
          ReDim Preserve mata(1 To contar)
          mata(contar).fechareg = fecha
          mata(contar).intencion = Trim(rmesa.Fields("INTENCION"))
          mata(contar).tValorMD = Trim(rmesa.Fields("T_VALOR"))
          mata(contar).emisionMD = Trim(rmesa.Fields("EMISION"))
          mata(contar).serieMD = Trim(rmesa.Fields("SERIE"))
          mata(contar).cEmisionMD = GeneraClaveEmision(mata(contar).tValorMD, mata(contar).emisionMD, mata(contar).serieMD)
          mata(contar).tReporto = Val(rmesa.Fields("T_PREMIO")) / 100
          mata(contar).fCompraMD = ConvertirTextoFecha(Trim(rmesa.Fields("F_COMPRA")), 0)
          mata(contar).pAsignadoMD = Trim(rmesa.Fields("P_COMPRA"))
          mata(contar).fVencMD = ConvertirTextoFecha(rmesa.Fields("F_VENCIMIENTO"), 0)
          mata(contar).noTitulosMD = Val(rmesa.Fields("N_TITULOS"))
          mata(contar).Tipo_Mov = TradCPosMD(rmesa.Fields("T_OPERACION"))
          toper = Val(rmesa.Fields("T_OPERACION"))
          cpos = Val(rmesa.Fields("M_POSICION"))
          port = Trim(rmesa.Fields("NOMPORTAFOLIO"))
          mata(contar).C_Posicion = DefinirPosMD(toper, cpos, port)
          mata(contar).c_operacion = contar
       End If
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Generando la posicion de mesa de dinero del " & fecha & " " & Format(AvanceProc, "#,##0.00 %")
       DoEvents
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0) As New propPosMD
End If
LeerPosMDIKOS = mata
End Function

Function DefinirPosMD(toper, cpos, port)
    If port = "PI DISPONIBLE PARA LA VENTA" Then
       DefinirPosMD = ClavePosPIDV
    ElseIf port = "PI CONSERVADOS A VENCIMIENTO" Then
       DefinirPosMD = ClavePosPICV
    ElseIf port = "PI DERIVADOS" Then
       DefinirPosMD = ClavePosPID
    ElseIf port = "FONDOS DE PENSIONES" Then
       DefinirPosMD = ClavePosPenMD
    Else
       If toper = 10 Then
          DefinirPosMD = ClavePosMD
       Else
          DefinirPosMD = cpos
       End If
    End If

End Function

Sub ImpPosMDineroSim(ByVal fecha As Date, ByVal nomarch As String, ByVal txtnompos As String, ByRef noreg As Long, ByVal exito As Boolean)
Dim matposmesa() As Variant
Dim sierrores As Boolean
Dim sihayarch As Boolean
Dim MatVPrecios() As Variant
Dim siposmesad As Boolean
'se importa la posicion total para una fecha
'de la mesa de dinero desde una tabla remota en oracle

    sihayarch = VerifAccesoArch(nomarch)
    If sihayarch Then
       'matposmesa = LeerArchTexto(nomarch, ",", "Leyendo la posicion de la Mesa de Dinero")
       matposmesa = LeerPosMDExcel(nomarch)
       matposmesa = AgregaParamPosPension(matposmesa)
       If UBound(matposmesa, 1) <> 0 Then
          Call GuardaPosMesaDSim(fecha, matposmesa, 2, txtnompos, "000000", noreg)
          exito = True
       End If
    End If
  'se lee la posicion desde la red
End Sub
Function AgregaParamPosPension(ByRef mata() As Variant)
Dim noreg1 As Long
Dim noreg2 As Long
Dim i As Long
Dim j As Long
noreg1 = UBound(mata, 1)
noreg2 = UBound(mata, 2)
ReDim matb(1 To noreg1, 1 To noreg2 + 1) As Variant
For i = 1 To noreg1
   For j = 1 To noreg2
   matb(i, j) = mata(i, j)
   Next j
Next i

For i = 1 To noreg1
   matb(i, 14) = ConvBolStr(DetermSiEmFlujos(mata(i, 5), mata(i, 6), mata(i, 7)))
Next i
AgregaParamPosPension = matb
End Function


Sub GuardaPosMD(ByVal txtfiltro As String, ByRef matpos() As propPosMD, ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horapos As String, ByRef exito As Boolean, ByRef nreg As Long)
Dim noreg As Long
Dim i As Long
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtcadena As String

noreg = UBound(matpos, 1)
If noreg <> 0 Then
ConAdo.Execute "DELETE FROM " & txtfiltro
nreg = 0
For i = 1 To noreg
    If matpos(i).Tipo_Mov <= 4 Then
       txtfecha = "to_date('" & Format(matpos(i).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"  'fecha de registro
       txtfecha1 = "to_date('" & Format(matpos(i).fCompraMD, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfecha2 = "to_date('" & Format(matpos(i).fVencMD, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPosMD & " VALUES("
       txtcadena = txtcadena & tipopos & ","                       'tipo de posicion
       txtcadena = txtcadena & txtfecha & ","                      'fecha DE REGISTRO
       txtcadena = txtcadena & "'" & txtnompos & "',"              'nombre de la posicion
       txtcadena = txtcadena & "'" & horapos & "',"                'hora de la posicion intradia
       txtcadena = txtcadena & "'" & matpos(i).intencion & "',"    'intencion
       txtcadena = txtcadena & matpos(i).C_Posicion & ","          'clave de posicion
       txtcadena = txtcadena & "'" & i & "',"                      'clave de la operacion
       txtcadena = txtcadena & matpos(i).Tipo_Mov & ","            'tipo de operacion
       txtcadena = txtcadena & "'" & matpos(i).tValorMD & "',"     'tv
       txtcadena = txtcadena & "'" & matpos(i).emisionMD & "',"    'emisor
       txtcadena = txtcadena & "'" & matpos(i).serieMD & "',"      'serie
       txtcadena = txtcadena & "'" & matpos(i).cEmisionMD & "',"   'clave emision
       txtcadena = txtcadena & matpos(i).noTitulosMD & ","         'no titulos
       txtcadena = txtcadena & txtfecha1 & ","                     'fecha de compra
       txtcadena = txtcadena & txtfecha2 & ","                     'fecha de vencimiento de operacion
       txtcadena = txtcadena & matpos(i).pAsignadoMD & ","         'precio asignado/COMPRA
       txtcadena = txtcadena & matpos(i).tReporto & ","            'tasa premio
       txtcadena = txtcadena & "null,"                             'subportafolio 1
       txtcadena = txtcadena & "null,"                             'subportafolio 2
       txtcadena = txtcadena & "'" & matpos(i).CalifFP & "',"      'calificacion fp
       txtcadena = txtcadena & "'" & matpos(i).CalifTMD & "',"     'calificacion mo
       txtcadena = txtcadena & "'" & matpos(i).SiFlujosMD & "',"   'si tiene flujos
       txtcadena = txtcadena & "'" & matpos(i).escala & "',"       'escala
       txtcadena = txtcadena & "'" & matpos(i).Sector & "',"       'sector
       txtcadena = txtcadena & matpos(i).recupera & ","            'recuperacion
       txtcadena = txtcadena & matpos(i).CalifLP & ")"             'calif num lp
  
   ConAdo.Execute txtcadena
  nreg = nreg + 1
End If
   AvanceProc = i / noreg
   MensajeProc = "Guardando las operaciones de Mesa de Dinero " & Format(AvanceProc, "##0.00 %")
   DoEvents
Next i
exito = True
MensajeProc = "Se guardaron " & nreg & " registros de la posición de Mesa de Dinero"
Else
 MensajeProc = "Atencion. Faltan registros de la mesa de dinero"
 MsgBox MensajeProc
 Call MostrarMensajeSistema(MensajeProc, frmProgreso.Label2, 1, Date, Time, NomUsuario)
 exito = False
End If
End Sub

Sub GuardaPosFPMD(ByRef mata() As Variant, ByRef matf() As Variant, ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horareg As String, ByVal id_pos As Integer, ByRef nreg As Long, ByRef exito As Boolean)
Dim noreg As Long
Dim i As Long
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtcadena As String
Dim txtborra As String

noreg = UBound(mata, 1)
If noreg <> 0 Then
   For i = 1 To noreg
       If (mata(i, 6) <> "1B" And mata(i, 6) <> "1" And mata(i, 6) <> "1A" And mata(i, 6) <> "1I" And mata(i, 6) <> "51" And mata(i, 6) <> "52" And mata(i, 6) <> "CF" And mata(i, 6) <> "41" And mata(i, 6) <> "FE") Then
          txtfecha1 = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"  'fecha de la posicion
          txtfecha2 = "to_date('" & Format(mata(i, 11), "dd/mm/yyyy") & "','dd/mm/yyyy')" 'fecha de compra
          txtfecha3 = "to_date('" & Format(mata(i, 12), "dd/mm/yyyy") & "','dd/mm/yyyy')" 'fecha de vencimiento
          txtcadena = "INSERT INTO " & TablaPosMD & " VALUES("
          txtcadena = txtcadena & tipopos & ","               'tipo de posicion
          txtcadena = txtcadena & txtfecha1 & ","             'fecha registro
          txtcadena = txtcadena & "'" & txtnompos & "',"      'nombre de la posicion si es simulada
          txtcadena = txtcadena & "'" & horareg & "',"        'hora de la posicion intradia
          txtcadena = txtcadena & "'" & mata(i, 2) & "',"     'intencion
          txtcadena = txtcadena & mata(i, 3) & ","            'clave de posicion
          txtcadena = txtcadena & "'" & mata(i, 4) & "',"     'clave de la operacion
          txtcadena = txtcadena & "'" & mata(i, 5) & "',"     'tipo de operacion
          txtcadena = txtcadena & "'" & mata(i, 6) & "',"     'tv
          txtcadena = txtcadena & "'" & mata(i, 7) & "',"     'emisor
          txtcadena = txtcadena & "'" & mata(i, 8) & "',"     'serie
          txtcadena = txtcadena & "'" & mata(i, 9) & "',"     'clave emision
          txtcadena = txtcadena & mata(i, 10) & ","           'no titulos
          txtcadena = txtcadena & txtfecha2 & ","             'fecha de compra
          txtcadena = txtcadena & txtfecha3 & ","             'fecha de vencimiento de operacion
          txtcadena = txtcadena & mata(i, 13) & ","           'precio asignado/COMPRA
          txtcadena = txtcadena & mata(i, 14) & ","           'tasa premio
          txtcadena = txtcadena & "'" & mata(i, 15) & "',"    'subportafolio 1
          txtcadena = txtcadena & "'" & mata(i, 16) & "',"    'subportafolio 2
          txtcadena = txtcadena & "'" & mata(i, 17) & "',"    'calificacion
          txtcadena = txtcadena & "'" & mata(i, 18) & "')"    'si flujos
          ConAdo.Execute txtcadena
       End If
       nreg = nreg + 1
       AvanceProc = i / noreg
       MensajeProc = "Guardando las operaciones de Mesa de Dinero " & Format(AvanceProc, "##0.00 %")
       DoEvents
    Next i
    If nreg = noreg Then
       exito = True
       MensajeProc = "Se guardaron " & nreg & " registros de la posición del fondo"
    Else
       exito = False
       MensajeProc = "No se guardaron todos los registros"
    End If
Else
 MensajeProc = "No hay registros"
 MsgBox MensajeProc
 exito = False
End If
End Sub

Sub GuardaPosFPDiv(ByRef mata() As Variant, ByRef matf() As Variant, ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horareg As String, ByVal id_pos As Integer, ByRef nreg As Long, ByRef exito As Boolean)
Dim noreg As Long
Dim i As Long
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtcadena As String
Dim txtborra As String

noreg = UBound(mata, 1)
If noreg <> 0 Then
   For i = 1 To noreg
       If (mata(i, 6) = "1B" Or mata(i, 6) = "1" Or mata(i, 6) = "1A" Or mata(i, 6) = "1I" Or mata(i, 6) = "51" Or mata(i, 6) = "52" Or mata(i, 6) = "CF" Or mata(i, 6) = "41" Or mata(i, 6) = "FE") And mata(i, 5) = 1 Then
       txtfecha1 = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"  'fecha de la posicion
       txtfecha2 = "to_date('" & Format(mata(i, 11), "dd/mm/yyyy") & "','dd/mm/yyyy')" 'fecha de compra
       txtfecha3 = "to_date('" & Format(mata(i, 12), "dd/mm/yyyy") & "','dd/mm/yyyy')" 'fecha de vencimiento
       txtcadena = "INSERT INTO " & TablaPosDiv & " VALUES("
       txtcadena = txtcadena & tipopos & ","               'tipo de posicion
       txtcadena = txtcadena & txtfecha1 & ","             'fecha posicion
       txtcadena = txtcadena & "'" & txtnompos & "',"      'nombre de la posicion si es simulada
       txtcadena = txtcadena & "'" & horareg & "',"        'hora de la posicion intradia
       txtcadena = txtcadena & "'" & mata(i, 2) & "',"     'intencion
       txtcadena = txtcadena & mata(i, 3) & ","            'clave de posicion
       txtcadena = txtcadena & "'" & mata(i, 4) & "',"     'clave de la operacion
       txtcadena = txtcadena & "'" & mata(i, 5) & "',"     'tipo de operacion
       txtcadena = txtcadena & "'" & mata(i, 6) & "',"     'tv
       txtcadena = txtcadena & "'" & mata(i, 7) & "',"     'emisor
       txtcadena = txtcadena & "'" & mata(i, 8) & "',"     'serie
       txtcadena = txtcadena & "'" & mata(i, 9) & "',"     'clave emision
       txtcadena = txtcadena & mata(i, 10) & ","           'no titulos
       txtcadena = txtcadena & txtfecha2 & ","             'fecha de compra
       txtcadena = txtcadena & txtfecha3 & ","             'fecha de vencimiento de operacion
       txtcadena = txtcadena & mata(i, 13) & ","           'precio asignado/COMPRA
       txtcadena = txtcadena & "'" & mata(i, 15) & "',"    'subportafolio 1
       txtcadena = txtcadena & "'" & mata(i, 16) & "',"    'subportafolio 2
       txtcadena = txtcadena & "'" & mata(i, 17) & "')"    'calificacion
       ConAdo.Execute txtcadena
       End If
       nreg = nreg + 1
       AvanceProc = i / noreg
       MensajeProc = "Guardando las operaciones de Mesa de Dinero " & Format(AvanceProc, "##0.00 %")
       DoEvents
    Next i
    If nreg = noreg Then
       exito = True
       MensajeProc = "Se guardaron " & nreg & " registros de la posición del fondo"
    Else
       exito = False
       MensajeProc = "No se guardaron todos los registros"
    End If
Else
 MensajeProc = "No hay registros"
 MsgBox MensajeProc
 exito = False
End If
End Sub

Sub GuardaPosMesaDSim(ByVal fecha As Date, ByRef mata() As Variant, ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horareg As String, ByRef noreg As Long)
Dim i As Long
Dim contar As Long
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtcadena As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

noreg = UBound(mata, 1)
If noreg <> 0 Then
ConAdo.Execute "DELETE FROM " & TablaPosMD & " WHERE NOMPOS = '" & txtnompos & "'"
contar = 0
For i = 1 To noreg
    If mata(i, 1) = "N" Then
       contar = contar + 1
       txtfecha1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfecha2 = "to_date('" & Format(mata(i, 10), "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfecha3 = "to_date('" & Format(mata(i, 13), "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaPosMD & " VALUES("
       txtcadena = txtcadena & tipopos & ","                        'TIPO DE POSICION
       txtcadena = txtcadena & txtfecha1 & ","                      'FECHA
       txtcadena = txtcadena & "'" & txtnompos & "',"               'nombre de la posicion
       txtcadena = txtcadena & "'" & horareg & "',"                 'HORA DE LA POSICION
       txtcadena = txtcadena & "'" & mata(i, 1) & "',"              'intencion
       txtcadena = txtcadena & mata(i, 2) & ","                     'CLAVE DE LA POSICION
       txtcadena = txtcadena & "'" & mata(i, 3) & "',"              'CLAVE DE LA OPERACION
       txtcadena = txtcadena & mata(i, 4) & ","                     'tipo de operacion
       txtcadena = txtcadena & "'" & mata(i, 5) & "',"              'tv
       txtcadena = txtcadena & "'" & mata(i, 6) & "',"              'emision
       txtcadena = txtcadena & "'" & mata(i, 7) & "',"              'serie
       txtcadena = txtcadena & "'" & mata(i, 8) & "',"              'clave emision
       txtcadena = txtcadena & mata(i, 9) & ","                     'no titulos
       txtcadena = txtcadena & txtfecha2 & ","                      'fecha valor/compra
       txtcadena = txtcadena & txtfecha3 & ","                      'fecha de vencimiento de la operacion
       txtcadena = txtcadena & mata(i, 11) & ","                    'PRECIO COMPRA
       txtcadena = txtcadena & mata(i, 12) & ","                    'tasa premio
       txtcadena = txtcadena & "null,"                              'subportafolio1
       txtcadena = txtcadena & "null,"                              'subportafolio2
       txtcadena = txtcadena & "null,"                              'calificacion
       txtcadena = txtcadena & "'" & ConvBolStr(mata(i, 14)) & "')"       'si flujos
       ConAdo.Execute txtcadena
    End If
  AvanceProc = i / noreg
  MensajeProc = "Guardando las operaciones de Mesa de Dinero " & Format(AvanceProc, "###0.00")
 DoEvents
Next i
MensajeProc = "Se guardaron " & noreg & " registros de la posición de Mesa de Dinero del "
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0

End Sub

Function LeerVPrecios(ByVal fecha As Date, ByRef mindvp() As Variant) As propVecPrecios()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

Dim noprecios As Long
Dim txtfecha As String
Dim noreg As Integer
Dim i As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim mata() As New propVecPrecios
Dim rprecios As New ADODB.recordset


txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaVecPrecios
txtfiltro2 = txtfiltro2 & " WHERE FECHA = " & txtfecha & " ORDER BY CLAVE_EMISION ASC"
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"
rprecios.Open txtfiltro1, ConAdo
noreg = rprecios.Fields(0)
rprecios.Close
If noreg <> 0 Then
ReDim mata(1 To noreg)
ReDim mindvp(1 To noreg, 1 To 2)
rprecios.Open txtfiltro2, ConAdo
rprecios.MoveFirst
For i = 1 To noreg
    mindvp(i, 1) = ReemplazaVacioValor(rprecios.Fields("clave_emision"), "")
    mindvp(i, 2) = i
    mata(i).fecha = rprecios.Fields("fecha")
    mata(i).tmercado = ReemplazaVacioValor(rprecios.Fields("tmercado"), 0)
    mata(i).tv = ReemplazaVacioValor(rprecios.Fields("TV"), "")
    mata(i).emision = ReemplazaVacioValor(rprecios.Fields("EMISION"), "")
    mata(i).serie = ReemplazaVacioValor(rprecios.Fields("SERIE"), "")
    mata(i).psucio = ReemplazaVacioValor(rprecios.Fields("PSUCIO"), 0)
    mata(i).plimpio = ReemplazaVacioValor(rprecios.Fields("PLIMPIO"), 0)
    mata(i).int_md = ReemplazaVacioValor(rprecios.Fields("INTERESESMD"), 0)
    mata(i).tasa_st = ReemplazaVacioValor(rprecios.Fields("tasa_sobret"), 0)
    mata(i).dxv = ReemplazaVacioValor(rprecios.Fields("dvencimiento"), 0)
    mata(i).vnominal = ReemplazaVacioValor(rprecios.Fields("vnominal"), 0)
    mata(i).fvenc = ReemplazaVacioValor(rprecios.Fields("fvencimiento"), 0)
    mata(i).femision = ReemplazaVacioValor(rprecios.Fields("femision"), 0)
    mata(i).moneda = ReemplazaVacioValor(rprecios.Fields("moneda"), 0)
    mata(i).tcupon = ReemplazaVacioValor(rprecios.Fields("tcupon"), 0)
    mata(i).pcupon = ReemplazaVacioValor(rprecios.Fields("pcupon"), 0)
    mata(i).yield = ReemplazaVacioValor(rprecios.Fields("yield"), 0)
    mata(i).calif_sp = ReemplazaVacioValor(rprecios.Fields("calif_sp"), 0)
    mata(i).calif_fitch = ReemplazaVacioValor(rprecios.Fields("calif_fitch"), 0)
    mata(i).calif_moodys = ReemplazaVacioValor(rprecios.Fields("calif_moodys"), 0)
    mata(i).calif_hr = ReemplazaVacioValor(rprecios.Fields("calif_hr"), 0)
    mata(i).c_emision = ReemplazaVacioValor(rprecios.Fields("clave_emision"), "")
    mata(i).regla_cupon = ReemplazaVacioValor(rprecios.Fields("REGLA_CUPON"), "")
    mata(i).frec_cupon = ReemplazaVacioValor(rprecios.Fields("FREC_CUPON"), "")
    mata(i).st_colocacion = ReemplazaVacioValor(rprecios.Fields("ST_COLOCACION"), 0)
    rprecios.MoveNext  'se lee el siguiente registro
    AvanceProc = i / noreg
    MensajeProc = "Leyendo el vector de precios: " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
rprecios.Close
mindvp = RutinaOrden(mindvp, 1, SRutOrden)
Else
   MensajeProc = "No hay precios para el dia " & fecha
 ReDim mata(0 To 0)
End If
LeerVPrecios = mata

On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerPVPrecios(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfiltro As String
Dim sql_precios As String
Dim sql_num_precios As String
Dim noprecios As Long
Dim i As Long
Dim noreg As Integer
Dim rprecios As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = TablaVecPrecios & " WHERE FECHA = " & txtfecha
sql_num_precios = "select count(*) from " & txtfiltro
rprecios.Open sql_num_precios, ConAdo
noreg = rprecios.Fields(0)
rprecios.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 15) As Variant
sql_precios = "select * from " & txtfiltro
rprecios.Open sql_precios, ConAdo
rprecios.MoveFirst
For i = 1 To noreg
    mata(i, 1) = rprecios.Fields("fecha")            'fecha
    mata(i, 2) = ReemplazaVacioValor(rprecios.Fields("psucio"), 0)
    mata(i, 3) = ReemplazaVacioValor(rprecios.Fields("plimpio"), 0)
    mata(i, 4) = ReemplazaVacioValor(rprecios.Fields("tcupon"), 0) / 100
    mata(i, 5) = ReemplazaVacioValor(rprecios.Fields("pcupon"), 0)
    mata(i, 6) = ReemplazaVacioValor(rprecios.Fields("clave_emision"), "")
    mata(i, 7) = ReemplazaVacioValor(rprecios.Fields("vnominal"), 0)
    mata(i, 8) = ReemplazaVacioValor(rprecios.Fields("fvencimiento"), 0)
    mata(i, 9) = ReemplazaVacioValor(rprecios.Fields("INTERESESMD"), 0)
    mata(i, 10) = ReemplazaVacioValor(rprecios.Fields("CALIF_SP"), 0)
    mata(i, 11) = ReemplazaVacioValor(rprecios.Fields("CALIF_FITCH"), 0)
    mata(i, 12) = ReemplazaVacioValor(rprecios.Fields("CALIF_MOODYS"), 0)
    mata(i, 13) = ReemplazaVacioValor(rprecios.Fields("CALIF_HR"), 0)
    mata(i, 14) = ReemplazaVacioValor(rprecios.Fields("REGLA_CUPON"), "")
    mata(i, 15) = ReemplazaVacioValor(rprecios.Fields("TV"), "")
    rprecios.MoveNext  'se lee el siguiente registro
AvanceProc = i / noreg
MensajeProc = "Leyendo el vector de precios del " & fecha & " : " & Format(AvanceProc, "#,##0.00 %")

DoEvents
Next i
rprecios.Close
mata = RutinaOrden(mata, 6, SRutOrden)
Else
 MensajeProc = "No hay precios para el dia " & fecha
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
'se ordena por la clave de emision
If UBound(mata, 1) <> 0 Then FechaVPrecios = fecha
LeerPVPrecios = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0

End Function

Function TPosPensiones(ByRef mat() As Variant) As Variant()
Dim noreg As Long
Dim i As Long
Dim n As Long

'esta rutina es para transformar los datos que
'se obtengan de la lectura de una tabla de datos
'del fondo de pensiones
noreg = UBound(mat, 1)
If noreg <> 0 Then
 n = UBound(mat, 1)
 ReDim mat1(1 To n) As Variant
 For i = 1 To n
 mat1(i).Tipo_Mov = 1   'tipo fr operacion todas compra
 mat1(i).Signo_Op = TraducirTMov(mat(i).Tipo_Mov)
 mat1(i).cEmisionMD = mat(i).cEmisionMD 'tipo de operacion
 mat1(i).noTitulosMD = mat(i).noTitulosMD 'tipo de operacion
 mat1(i).fVencMD = mat(i).fVencMD 'tipo de operacion
 mat1(i).pAsignadoMD = mat(i).pAsignadoMD       'tipo de operacion
 mat1(i).tReporto = mat(i).tReporto      'fecha de
 mat1(i).fCompraMD = mat(i).fCompraMD      'fecha de
 mat1(i).vNominalMD = mat(i).vNominalMD     'tipo de operacion
  If Not IsNull(mat(i, CDescTitulo)) Then
   mat1(i, CDescTitulo) = mat(i, CDescTitulo)   '
  Else
   mat1(i, CDescTitulo) = 0
  End If
 Next i
 TPosPensiones = mat1
End If
End Function

Function TPosDerivados(ByRef mat() As Variant, ByVal noreg As Long)
Dim n As Long
Dim i As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta rutina es para transformar los datos que
'se obtengan de la lectura de una tabla de datos
'de la posicion de derivados
If noreg <> 0 Then
 n = UBound(mat, 1)
 ReDim mat1(1 To n) As Variant
 For i = 1 To n
 mat1(i).Tipo_Mov = mat(i).Tipo_Mov   'tipo fr operacion todas compra
 mat1(i).Signo_Op = TraducirTMov(mat1(i).Tipo_Mov)
 mat1(i).cEmisionMD = mat(i).cEmisionMD 'tipo de operacion
 mat1(i).noTitulosMD = mat(i).noTitulosMD 'tipo de operacion
 mat1(i).fVencMD = mat(i).fVencMD 'tipo de operacion
 mat1(i).pAsignadoMD = mat(i).pAsignadoMD      'precio de compra
 mat1(i).tReporto = mat(i).tReporto      'tasa de compra
 mat1(i).tCuponMD = mat(i).tCuponMD      '
 mat1(i).fCompraMD = mat(i).fCompraMD      'fecha de compra
 mat1(i).vNominalMD = mat(i).vNominalMD     '
 If Not IsNull(mat(i, CDescTitulo)) Then
  mat1(i, CDescTitulo) = mat(i, CDescTitulo)   '
 Else
  mat1(i, CDescTitulo) = 0
 End If
 
 Next i
 TPosDerivados = mat1
End If


On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerPosPensiones(ByVal txtport As String, ByVal txtmsg As String)
Dim sql_num_mesa As String
Dim sql_mesa As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

'esta rutina sirve para leer la posicion
'del fondo de pensiones de la tabla de oracle
sql_num_mesa = "select count(*) from " & txtport

rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close

If noreg <> 0 Then
ReDim mata(1 To noreg) As New propPosMD
sql_mesa = "select * from " & txtport
rmesa.Open sql_mesa, ConAdo
rmesa.MoveFirst

For i = 1 To noreg
mata(i).C_Posicion = 12
mata(i).fechareg = rmesa.Fields("FECHA")
If IsNull(rmesa.Fields("toperacion")) Then
 mata(i).Tipo_Mov = 0
Else
 mata(i).Tipo_Mov = rmesa.Fields("toperacion")
End If
mata(i).Signo_Op = TraducirTMov(mata(i).Tipo_Mov)
If IsNull(Trim(rmesa.Fields("tv"))) Then
 mata(i).tValorMD = "Desconocido"
Else
  mata(i).tValorMD = Trim(rmesa.Fields("tv"))
End If
If IsNull(Trim(rmesa.Fields("emision"))) Then
 mata(i).emisionMD = "Desconocido"
Else
  mata(i).emisionMD = Trim(rmesa.Fields("emision"))
End If
If IsNull(Trim(rmesa.Fields("serie"))) Then
 mata(i).serieMD = "Desconocido"
Else
  mata(i).serieMD = Trim(rmesa.Fields("serie"))
End If
If IsNull(Trim(rmesa.Fields("cemision"))) Then
 mata(i).cEmisionMD = "Desconocido"
Else
  mata(i).cEmisionMD = Trim(rmesa.Fields("cemision"))
End If
If IsNull(rmesa.Fields("ntitulos")) Then
  mata(i).noTitulosMD = 0
Else
  mata(i).noTitulosMD = Val(rmesa.Fields("ntitulos"))
End If
If IsNull(rmesa.Fields("fvencimiento")) Then
  mata(i).fVencMD = 0
Else
  mata(i).fVencMD = CDate(rmesa.Fields("fvencimiento"))
End If
If IsNull(rmesa.Fields("VNominal")) Then
 mata(i).vNominalMD = 0
Else
 mata(i).vNominalMD = Val(Trim(rmesa.Fields("VNominal")))
End If
If IsNull(rmesa.Fields("pcupon")) Then
 mata(i).pCuponMD = 0
Else
 mata(i).pCuponMD = Val(Trim(rmesa.Fields("pcupon")))
End If
If IsNull(rmesa.Fields("tcupon")) Then
  mata(i).tCuponMD = 0
Else
  mata(i).tCuponMD = Val(rmesa.Fields("tcupon"))
End If
If IsNull(rmesa.Fields("tpremio")) Then
  mata(i).tReporto = 0
Else
  mata(i).tReporto = Val(rmesa.Fields("tpremio"))
End If
If IsNull(rmesa.Fields("pcompra")) Then
  mata(i).pAsignadoMD = 0
Else
  mata(i).pAsignadoMD = Val(rmesa.Fields("pcompra"))
End If


  rmesa.MoveNext  'se lee el siguiente registro
  AvanceProc = i / noreg
  MensajeProc = txtmsg & ": " & Format(AvanceProc, "##0.00 %")
Next i
 Call MostrarMensajeSistema("Se leyeron " & noreg & " de la posicion del fondo de pensiones", frmProgreso.Label2, 0, Date, Time, NomUsuario)
 rmesa.Close
Else
 ReDim mata(0 To 0) As New propPosMD
End If
LeerPosPensiones = mata
End Function

Function AnexarClaveEmision(ByRef matpos() As Variant, ByVal ind1 As Integer, ByVal ind2 As Integer, ByVal ind3 As Integer, ByVal ind4 As Integer)
Dim noreg As Long
Dim i As Long
'ind1 posicion o mesa
'ind2 tipo mercado
'ind3 tipo valor
'ind4 emision
'ind5 serie
'ind6 columna con el resultado
'en un segundo proceso se colocan algunas claves necesarias para
'este es un proceso heuristico que debe ser mejorado
noreg = UBound(matpos, 1)
For i = 1 To noreg
    matpos(i, ind4) = GeneraClaveEmision(matpos(i, ind1), matpos(i, ind2), matpos(i, ind3))
    AvanceProc = i / noreg
    MensajeProc = "generando la clave de la emision " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
AnexarClaveEmision = matpos

End Function

Function GeneraClaveEmision(ByVal texto1 As String, ByVal texto2 As String, ByVal texto3 As String) As String
Dim noreg2 As Long
Dim indice As Long
Dim maxvars As Integer
Dim j As Long
Dim nocond As Integer
Dim valor As Boolean
Dim sicond1 As Boolean
Dim sicond2 As Boolean
Dim sicond3 As Boolean
Dim clave As Integer
Dim nomemision As String

'ind1 posicion o mesa
'ind2 tipo mercado
'ind3 tipo valor
'ind4 emision
'ind5 serie
'ind6 columna con el resultado

'en un segundo proceso se colocan algunas claves necesarias para
'este es un proceso heuristico que debe ser mejorado
 noreg2 = UBound(MatClavesEmision, 1)
 indice = 0
 maxvars = 0
 For j = 1 To noreg2
 nocond = 0
 valor = True

 If Not EsVariableVacia(texto1) Then                    'primero se clasifica por tv
  If MatClavesEmision(j, 1) <> "*" Then
      If texto1 = MatClavesEmision(j, 1) Then
       sicond1 = True
       nocond = nocond + 1
      Else
       sicond1 = False
      End If
  Else
   sicond1 = True
  End If
  If sicond1 = True And valor Then
     valor = True
  Else
     valor = False
  End If
 End If
 If Not EsVariableVacia(texto2) Then                    'emision
  If MatClavesEmision(j, 2) <> "*" Then
      If texto2 = MatClavesEmision(j, 2) Then
       sicond2 = True
       nocond = nocond + 1
      Else
       sicond2 = False
      End If
  Else
   sicond2 = True
  End If
 If sicond2 = True And valor Then
    valor = True
 Else
    valor = False
 End If
 End If
 If Not EsVariableVacia(texto3) Then                    'serie
  If MatClavesEmision(j, 3) <> "*" Then
   If texto3 = MatClavesEmision(j, 3) Then
    sicond3 = True
    nocond = nocond + 1
   Else
    sicond3 = False
   End If
  Else
   sicond3 = True
  End If
  If sicond3 = True And valor Then
     valor = True
  Else
     valor = False
  End If
  End If
' If not esvariablevacia(texto4) Then                    'serie
'  If MatClavesEmision(j, 4) <> "*" Then
'      If texto4 = MatClavesEmision(j, 4) Then
'       sicond4 =true
'       nocond = nocond + 1
'      Else
'       sicond4 = false
'      End If
'  Else
'    sicond4 =true
'  End If
'  valor = valor And sicond4
' End If
 'If not esvariablevacia(texto5) Then                   'serie
'  If MatClavesEmision(j, 5) <> "*" Then
      'If texto5 = MatClavesEmision(j, 5) Then
'       sicond5 =true
       'nocond = nocond + 1
      'Else
'       sicond5 = false
      'End If
  'Else
    'sicond5 =true
  'End If
  'valor = valor And sicond5
 'End If
  If valor = True And nocond > 0 Then
   If nocond > maxvars Then indice = j
   maxvars = Maximo(maxvars, nocond)
  End If
 Next j
'se selecciona el registro con mayor numero de coincidencias
If indice <> 0 Then
   clave = MatClavesEmision(indice, 4)
   If clave = 1 Then
      nomemision = texto1 & texto2 & texto3
   ElseIf clave = 2 Then
      nomemision = texto1 & texto3
   ElseIf clave = 3 Then
      nomemision = texto1 & texto3
   Else
      nomemision = texto1 & texto2 & texto3
   End If
Else
'en cualquier otro caso se ejecuta la regla tipovalor-emision-serie
   nomemision = texto1 & texto2 & texto3
End If
   GeneraClaveEmision = nomemision
End Function

Function EncuentraEmisionBase(ByVal texto1 As String, ByVal texto2 As String, ByVal texto3 As String)
Dim noreg2 As Integer
Dim indice As Integer
Dim maxvars As Integer
Dim j As Integer
Dim nocond As Integer
Dim valor As Boolean
Dim sicond1 As Boolean
Dim sicond2 As Boolean
Dim sicond3 As Boolean
'ind1 posicion o mesa
'ind2 tipo mercado
'ind3 tipo valor
'ind4 emision
'ind5 serie
'ind6 columna con el resultado

'en un segundo proceso se colocan algunas claves necesarias para
'este es un proceso heuristico que debe ser mejorado
 noreg2 = UBound(MatParamEmisiones, 1)
 indice = 0
 maxvars = 0
 For j = 1 To noreg2
 nocond = 0
 valor = True

 If Not EsVariableVacia(texto1) Then              'tipo valor
  If MatParamEmisiones(j, 1) <> "*" Then
      If texto1 = MatParamEmisiones(j, 1) Then
       sicond1 = True
       nocond = nocond + 1
      Else
       sicond1 = False
      End If
  Else
   sicond1 = True
  End If
  If valor = True And sicond1 Then
     valor = True
  Else
     valor = False
  End If
 End If
 If Not EsVariableVacia(texto2) Then               'emision
  If MatParamEmisiones(j, 2) <> "*" Then
      If texto2 = MatParamEmisiones(j, 2) Then
       sicond2 = True
       nocond = nocond + 1
      Else
       sicond2 = False
      End If
  Else
   sicond2 = True
  End If
  If valor = True And sicond2 Then
  valor = True
  Else
  valor = False
  End If
 End If
 If Not EsVariableVacia(texto3) Then               'serie
  If MatParamEmisiones(j, 3) <> "*" Then
   If texto3 = MatParamEmisiones(j, 3) Then
    sicond3 = True
    nocond = nocond + 1
   Else
    sicond3 = False
   End If
  Else
   sicond3 = True
  End If
  If valor = True And sicond3 Then
     valor = True
  Else
     valor = False
  End If
  End If
  If valor = True And nocond > 0 Then
   If nocond > maxvars Then indice = j
   maxvars = Maximo(maxvars, nocond)
  End If
 Next j
'se selecciona el registro con mayor numero de coincidencias
 EncuentraEmisionBase = indice

End Function

Function CargaPortReporteCVAR(ByVal txtport As String)
'se cargan los grupos del portafolio que se quiere desglosar

Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
'====================================================
txtfiltro2 = "select * from " & PrefijoBD & TablaReporteCVaR & " WHERE REPORTE = '" & txtport & "' ORDER BY ORDEN"
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 ReDim mata(1 To noreg, 1 To 4) As Variant
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("REPORTE")            'reporte
       mata(i, 2) = rmesa.Fields("ORDEN")              'orden
       mata(i, 3) = rmesa.Fields("CONCEPTO")           'concepto
       mata(i, 4) = rmesa.Fields("DESCRIPCION")        'var 100%
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Cargando las estructuras de portafolio " & Format(AvanceProc, "##0.00 %")
   Next i
   rmesa.Close
Else
  ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaPortReporteCVAR = mata
End Function

Function DefinePortSegRiesgo() As String()
ReDim mata(1 To 12, 1 To 2) As String
mata(1, 1) = "CONSOLIDADO"
mata(2, 1) = "MERCADO DE DINERO"
mata(3, 1) = "MESA DE DINERO"
mata(4, 1) = "TESORERIA"
mata(5, 1) = "MESA DE CAMBIOS"
mata(6, 1) = "DERIVADOS DE NEGOCIACION"
mata(7, 1) = "DERIVADOS ESTRUCTURALES"
mata(8, 1) = "DERIVADOS NEGOCIACION RECLASIFICACION"
mata(9, 1) = "PORTAFOLIO DE INVERSION"
mata(10, 1) = "PI DISPONIBLES PARA LA VENTA"
mata(11, 1) = "PI CONSERVADOS A VENCIMIENTO"
mata(12, 1) = "PI DERIVADOS"

mata(1, 2) = "BANOBRAS"
mata(2, 2) = "DINEROB"
mata(3, 2) = "MDINERO"
mata(4, 2) = "TESO"
mata(5, 2) = "CAMBIOS"
mata(6, 2) = "DERIV"
mata(7, 2) = "DERIVEST"
mata(8, 2) = "DERIVR"
mata(9, 2) = "PI"
mata(10, 2) = "PIDV"
mata(11, 2) = "PICV"
mata(12, 2) = "PID"
DefinePortSegRiesgo = mata
End Function

Function CargaGruposPortPos(ByVal txtport As String)
'se cargan los grupos del portafolio que se quiere desglosar
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim rmesa As New ADODB.recordset
'====================================================
If txtport = "*" Then
   txtfiltro2 = "SELECT * FROM " & PrefijoBD & TablaGruposPortPos & " ORDER BY CONCEPTO"
Else
   txtfiltro2 = "select * from " & PrefijoBD & TablaGruposPortPos & " WHERE GRUPO = '" & txtport & "' ORDER BY ORDEN"
End If
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
NoGruposPort = rmesa.Fields(0)
rmesa.Close
If NoGruposPort <> 0 Then
 ReDim mata(1 To NoGruposPort, 1 To 5) As Variant
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To NoGruposPort
       mata(i, 1) = rmesa.Fields("GRUPO")           'reporte
       mata(i, 2) = rmesa.Fields("ORDEN")           'orden
       mata(i, 3) = rmesa.Fields("CONCEPTO")        'concepto
       mata(i, 4) = rmesa.Fields("VAR100")          'var 100%
       rmesa.MoveNext
       AvanceProc = i / NoGruposPort
       MensajeProc = "Cargando las estructuras de portafolio " & Format(AvanceProc, "##0.00 %")
   Next i
   rmesa.Close
Else
  ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaGruposPortPos = mata
End Function

Sub LeerPortafolioFRiesgo(ByVal txtnomportfr As String, ByRef matnodosfr() As propNodosFRiesgo, ByRef nofact As Long)
Dim txtfiltro As String
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim i As Integer
Dim j As Integer
Dim indice As Integer
Dim rmesa As New ADODB.recordset

'matndfr tiene los datos desglosados
'MatResFRiesgo tiene los datos por grupo


'se leen los plazos de las curvas que componen este
'grupo o portafolio de curvas
'aqui se determina el orden que los factores tendran en la tabla de factores
'1   NOMBRE del factor
'2   NO DE NODOS
'3   INDICE DE FACTORES ACUMULADOS
'4   TIPO VALOR
'    SELECT CONCEPTO, Count(CONCEPTO) AS CuentaDeCURVA, first([TIPO VALOR]) as TipoValor,first([TIPO FUNCION]) as TipoFuncion From [PORTAFOLIOS FACTORES R] WHERE PORTAFOLIO = '" & txtnomportfr & "' GROUP BY CONCEPTO ORDER BY  CONCEPTO"
'====================================================
'se determinan el numero de curvas
txtfiltro = "SELECT COUNT(distinct CONCEPTO) From " & PrefijoBD & TablaPortFR & " Where PORTAFOLIO = '" & txtnomportfr & "'"
rmesa.Open txtfiltro, ConAdo
NoGruposFR = rmesa.Fields(0)
rmesa.Close
If NoGruposFR <> 0 Then
   txtfiltro = "SELECT CONCEPTO, Count(CONCEPTO) AS CuentaDeCONCEPTO, MIN(TIPO_VALOR) AS PrimeroDeTIPO_VALOR, MIN(TIPO_FUNCION) AS PrimeroDeTIPO_FUNCION"
   txtfiltro = txtfiltro & " From " & PrefijoBD & TablaPortFR & " Where PORTAFOLIO = '" & txtnomportfr & "' GROUP BY CONCEPTO ORDER BY CONCEPTO"
   rmesa.Open txtfiltro, ConAdo
'ESTRUCTURA DE matresfr1iesgo
'1  nombre de la curva
'2  no de plazos
'3  total acumulado de plazos
'4  tipo factor
'5  tipo funcion interpolacion aplicar
ReDim MatResFRiesgo(1 To NoGruposFR) As New resPropFRiesgo
ReDim MatResFRiesgo1(1 To NoGruposFR, 1 To 5) As Variant
   nofact = 0
   For i = 1 To NoGruposFR
       MatResFRiesgo1(i, 1) = rmesa.Fields("CONCEPTO")
       MatResFRiesgo1(i, 2) = rmesa.Fields("CuentaDeCONCEPTO")
       MatResFRiesgo1(i, 4) = rmesa.Fields("PrimeroDeTIPO_VALOR")
       MatResFRiesgo1(i, 5) = rmesa.Fields("PrimeroDeTIPO_FUNCION")
       rmesa.MoveNext
       nofact = nofact + MatResFRiesgo1(i, 2)
   Next i
   rmesa.Close
   MatResFRiesgo1 = RutinaOrden(MatResFRiesgo1, 1, SRutOrden)
   For i = 1 To NoGruposFR
       MatResFRiesgo(i).nomFactor = MatResFRiesgo1(i, 1)
       MatResFRiesgo(i).nonodos = MatResFRiesgo1(i, 2)
       If i = 1 Then
          MatResFRiesgo(i).nonodosacum = MatResFRiesgo1(i, 2)
       Else
          MatResFRiesgo(i).nonodosacum = MatResFRiesgo(i - 1).nonodosacum + MatResFRiesgo1(i, 2)
       End If
       MatResFRiesgo(i).tfactor = MatResFRiesgo1(i, 4)
       MatResFRiesgo(i).escurva = MatResFRiesgo1(i, 5)
   Next i

'indica que deben incluirse todos los factores de riesgo en la matriz de factores
ReDim SiFactorRiesgo(1 To nofact) As Boolean
   For i = 1 To nofact
       SiFactorRiesgo(i) = True
   Next i
   noreg1 = 0
   For i = 1 To NoGruposFR
       noreg1 = Maximo(noreg1, MatResFRiesgo(i).nonodos)
   Next i
'se crea la una matriz que va a tener los plazos de todas las curvas
'primero se ordena por factor y luego por plazo

ReDim MatPlazos(1 To noreg1, 1 To NoGruposFR) As Long
ReDim MatDescripFR(1 To noreg1, 1 To NoGruposFR) As Variant
   For i = 1 To NoGruposFR
       txtfiltro = "SELECT count(*) FROM " & PrefijoBD & TablaPortFR & " WHERE CONCEPTO = '" & MatResFRiesgo(i).nomFactor & "' AND PORTAFOLIO = '" & txtnomportfr & "' ORDER BY PLAZO"
       rmesa.Open txtfiltro, ConAdo
       noreg2 = rmesa.Fields(0)
       rmesa.Close
       If noreg2 <> 0 Then
          txtfiltro = "SELECT * FROM " & PrefijoBD & TablaPortFR & " WHERE CONCEPTO = '" & MatResFRiesgo(i).nomFactor & "' AND PORTAFOLIO = '" & txtnomportfr & "' ORDER BY PLAZO"
          rmesa.Open txtfiltro, ConAdo
          For j = 1 To noreg2
              MatPlazos(j, i) = rmesa.Fields("PLAZO")
              MatDescripFR(j, i) = rmesa.Fields("DESCRIPCION")     'descripcion larga
              rmesa.MoveNext
          Next j
          rmesa.Close
       End If
       AvanceProc = i / NoGruposFR
       MensajeProc = "Leyendo la estructura de factores de riesgo " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
ReDim matndfr(1 To nofact) As New propNodosFRiesgo
For i = 1 To NoGruposFR
    For j = 1 To MatResFRiesgo(i).nonodos  'no de nodos para la curva i
' se define el nombre para el factor de riesgo
        If i = 1 Then
           indice = 0
        Else
           indice = MatResFRiesgo(i - 1).nonodosacum
        End If
        matndfr(indice + j).indFactor = MatResFRiesgo(i).nomFactor & " " & MatPlazos(j, i)    'clave de indexacion
        matndfr(indice + j).tfactor = MatResFRiesgo(i).tfactor                                'tipo de factor (tasa,sobretasa, indice)
        matndfr(indice + j).nomFactor = MatResFRiesgo(i).nomFactor                            'nombre de la curva o factor
        matndfr(indice + j).plazo = MatPlazos(j, i)                                       'plazo factor
        matndfr(indice + j).descFactor = MatDescripFR(j, i)                               'descripcion larga
        If j <> 1 Then
            If MatPlazos(j - 1, i) > MatPlazos(j, i) Then MsgBox "Algo esta mal"
        End If
    Next j
    If NoGruposFR <> 0 Then
       AvanceProc = i / NoGruposFR
    Else
       AvanceProc = 0
    End If
    MensajeProc = "Leyendo el portafolio de factores de riesgo " & Format(AvanceProc, "#,##0.00 %")
    DoEvents
Next i
matnodosfr = matndfr
End If
End Sub

Function BuscarTProd(ByVal i As Long, ByRef matpos() As propPosMD, ByRef matclas() As Variant)
'busca el mayor numero de coincidencias para el tipo valor,emsision y serie del instrumento
'en la tabla matclas(), si encuentra al menos una coincidencia, le asigna un indice
'distinto de cero
Dim noport As Long
Dim indice As Long
Dim maxvars As Long
Dim j As Long
Dim largo As Long

 noport = UBound(matclas, 1)
 indice = 0
 maxvars = 0
 For j = 1 To noport
     largo = 0
     If matclas(j, 1) <> "*" And matclas(j, 2) <> "*" And matclas(j, 3) <> "*" Then
       If matclas(j, 1) = matpos(i).tValorMD And matclas(j, 2) = matpos(i).emisionMD And matclas(j, 3) = matpos(i).serieMD Then
         largo = largo + 1
       Else
         largo = 0
       End If
     ElseIf matclas(j, 1) <> "*" And matclas(j, 2) <> "*" And matclas(j, 3) = "*" Then
        If matclas(j, 1) = matpos(i).tValorMD And matclas(j, 2) = matpos(i).emisionMD Then
           largo = largo + 1
        Else
           largo = 0
        End If
     ElseIf matclas(j, 1) <> "*" And matclas(j, 2) = "*" And matclas(j, 3) <> "*" Then
        If matclas(j, 1) = matpos(i).tValorMD And matclas(j, 3) = matpos(i).serieMD Then
           largo = largo + 1
        Else
           largo = 0
        End If
     ElseIf matclas(j, 1) = "*" And matclas(j, 2) <> "*" And matclas(j, 3) <> "*" Then
        If matclas(j, 2) = matpos(i).emisionMD And matclas(j, 3) = matpos(i).serieMD Then
           largo = largo + 1
        Else
           largo = 0
        End If
     ElseIf matclas(j, 1) <> "*" And matclas(j, 2) = "*" And matclas(j, 3) = "*" Then
        If matclas(j, 1) = matpos(i).tValorMD Then
           largo = largo + 1
        Else
           largo = 0
        End If
     ElseIf matclas(j, 1) = "*" And matclas(j, 2) <> "*" And matclas(j, 3) = "*" Then
        If matclas(j, 2) = matpos(i).emisionMD Then
           largo = largo + 1
        Else
           largo = 0
        End If
     ElseIf matclas(j, 1) = "*" And matclas(j, 2) = "*" And matclas(j, 3) <> "*" Then
        If matclas(j, 3) = matpos(i).serieMD Then
           largo = largo + 1
        Else
           largo = 0
        End If
     End If
     If largo > maxvars Then
        indice = j
        maxvars = Maximo(maxvars, largo)
     End If
 Next j
BuscarTProd = indice
End Function

Function BuscarTProd2(ByVal i As Long, ByRef matpos() As propPosDiv, ByRef matclas() As Variant)
Dim noport As Long
Dim indice As Long
Dim maxvars As Long
Dim j As Long
Dim largo As Long

 noport = UBound(matclas, 1)
 indice = 0
 maxvars = 0
 For j = 1 To noport
     largo = 0
     If matclas(j, 1) <> "*" And matclas(j, 2) <> "*" And matclas(j, 3) <> "*" Then
       If matclas(j, 1) = matpos(i).TValorDiv And matclas(j, 2) = matpos(i).EmisionDiv And matclas(j, 3) = matpos(i).SerieDiv Then
         largo = largo + 1
       Else
         largo = 0
       End If
     ElseIf matclas(j, 1) <> "*" And matclas(j, 2) <> "*" And matclas(j, 3) = "*" Then
        If matclas(j, 1) = matpos(i).TValorDiv And matclas(j, 2) = matpos(i).EmisionDiv Then
           largo = largo + 1
        Else
           largo = 0
        End If
     ElseIf matclas(j, 1) <> "*" And matclas(j, 2) = "*" And matclas(j, 3) <> "*" Then
        If matclas(j, 1) = matpos(i).TValorDiv And matclas(j, 3) = matpos(i).SerieDiv Then
           largo = largo + 1
        Else
           largo = 0
        End If
     ElseIf matclas(j, 1) = "*" And matclas(j, 2) <> "*" And matclas(j, 3) <> "*" Then
        If matclas(j, 2) = matpos(i).EmisionDiv And matclas(j, 3) = matpos(i).SerieDiv Then
           largo = largo + 1
        Else
           largo = 0
        End If
     ElseIf matclas(j, 1) <> "*" And matclas(j, 2) = "*" And matclas(j, 3) = "*" Then
        If matclas(j, 1) = matpos(i).TValorDiv Then
           largo = largo + 1
        Else
           largo = 0
        End If
     ElseIf matclas(j, 1) = "*" And matclas(j, 2) <> "*" And matclas(j, 3) = "*" Then
        If matclas(j, 2) = matpos(i).EmisionDiv Then
           largo = largo + 1
        Else
           largo = 0
        End If
     ElseIf matclas(j, 1) = "*" And matclas(j, 2) = "*" And matclas(j, 3) <> "*" Then
        If matclas(j, 3) = matpos(i).SerieDiv Then
           largo = largo + 1
        Else
           largo = 0
        End If
     End If
     If largo > maxvars Then
        indice = j
        maxvars = Maximo(maxvars, largo)
     End If
 Next j
BuscarTProd2 = indice
End Function

Sub ClasTProdMD(ByRef matpost() As propPosRiesgo, ByRef matpos() As propPosMD, ByRef txtmsg As String, ByRef exito As Boolean)
'asigna los parametros de valuacion para la posicion de deuda
'esta funcion asigna el modelo de valuacion buscando el tipo valor,emision y serie
'en alguno de los catalogos de valuacion del sistema
'asignando el modelo de valuacion cuando encuentra alguna coincidencia en alguna de
'las tablas
'datos de entrada:
'el array con la posicion de instrumentos de deuda
'salida:
'la misma posición con el modelo y los parameros para valuar cada operacion
'txtmsg - los mensajes arrojados por el subproceso
'exito  - el estado de la finalizacion del proceso

Dim i As Integer
Dim j As Long
Dim indice1 As Integer
Dim indice2 As Integer
Dim indice3 As Integer
Dim indice4 As Integer
Dim indice5 As Integer
Dim indice6 As Integer
'esta funcion solo es heuristica para la posicion de mercado de dinero
'para las demas posiciones hara una busqueda estricta
exito = True
txtmsg = ""
For i = 1 To UBound(matpos, 1)
    indice1 = 0: indice2 = 0: indice3 = 0: indice4 = 0: indice5 = 0: indice6 = 0
    If matpos(i).Tipo_Mov = ClaveCDirec Or matpos(i).Tipo_Mov = ClaveVDirec Then  'operaciones en directo
       indice1 = BuscarTProd(i, matpos, MatTValBC0)
       If indice1 <> 0 Then
          Call AgregarPValBC0(i, indice1, matpos, MatTValBC0)
       End If
       indice2 = BuscarTProd(i, matpos, MatTValBonos)
       If indice2 <> 0 Then
          Call AgregarPValBono(i, indice2, matpos, MatTValBonos)
       End If
       indice3 = BuscarTProd(i, matpos, MatTValSTCupon)
       If indice3 <> 0 Then
          Call AgregarPValBonoSTC(i, indice3, matpos, MatTValSTCupon)
       End If
       indice4 = BuscarTProd(i, matpos, MatTValSTDesc)
       If indice4 <> 0 Then
          Call AgregarPValBonoSTD(i, indice4, matpos, MatTValSTDesc)
       End If
       indice5 = BuscarTProd(i, matpos, MatTValInd)
       If indice5 <> 0 Then
          Call AgregarPValInd(i, indice5, matpos, MatTValInd)
       End If
    Else
       indice6 = BuscarTProd(i, matpos, MatTValReportos)
       If indice6 <> 0 Then
          Call AgregarPValRep(i, indice6, matpos, MatTValReportos)
       End If
    End If
    If indice1 = 0 And indice2 = 0 And indice3 = 0 And indice4 = 0 And indice5 = 0 And indice6 = 0 Then
       txtmsg = txtmsg & "No se clasificó " & matpos(i).C_Posicion & " " & matpos(i).Tipo_Mov & "  " & matpos(i).c_operacion & " " & matpos(i).tValorMD & " " & matpos(i).emisionMD & " " & matpos(i).serieMD & ","
       exito = False
    End If
    For j = 1 To UBound(matpost, 1)
        If matpost(j).IndPosicion = i And matpost(j).No_tabla = 1 Then
           matpost(j).fValuacion = matpos(i).fValuacion
           Exit For
        End If
    Next j
Next i
End Sub

Sub ClasTProdMD2(ByRef matpos() As propPosMD, ByRef txtmsg As String, ByRef exito As Boolean)
'asigna los parametros de valuacion para la posicion de deuda
'buscando el tipo valor,emision y serie
'en alguno de los catalogos de valuacion del sistema
'asignando el modelo de valuacion cuando encuentra alguna coincidencia en alguna de
'las tablas
'datos de entrada:
'el array con la posicion de instrumentos de deuda
'salida:
'la misma posición con los parameros para valuar cada operacion
'txtmsg - los mensajes arrojados por el subproceso
'exito  - el estado de la finalizacion del proceso

Dim i As Integer
Dim j As Long
Dim indice1 As Integer
Dim indice2 As Integer
Dim indice3 As Integer
Dim indice4 As Integer
Dim indice5 As Integer
Dim indice6 As Integer
exito = True
txtmsg = ""
For i = 1 To UBound(matpos, 1)
    indice1 = 0: indice2 = 0: indice3 = 0: indice4 = 0: indice5 = 0: indice6 = 0
    If matpos(i).Tipo_Mov = ClaveCDirec Or matpos(i).Tipo_Mov = ClaveVDirec Then  'operaciones en directo
       indice1 = BuscarTProd(i, matpos, MatTValBC0)
       If indice1 <> 0 Then
          Call AgregarPValBC0(i, indice1, matpos, MatTValBC0)
       End If
       indice2 = BuscarTProd(i, matpos, MatTValBonos)
       If indice2 <> 0 Then
          Call AgregarPValBono(i, indice2, matpos, MatTValBonos)
       End If
       indice3 = BuscarTProd(i, matpos, MatTValSTCupon)
       If indice3 <> 0 Then
          Call AgregarPValBonoSTC(i, indice3, matpos, MatTValSTCupon)
       End If
       indice4 = BuscarTProd(i, matpos, MatTValSTDesc)
       If indice4 <> 0 Then
          Call AgregarPValBonoSTD(i, indice4, matpos, MatTValSTDesc)
       End If
       indice5 = BuscarTProd(i, matpos, MatTValInd)
       If indice5 <> 0 Then
          Call AgregarPValInd(i, indice5, matpos, MatTValInd)
       End If
    Else
       indice6 = BuscarTProd(i, matpos, MatTValReportos)
       If indice6 <> 0 Then
          Call AgregarPValRep(i, indice6, matpos, MatTValReportos)
       End If
    End If
    If indice1 = 0 And indice2 = 0 And indice3 = 0 And indice4 = 0 And indice5 = 0 And indice6 = 0 Then
       txtmsg = txtmsg & "No se clasificó " & matpos(i).C_Posicion & " " & matpos(i).Tipo_Mov & "  " & matpos(i).c_operacion & " " & matpos(i).tValorMD & " " & matpos(i).emisionMD & " " & matpos(i).serieMD & ","
       exito = False
    End If
Next i
End Sub

Sub ClasTprod2(ByRef matpost() As propPosRiesgo, ByRef matpos() As propPosDiv, ByRef txtmsg As String, ByRef exito As Boolean)
Dim i As Long
Dim j As Long
Dim indice5 As Integer
exito = True
txtmsg = ""
For i = 1 To UBound(matpos, 1)
    indice5 = BuscarTProd2(i, matpos, MatTValInd)
    If indice5 <> 0 Then
       Call AgregarPValIndDiv(i, indice5, matpos, MatTValInd)
    End If
    If indice5 = 0 Then
       txtmsg = "No se clasificó " & matpos(i).C_Posicion & " " & matpos(i).Tipo_Mov & "  " & matpos(i).c_operacion & " " & matpos(i).TValorDiv & " " & matpos(i).EmisionDiv & " " & matpos(i).SerieDiv
       exito = False
       Exit Sub
    End If
    For j = 1 To UBound(matpost, 1)
        If matpost(j).IndPosicion = i And matpost(j).No_tabla = 2 Then
           matpost(j).fValuacion = matpos(i).fValuacion
           Exit For
        End If
    Next j

Next i

End Sub

Sub AgregarPValBC0(i, indice, matpos, matclas)
    matpos(i).fValuacion = matclas(indice, 4)               'funcion de valuacion
    matpos(i).fRiesgo1MD = matclas(indice, 5)            'curva desc 1
    matpos(i).tInterpol1MD = matclas(indice, 6)          'tipo interpolacion 1
    matpos(i).tCambioMD = ReemplazaVacioValor(matclas(indice, 7), "")           'tipo cambio
End Sub

Sub AgregarPValBono(i, indice, matpos, matclas)
    matpos(i).fValuacion = matclas(indice, 4)                   'funcion de valuacion
    If matclas(indice, 5) <> 0 Then
       matpos(i).pCuponMD = matclas(indice, 5)              'pcupon general
    End If
    matpos(i).fRiesgo1MD = matclas(indice, 6)               'factor de descuento
    matpos(i).tInterpol1MD = matclas(indice, 7)             'tipo interpolacion 1
    matpos(i).tCambioMD = ReemplazaVacioValor(matclas(indice, 8), "")               'tipo cambio acttiva
End Sub

Sub AgregarPValRep(i, indice, matpos, matclas)
    matpos(i).fValuacion = matclas(indice, 4)                 'funcion de valuacion
    matpos(i).fRiesgo1MD = matclas(indice, 5)                 'curva desc 1
    matpos(i).tInterpol1MD = matclas(indice, 6)               'tipo interpolacion 1
    matpos(i).tCambioMD = ReemplazaVacioValor(matclas(indice, 7), "")                 'TIPO DE cambio
End Sub

Sub AgregarPValBonoSTD(i, indice, matpos, matclas)
       matpos(i).fValuacion = matclas(indice, 4)              'funcion de valuacion
       If matclas(indice, 5) <> 0 Then
          matpos(i).pCuponMD = matclas(indice, 5)             'periodo cupon
       End If
       matpos(i).fRiesgo1MD = matclas(indice, 6)              'factor 1
       matpos(i).tInterpol1MD = matclas(indice, 7)            'tipo interpolacion 1
       matpos(i).fRiesgo2MD = matclas(indice, 8)              'factor 2
       matpos(i).tInterpol2MD = matclas(indice, 9)            'tipo interpolacion
End Sub

Sub AgregarPValBonoSTC(i, indice, matpos, matclas)
    matpos(i).fValuacion = matclas(indice, 4)              'funcion de valuacion
    If matclas(indice, 7) <> 0 Then
       matpos(i).pCuponMD = matclas(indice, 7)             'pcupon general
    End If
    matpos(i).sTCuponMD = matclas(indice, 8)               'sobretasa de cupon
    matpos(i).fRiesgo1MD = matclas(indice, 9)              'yield
    matpos(i).tInterpol1MD = matclas(indice, 10)           'tipo interpolacion 1
    matpos(i).fRiesgo2MD = matclas(indice, 11)             'factor 2
    matpos(i).tInterpol2MD = matclas(indice, 12)           'interpolacion 2
    matpos(i).tCambioMD = ReemplazaVacioValor(matclas(indice, 13), "")         'tipo de cambio
End Sub

Sub AgregarPValInd(i, indice, matpos, matclas)
    matpos(i).fValuacion = matclas(indice, 4)                'funcion de valuacion
    matpos(i).tCambioMD = matclas(indice, 5)            'tipo de cambio a aplicar
End Sub


Sub AgregarPValIndDiv(ByVal i As Integer, ByVal indice As Integer, ByRef matpos() As propPosDiv, ByRef matclas() As Variant)
    matpos(i).fValuacion = matclas(indice, 4)                'funcion de valuacion
    matpos(i).TCambioDiv = matclas(indice, 5)            'tipo de cambio a aplicar
End Sub


Sub AgregarPValSwaps(ByRef matpost() As propPosRiesgo, ByRef matpos() As propPosSwaps, ByRef matclas() As Variant, ByRef txtmsg As String, ByRef sicontinua As Boolean)
    Dim i As Integer, indice As Integer, noreg As Integer
    Dim j As Long
    Dim txtclave As String
    sicontinua = True
    txtmsg = ""
    noreg = UBound(matpos, 1)
    For i = 1 To noreg
        indice = BuscarValorArray(matpos(i).cProdSwapGen, matclas, 1)
        If indice <> 0 Then                                           'se aplican las definiciones de la tabla t_var_valuacion y algunos parametros de valuacion
           matpos(i).fValuacion = matclas(indice, 2)                 'funcion de valuacion
           If matclas(indice, 3) <> 0 Then
              matpos(i).PCuponActSwap = matclas(indice, 3)        'pcupon activa
           End If
           If matclas(indice, 4) <> 0 Then
              matpos(i).PCuponPasSwap = matclas(indice, 4)          'pcupon pasiva
           End If
           matpos(i).FRiesgo1Swap = matclas(indice, 5)            'curva desc 1
           matpos(i).TInterpol1Swap = matclas(indice, 6)          'tipo interpolacion desc 1
           matpos(i).FRiesgo2Swap = matclas(indice, 7)            'curva desc 2
           matpos(i).TInterpol2Swap = matclas(indice, 8)          'tipo interpolacion desc 2
           matpos(i).FRiesgo3Swap = ReemplazaVacioValor(matclas(indice, 9), "")           'curva pago 1
           matpos(i).TInterpol3Swap = ReemplazaVacioValor(matclas(indice, 10), 0)        'interpolacion pago 1
           matpos(i).FRiesgo4Swap = ReemplazaVacioValor(matclas(indice, 11), "")           'curva pago 2
           matpos(i).TInterpol4Swap = ReemplazaVacioValor(matclas(indice, 12), 0)        'interpolacion pago 2
           If Not EsVariableVacia(matclas(indice, 13)) Then
              matpos(i).TCambio1Swap = matclas(indice, 13)          'tipo cambio acttiva
           Else
              matpos(i).TCambio1Swap = ""
           End If
           If Not EsVariableVacia(matclas(indice, 14)) Then
              matpos(i).TCambio2Swap = matclas(indice, 14)          'tipo cambio pasiva
           Else
              matpos(i).TCambio2Swap = ""
           End If
           For j = 1 To UBound(matpost, 1)
               If matpost(j).IndPosicion = i And matpost(j).No_tabla = 3 Then
                  matpost(j).fValuacion = matpos(i).fValuacion
                  Exit For
               End If
           Next j
        Else
           MensajeProc = "No se encuentra el swap: " & matpos(i).c_operacion & "  " & matpos(i).cProdSwapGen & " en el catalogo de Valuacion. No es posible valuar la operacion"
           txtmsg = MensajeProc
           sicontinua = False
        End If
        AvanceProc = i / noreg
        MensajeProc = "Obteniendo informacion de valuacion " & Format$(AvanceProc, "##0.00 %")
        DoEvents
    Next i
End Sub

Sub AgregarPValDeuda(ByRef matpost() As propPosRiesgo, ByRef matpos() As propPosDeuda, ByRef matclas() As Variant, ByRef txtmsg As String, ByRef sicontinua As Boolean)
    Dim i As Integer, indice As Integer, noreg As Integer
    Dim j As Integer
    Dim txtclave As String
    sicontinua = True
    txtmsg = ""
    noreg = UBound(matpos, 1)
    For i = 1 To noreg
        If Not EsVariableVacia(matpos(i).ProductoDeuda) Then
           indice = BuscarValorArray(matpos(i).ProductoDeuda, matclas, 1)
           If indice <> 0 Then                                                'se aplican las definiciones de la tabla t_var_valuacion y algunos parametros de valuacion
               matpos(i).fValuacion = matclas(indice, 2)                     'funcion de valuacion
               If matclas(indice, 3) <> 0 Then
                   matpos(i).PCuponDeuda = matclas(indice, 3)               'pcupon de referencia
               End If
               matpos(i).FRiesgo1Deuda = matclas(indice, 4)                            'curva desc
               matpos(i).TInterpol1Deuda = matclas(indice, 5)                          'tipo interpolacion desc
               matpos(i).FRiesgo2Deuda = ReemplazaVacioValor(matclas(indice, 6), "")                'curva pago
               matpos(i).TInterpol2Deuda = ReemplazaVacioValor(matclas(indice, 7), 0)              'interpolacion pago
               matpos(i).TCambioDeuda = ReemplazaVacioValor(matclas(indice, 8), "")                  'tipo cambio acttiva
               For j = 1 To UBound(matpost, 1)
                   If matpost(j).IndPosicion = i And matpost(j).No_tabla = 5 Then
                      matpost(j).fValuacion = matpos(i).fValuacion
                      Exit For
                   End If
               Next j
           Else
               MensajeProc = "No se encuentra la operacion: " & matpos(i).c_operacion & "  " & matpos(i).ProductoDeuda & ", mesa o posicion: " & matpos(i).C_Posicion & " en el catalogo de Valuacion. No es posible valuar la operacion"
               MsgBox MensajeProc
               sicontinua = False
           End If
        Else
           MensajeProc = "La clave de funcion de la operacion: " & matpos(i).c_operacion & "  " & matpos(i).ProductoDeuda & " no es valida"
           txtmsg = MensajeProc
           sicontinua = False
        End If
        AvanceProc = i / noreg
        MensajeProc = "Obteniendo informacion de valuacion " & Format$(AvanceProc, "##0.00 %")
        DoEvents
    Next i


End Sub

Function TraduceDerivadoEstandar2(ByVal txtclave As String) As String
'objetivo de la funcion: asigna un nombre en catalogo de una clave
'del sistema ikos derivados
'con la tabla DERIVADO_ESTANDAR y los campos de esta determinar que tipo de swap
'esta tomado en el campo C_PRODUCTO de la tabla VAR_TC_REL_S_IKOS_S
'de esta manera el sistema sabe como debe de valuar el instrumento a comparar C_PRODUCTO
'con el campo CPRODUCTO de las tablas VAR_TC_VAL_SWAPS1 y VAR_TC_VAL_SWAPS2

'la tabla DERIVADO_ESTANDAR la actualiza el sistema IKOS derivados


Dim moneda_act As String
Dim moneda_pas As String
Dim tasa_act As String
Dim tasa_pas As String
Dim i As Integer, valor As String
Dim j As Integer
    For i = 1 To UBound(MatDerEstandar, 1)
        If Trim$(Left$(txtclave, 49)) = Trim$(Left$(MatDerEstandar(i, 2), 49)) Then
           valor = ""
           moneda_act = MatDerEstandar(i, 4)
           moneda_pas = MatDerEstandar(i, 5)
           tasa_act = MatDerEstandar(i, 8)
           tasa_pas = MatDerEstandar(i, 9)
           For j = 1 To UBound(MatRelSwapIS, 1)
               If moneda_act = MatRelSwapIS(j, 2) And moneda_pas = MatRelSwapIS(j, 3) And tasa_act = MatRelSwapIS(j, 4) And tasa_pas = MatRelSwapIS(j, 5) Then
                  valor = MatRelSwapIS(j, 1)
                  Exit For
               End If
           Next j
         Exit For
         End If
    Next i
    If Not EsVariableVacia(valor) Then
        TraduceDerivadoEstandar2 = valor
    Else
        MensajeProc = "No se encontro clave para este derivado " & txtclave
        MsgBox MensajeProc
        TraduceDerivadoEstandar2 = ""
    End If

End Function


Sub AgregarPValFwd(ByRef matpost() As propPosRiesgo, ByRef matpos() As propPosFwd, ByRef matclas() As Variant, ByRef txtmsg As String, ByRef sicontinua As Boolean)
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim indice As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
sicontinua = True
txtmsg = ""
noreg = UBound(matpos, 1)
For i = 1 To noreg
    If matpos(i).C_Posicion = 4 Then
       indice = BuscarValorArray(matpos(i).ClaveProdFwd, matclas, 1)
       If indice <> 0 Then  'se aplican las definiciones de la tabla t_var_valuacion y algunos parametros de valuacion
          matpos(i).fValuacion = matclas(indice, 2)              'funcion de valuacion
          matpos(i).FRiesgo1Fwd = matclas(indice, 3)            'curva desc 1
          matpos(i).TInterpol1Fwd = matclas(indice, 4)         'tipo interpolacion desc 1
          matpos(i).FRiesgo2Fwd = matclas(indice, 5)          'curva desc 2
          matpos(i).TInterpol2Fwd = matclas(indice, 6)         'tipo interpolacion desc 2
          matpos(i).FRiesgo3Fwd = matclas(indice, 7)           'curva pago 1
          matpos(i).TInterpol3Fwd = matclas(indice, 8)         'interpolacion pago 1
          matpos(i).TCambioFwd = matclas(indice, 11)          'tipo cambio
          For j = 1 To UBound(matpost, 1)
              If matpost(j).IndPosicion = i And matpost(j).No_tabla = 4 Then
                 matpost(j).fValuacion = matpos(i).fValuacion
                 Exit For
              End If
          Next j

       Else
         MensajeProc = "No se encuentra la operacion: " & matpos(i).c_operacion & ", mesa o posicion: " & matpos(i).C_Posicion & " en el catalogo de Valuacion. No es posible valuar la operacion"
         txtmsg = MensajeProc
         sicontinua = False
       End If
    End If
    
    AvanceProc = i / noreg
    MensajeProc = "Obteniendo informacion de valuacion " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0

End Sub

Function LeerArchTexto(ByVal nombre As String, ByVal letra As String, ByVal txtmensaje As String) As Variant()
Dim maxcampos As Integer
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim matc() As String
Dim novar As Integer
Dim contar As Long
Dim contar1 As Long

'se abre el archivo de texto con la tabla separada por comas
'la unica manera de saber cuantos registros va uno a abrir es
'leyendo 2 veces la tabla de datos
'PRIMERO SE ABRE UNA VEZ PARA VER CUANTOS RENGLONES HAY EN LA BASE
' Y TAMBIEN PARA SABER CUAL ES EL NO MAXIMO DE CAMPOS
'EN CADA RENGLON

 maxcampos = 0
 contar = 0
 Dim mat1() As String
 Open nombre For Input As #1
 Do While Not EOF(1)
 contar = contar + 1
 ReDim Preserve mat1(1 To contar) As String
 Line Input #1, mat1(contar)
 Loop
 Close #1
 
noreg = UBound(mat1, 1)
contar = 0
For i = 1 To noreg
    If Not EsVariableVacia(mat1(i)) Then
       contar = contar + 1
       matc = EncontrarSubCadenas(mat1(i), letra)
       novar = UBound(matc, 1)
       maxcampos = Maximo(maxcampos, novar)
    End If
    AvanceProc = i / noreg
    MensajeProc = "Determinando el no de variables " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i

'ahora si se lee la tabla para su importacion
contar1 = 0
If contar <> 0 And maxcampos <> 0 Then
   ReDim mata(1 To contar, 1 To maxcampos) As Variant
   For i = 1 To contar
       If Not EsVariableVacia(mat1(i)) Then
          contar1 = contar1 + 1
          matc = EncontrarSubCadenas(mat1(i), letra)
          For j = 1 To UBound(matc, 1)
               mata(contar1, j) = matc(j)
          Next j
       End If
       AvanceProc = i / noreg
       MensajeProc = txtmensaje & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
Else
  ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerArchTexto = mata
End Function

Sub GuardaPosPensionO(ByVal fecha As Date, ByRef mata() As Variant, ByVal txtbase As String)
Dim txtfecha As String
Dim noreg As Integer
Dim i As Integer
Dim nocol As Integer
Dim txtcadena As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta es la estructura actual de la posición del fondo de pensiones


'1 fecha pos
'2 intencion
'3 tipo de operacion
'4 tipo valor
'5 emision
'6 serie
'7 no titulos
'8 fecha vencimiento
'9 valor nominal
'10 plazo cupon
'11 tasa cupon
'12 tasa reporto
'13 precio pactado
'14 propietario
'15 fecha compra
'16 tipo mercado
'17 mesa o posicion
'18 emision



'1  FECHA
'2  intencion
'3  toperacion
'4  CPORTAFOLIO
'5  TV
'6  EMISION
'7  serie
'8  CEMISION
'9  ntitulos
'10 FCOMPRA
'11 PCOMPRA
'12 FVENCIMIENTO
'13 VNOMINAL
'14 PCUPON
'15 tcupon
'16 TPREMIO
'17 MONEDA

txtfecha = Format(fecha, "dd/mm/yyyy")
ConAdo.Execute "DELETE FROM " & txtbase & " WHERE FECHA = '" & txtfecha & "'"
noreg = UBound(mata, 1)
If noreg <> 0 Then
nocol = UBound(mata, 2)
For i = 1 To noreg
If mata(i, 17) <> "0" And mata(i, 6) <> "B8" Then
 txtcadena = "INSERT INTO " & txtbase & " VALUES("
 txtfecha = Format(fecha, "dd/mm/yyyy")
 txtcadena = txtcadena & "'" & txtfecha & "',"                   '1 fecha de la posicion
 txtcadena = txtcadena & "'" & mata(i, 2) & "',"                 '2 negociacion o vencimiento
 txtcadena = txtcadena & "'" & Trim(mata(i, 3)) & "',"           '3 tipo de operacion
 txtcadena = txtcadena & "'" & Trim(mata(i, 14)) & "',"           '4 clave portafolio
 txtcadena = txtcadena & "'" & Trim(mata(i, 4)) & "',"           '5 tv
 txtcadena = txtcadena & "'" & Trim(mata(i, 5)) & "',"           '6 emision
 txtcadena = txtcadena & "'" & Trim(mata(i, 6)) & "',"           '7 serie
 txtcadena = txtcadena & "'" & Trim(mata(i, 18)) & "',"          '8 clave emision generada por el sistema
 txtcadena = txtcadena & CDbl(Trim(mata(i, 7))) & ","            '9 no titulos
 If IsDate(mata(i, 15)) And mata(i, 15) <> 0 Then                '10 fecha de compra
    txtfecha = "to_date('" & Format(mata(i, 15), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = txtcadena & txtfecha & ","
 Else
  txtcadena = txtcadena & "null" & ","
 End If
 txtcadena = txtcadena & CDbl(Trim(mata(i, 13))) & ","           '11 precio de compra
 If IsDate(mata(i, 8)) Then                                      '12 fecha de vencimiento
  txtfecha = "to_date('" & Format(mata(i, 8), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtcadena = txtcadena & txtfecha & ","
 Else
  txtcadena = txtcadena & "null" & ","
 End If
 txtcadena = txtcadena & Val(Trim(mata(i, 9))) & ","            '13 valor nominal
 If Not EsVariableVacia(mata(i, 10)) Then                             '14 periodo cupon
  txtcadena = txtcadena & Val(mata(i, 10)) & ","
 Else
  txtcadena = txtcadena & "null,"
 End If
 If Not EsVariableVacia(mata(i, 11)) Then                            '15 tasa cupon
  txtcadena = txtcadena & Val(mata(i, 11)) & ","
 Else
  txtcadena = txtcadena & "null,"
 End If
 If Not EsVariableVacia(mata(i, 12)) Then                             '16 tasa premio
  txtcadena = txtcadena & Val(mata(i, 12)) & ","
 Else
  txtcadena = txtcadena & "null,"
 End If
 txtcadena = txtcadena & "1)"                 'moneda
' MsgBox txtcadena
 ConAdo.Execute txtcadena
End If
 AvanceProc = i / noreg
 MensajeProc = "Guardando la posicion de fondo de pensiones del " & fecha & " : " & Format(AvanceProc, "##0.00 %")

 
Next i
 Call MostrarMensajeSistema("Se guardaron " & noreg & "  registros de la posicion del fondo de pensiones del " & fecha & " : " & Format(AvanceProc, "#,##0.00 %"), frmProgreso.Label2, 1, Date, Time, NomUsuario)
Else
 Call MostrarMensajeSistema("Atencion. Faltan registros de la posición de pensiones", frmProgreso.Label2, 1, Date, Time, NomUsuario)
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub GuardarVPrecios(fecha, nombase, mata, nreg1, exito)
Dim noreg As Long
Dim nreg As Long
Dim contar As Long
Dim sql_precios As String
Dim sql_num_precios As String
Dim i As Long
Dim txtfecha As String
Dim cadborra As String
Dim txtcadena As String


'se guardan los datos del vector de precios de la matriz mata
noreg = UBound(mata, 1)
nreg = 0
contar = 0
If noreg <> 0 Then
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   cadborra = "DELETE FROM " & nombase & " WHERE FECHA = " & txtfecha
   ConAdo.Execute cadborra
   For i = 1 To noreg
       If CDate(mata(i, 1)) = fecha And (mata(i, 2) <> "Mercado" And mata(i, 2) <> "Mercado") Then
          txtcadena = "INSERT INTO " & nombase & " VALUES("
          txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtcadena = txtcadena & txtfecha & ",'"         'fecha
          txtcadena = txtcadena & mata(i, 2) & "','"      'mercado
          txtcadena = txtcadena & mata(i, 3) & "','"      'tv
          txtcadena = txtcadena & mata(i, 4) & "','"      'emision
          txtcadena = txtcadena & mata(i, 5) & "',"       'serie
          txtcadena = txtcadena & mata(i, 6) & ","        'psucio
          txtcadena = txtcadena & mata(i, 7) & ","        'plimpio
          txtcadena = txtcadena & mata(i, 8) & ","        'intereses md
          txtcadena = txtcadena & mata(i, 9) & ","        '
          txtcadena = txtcadena & mata(i, 10) & ","
          txtcadena = txtcadena & mata(i, 11) & ","
          txtfecha = "to_date('" & Format(mata(i, 12), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtcadena = txtcadena & txtfecha & ","
          txtfecha = "to_date('" & Format(mata(i, 13), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtcadena = txtcadena & txtfecha & ","
          txtcadena = txtcadena & mata(i, 14) & ","
          txtcadena = txtcadena & mata(i, 15) & ","
          txtcadena = txtcadena & mata(i, 16) & ","
          txtcadena = txtcadena & Val(mata(i, 17)) & ","                  'yield
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 18), "") & "',"               'calificacion de sp
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 19), "") & "',"               'calificacion de fitch
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 20), "") & "',"               'calificacion de moodys
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 21), "") & "',"               'calificacion de hr
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 22), "") & "',"               'calificacion de verum
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 23), "") & "',"               'calificacion de verum
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 24), "") & "',"               'calificacion de dbrs
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 25), "") & "',"               'frecuencia cupon
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 26), "") & "',"               'regla cupon
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 27), 0) & ","                       'sobretasa de colocacion
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 28), 0) & ","                       'monto emitido
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 29), 0) & ","                       'monto en circulacion
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 30), 0) & "',"                'sector
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 31), 0) & "',"                'isin
          txtcadena = txtcadena & "'" & ReemplazaVacioValor(mata(i, 32), "") & "')"               'la clave de la emision
          ConAdo.Execute txtcadena
          nreg1 = nreg1 + 1
          contar = contar + 1
  Else
    'MsgBox "para esta emison no se guarda " & mata(i, 30)
  End If
  AvanceProc = i / noreg
  MensajeProc = "Guardando el vector de precios de " & fecha & " : " & Format(AvanceProc, "##0.00 %")
  DoEvents
 Next i
 If contar = noreg - 1 Then
    exito = True
    MensajeProc = "Se guardo el vector de precios del dia " & fecha
 Else
   MensajeProc = "No se guardo correctamente el vector de precios de " & fecha
   exito = False
 End If
End If
End Sub

Sub ImportarPosCambios(ByVal fecha As Date, ByVal siarch As Boolean, ByVal direc As String, ByRef noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
If siarch Then
   Call ImportarPosCamArch(fecha, direc, noreg, txtmsg, exito)
Else
   Call ImportarPosCamRed(fecha, noreg, txtmsg, exito)
End If
End Sub

Sub ImportarPosCamRed(ByVal fecha As Date, ByRef noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfiltro As String
Dim noreg1 As Integer
Dim contar As Integer
Dim i As Integer
Dim txtfecha As String
Dim txtcadena As String
Dim j As Integer
Dim tv As String
Dim emision As String
Dim serie As String
Dim cemision As String
Dim coperacion As String
Dim TOPERACION As Integer
Dim intencion As String
Dim tiempo As String
Dim rmesa As New ADODB.recordset


txtfiltro = "SELECT COUNT(*) FROM " & TablaInterfDiv
rmesa.Open txtfiltro, ConAdo
noreg1 = rmesa.Fields(0)
rmesa.Close
contar = 0
exito = False
If noreg1 <> 0 Then
   txtfiltro = "SELECT * FROM " & TablaInterfDiv
   rmesa.Open txtfiltro, ConAdo
   ReDim mata(1 To noreg1, 1 To 4) As Variant
   rmesa.MoveFirst
   For i = 1 To noreg1
      For j = 1 To 4
         mata(i, j) = rmesa.Fields(j - 1)
      Next j
      rmesa.MoveNext
   Next i
   rmesa.Close
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtcadena = "DELETE FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha & " AND TIPOPOS = 1"
   ConAdo.Execute txtcadena
   For i = 1 To noreg1
      If Val(mata(i, 4)) <> 0 Then
       contar = contar + 1
       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       tiempo = Format(Time, "hhmmss")
       If mata(i, 1) = "02" Then        'dolares
          coperacion = "1"
          tv = "*C"
          emision = "MXPUSD"
          serie = "FIX"
       ElseIf mata(i, 1) = "27" Then    'euros
          coperacion = "3"
          tv = "*C"
          emision = "MXPEUR"
          serie = "EUR"
       ElseIf mata(i, 1) = "06" Then    'yenes japoneses
          coperacion = "2"
          tv = "*C"
          emision = "MXPJPY"
          serie = "JPY"
       Else
          coperacion = "4"
          tv = "x"
          emision = "x"
          serie = "x"
       End If
       cemision = tv & emision & serie
       If mata(i, 4) >= 0 Then
          TOPERACION = 1
       Else
          TOPERACION = 4
       End If
       If mata(i, 3) = "5011" Then
          intencion = "N"
       Else
          intencion = "C"
       End If
       txtcadena = "INSERT INTO " & TablaPosDiv & " VALUES("
       txtcadena = txtcadena & "1,"
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "null,"
       txtcadena = txtcadena & tiempo & ","
       txtcadena = txtcadena & "3,"
       txtcadena = txtcadena & "'" & intencion & "',"
       txtcadena = txtcadena & "'" & coperacion & "',"
       txtcadena = txtcadena & TOPERACION & ","
       txtcadena = txtcadena & "'" & tv & "',"
       txtcadena = txtcadena & "'" & emision & "',"
       txtcadena = txtcadena & "'" & serie & "',"
       txtcadena = txtcadena & "'" & cemision & "',"
       txtcadena = txtcadena & "'" & Abs(mata(i, 4)) & "',"
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "0)"
       ConAdo.Execute txtcadena
      End If
   Next i
   exito = True
End If
noreg = contar
End Sub

Sub ImportarPosCamArch(ByVal fecha As Date, ByVal direc As String, ByRef noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim a As Boolean
Dim noregmc As Long
Dim nomarch As String
Dim sihayarch As Boolean
Dim matpos() As Variant
Dim txtfecha As String
Dim cont As Integer
Dim i As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'rutina que genera la posición de divisas
'se lee la posicion de divisas
'1 fecha de la posicion
'2 monto de la divisa
'3 clave de moneda
'4 fecha valor
'5 vacia
'6 cero
'7 compra venta o posicion
'8 descripcion
'la clave de la moneda es muy importante debido a que con esta
'clave hace del desglose de la posicion en el resumen ejecutivo
noreg = 0
exito = False
a = NoLabMX(fecha)
If Not a Then
nomarch = direc & "\" & "DIV" & Format(fecha, "yymmdd") & ".TXT"
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
 matpos = LeerArchTexto(nomarch, ",", "Leyendo la posicion de la mesa de cambios")
 txtfecha = Format(fecha, "dd/mm/yyyy")
 cont = UBound(matpos, 1)
 ReDim mata(1 To cont, 1 To 13) As Variant
 noregmc = 0
 '1  fecha de la posicion
 
 For i = 1 To cont
      mata(i, 1) = fecha                             'fecha de la posicion
      mata(i, 2) = ClavePosMC                        'clave de la posicion
      mata(i, 3) = "N"                               'intencion
      mata(i, 4) = i                                 'clave de la operacion
     If Val(matpos(i, 2)) >= 0 Then                  'tipo de operacion, larga o corta
        mata(i, 5) = 1
     Else
        mata(i, 5) = 4
     End If
  If matpos(i, 3) = "DOLAR AMERICANO" Then
     mata(i, 6) = "*C"
     mata(i, 7) = "MXPUSD"
     mata(i, 8) = "FIX"
     mata(i, 9) = "*CMXPUSDFIX"
  End If
  If matpos(i, 3) = "YEN JAPONES" Then
     mata(i, 6) = "*C"
     mata(i, 7) = "MXPJPY"
     mata(i, 8) = "JPY"
     mata(i, 9) = "*CMXPJPYJPY"
  End If
  If matpos(i, 3) = "EUROS" Then
     mata(i, 6) = "*C"
     mata(i, 7) = "MXPEUR"
     mata(i, 8) = "EUR"
     mata(i, 9) = "*CMXPEUREUR"
  End If
    mata(i, 10) = Abs(CDbl(matpos(i, 2)))          'no de titulos
    mata(i, 11) = CDate(matpos(i, 1))              'fecha de compra
    mata(i, 12) = 0                                'precio de compra
    mata(i, 13) = CDate(matpos(i, 1))              'fecha de liquidacion
 Next i
 Call GuardaPosDiv(mata, 1, "Real", "000000", noregmc, exito)
 MensajeProc = "Se guardaron " & noregmc & " registros de la posicion de riesgo de la mesa de cambios " & fecha
 If exito Then txtmsg = "El proceso finalizo correctamente"
 Else
  exito = False
  MensajeProc = "No hay posicion de divisas " & fecha
  txtmsg = MensajeProc
  noregmc = 0
End If
End If
noreg = noregmc
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ImpPosMCambiosSim(ByVal fecha As Date, ByVal nomarch As String, ByVal nompos As String, ByRef noreg As Long, ByRef exito As Boolean)
Dim sihayarch As Boolean
Dim matpos() As Variant
Dim txtfecha As String
Dim cont As Integer
Dim i As Integer
Dim noregmc As Integer

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'rutina que genera la posición de divisas
'se lee la posicion de divisas
'1 fecha de la posicion
'2 monto de la divisa
'3 clave de moneda
'4 fecha valor
'5 vacia
'6 cero
'7 compra venta o posicion
'8 descripcion
'la clave de la moneda es muy importante debido a que con esta
'clave hace del desglose de la posicion en el resumen ejecutivo
noreg = 0

exito = False


sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
   matpos = LeerArchTexto(nomarch, ",", "Leyendo la posicion de la mesa de cambios")
   txtfecha = Format(fecha, "dd/mm/yyyy")
   cont = UBound(matpos, 1)
 ReDim mata(1 To cont, 1 To 13) As Variant
   noregmc = 0
 For i = 1 To cont
      mata(i, 1) = fecha                             'fecha de la posicion
      mata(i, 2) = ClavePosMC                        'clave de la posicion
      mata(i, 3) = "N"                               'intencion
      mata(i, 4) = i                                 'clave de la operacion
     If Val(matpos(i, 2)) >= 0 Then                  'tipo de operacion, larga o corta
        mata(i, 5) = 1
     Else
        mata(i, 5) = 4
     End If
  If matpos(i, 3) = "DOLAR AMERICANO" Then
     mata(i, 6) = "*C"
     mata(i, 7) = "MXPUSD"
     mata(i, 8) = "FIX"
     mata(i, 9) = "*CMXPUSDFIX"
  End If
  If matpos(i, 3) = "YEN JAPONES" Then
     mata(i, 6) = "*C"
     mata(i, 7) = "MXPJPY"
     mata(i, 8) = "JPY"
     mata(i, 9) = "*CMXPJPYJPY"
  End If
  If matpos(i, 3) = "EUROS" Then
     mata(i, 6) = "*C"
     mata(i, 7) = "MXPEUR"
     mata(i, 8) = "EUR"
     mata(i, 9) = "*CMXPEUREUR"
  End If
    mata(i, 10) = Abs(CDbl(matpos(i, 2)))          'no de titulos
    mata(i, 11) = CDate(matpos(i, 1))              'fecha de compra
    mata(i, 12) = 0                                'precio de compra
    mata(i, 13) = CDate(matpos(i, 1))              'fecha de liquidacion
 Next i
    Call GuardaPosDivSim(mata, 2, nompos, noregmc, exito)
 Else
 noregmc = 0
End If
MensajeProc = "Se guardaron " & noregmc & " registros de la posicion de riesgo de la mesa de cambios " & fecha
noreg = noreg + noregmc
If noreg = 0 Then Call MostrarMensajeSistema("Se guardaron " & noreg & " registros de la posicion de riesgo de la mesa de cambios ", frmProgreso.Label2, 1, Date, Time, NomUsuario)
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub GuardaPosDiv(ByRef mata() As Variant, ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horapos As String, ByRef noreg As Long, ByRef exito As Boolean)
Dim txtfecha As String
Dim i As Long
Dim txtcadena As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String

txtfecha = "to_date('" & Format(mata(1, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
ConAdo.Execute "DELETE FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
noreg = 0
 For i = 1 To UBound(mata, 1)
   txtcadena = "INSERT INTO " & TablaPosDiv & " VALUES("
   txtfecha1 = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha2 = "to_date('" & Format(mata(i, 11), "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha3 = "to_date('" & Format(mata(i, 13), "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtcadena = txtcadena & "'" & tipopos & "',"           'tipo de posicion
   txtcadena = txtcadena & txtfecha1 & ","                'fecha de la posicion
   txtcadena = txtcadena & "'" & txtnompos & "',"         'nombre de la posicion
   txtcadena = txtcadena & "'" & horapos & "',"           'hora de la posicion
   txtcadena = txtcadena & "'" & mata(i, 3) & "',"        'intencion
   txtcadena = txtcadena & mata(i, 2) & ","               'clave de la posicion
   txtcadena = txtcadena & "'" & mata(i, 4) & "',"        'clave de operacion
   txtcadena = txtcadena & mata(i, 5) & ","               'tipo de operacion
   txtcadena = txtcadena & "'" & mata(i, 6) & "',"        'tipo valor
   txtcadena = txtcadena & "'" & mata(i, 7) & "',"        'emision
   txtcadena = txtcadena & "'" & mata(i, 8) & "',"        'serie
   txtcadena = txtcadena & "'" & mata(i, 9) & "',"        'clave de emision
   txtcadena = txtcadena & mata(i, 10) & ","              'no de titulos
   txtcadena = txtcadena & txtfecha2 & ","                'fecha de compra
   txtcadena = txtcadena & txtfecha3 & ","                'fecha de liquidacion
   txtcadena = txtcadena & mata(i, 12) & ","              'precio de compra
   txtcadena = txtcadena & "null,"                        'subportafolio 1
   txtcadena = txtcadena & "null,"                        'subportafolio 2
   txtcadena = txtcadena & "null)"                        'calificacion
   ConAdo.Execute txtcadena
   noreg = noreg + 1
   MensajeProc = "Creando posicion de riesgo de la mesa de cambios "
   DoEvents
 Next i
 If noreg = UBound(mata, 1) Then
    exito = True
 Else
    exito = False
 End If
End Sub

Sub GuardaPosDivSim(ByRef mata() As Variant, ByVal tipopos As Integer, ByVal txtnompos As String, ByRef noreg As Integer, ByRef exito As Boolean)
Dim i As Integer
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtcadena As String

ConAdo.Execute "DELETE FROM " & TablaPosDiv & " WHERE NOMPOS = '" & txtnompos & "'"
noreg = 0
 For i = 1 To UBound(mata, 1)
   txtfecha1 = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha2 = "to_date('" & Format(mata(i, 11), "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha3 = "to_date('" & Format(mata(i, 13), "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtcadena = "INSERT INTO " & TablaPosDiv & " VALUES("
   txtcadena = txtcadena & tipopos & ","                       'TIPO DE POSICION
   txtcadena = txtcadena & txtfecha1 & ","                     'fecha de la posicion
   txtcadena = txtcadena & "'" & txtnompos & "',"              'nombre de la posicion simulada
   txtcadena = txtcadena & "'000000',"                         'hora de la posicion
   txtcadena = txtcadena & "'" & mata(i, 3) & "',"             'intencion
   txtcadena = txtcadena & mata(i, 2) & ","                    'clave de la posicion
   txtcadena = txtcadena & "'" & mata(i, 4) & "',"             'clave de operacion
   txtcadena = txtcadena & mata(i, 5) & ","                    'tipo de operacion
   txtcadena = txtcadena & "'" & mata(i, 6) & "',"             'tipo valor
   txtcadena = txtcadena & "'" & mata(i, 7) & "',"             'emision
   txtcadena = txtcadena & "'" & mata(i, 8) & "',"             'serie
   txtcadena = txtcadena & "'" & mata(i, 9) & "',"             'clave de emision
   txtcadena = txtcadena & mata(i, 10) & ","                   'no de titulos
   txtcadena = txtcadena & txtfecha2 & ","                     'fecha de compra
   txtcadena = txtcadena & txtfecha3 & ","                     'fecha de liquidacion
   txtcadena = txtcadena & mata(i, 12) & ","                   'precio de compra
   txtcadena = txtcadena & "null,"                              'subport 1
   txtcadena = txtcadena & "null,"                              'subport 2
   txtcadena = txtcadena & "null)"                               'CALIFICACION
   ConAdo.Execute txtcadena
   noreg = noreg + 1
   MensajeProc = "Creando posicion de riesgo de la mesa de cambios "
   DoEvents
 Next i
 If noreg = UBound(mata, 1) Then
 exito = True
 Else
 exito = False
 End If
End Sub

Function CrearFiltroPosPort(ByVal fecha As Date, ByVal txtport As String) As String()
On Error GoTo hayerror

Dim txtfecha As String
Dim mattxt(1 To 5) As String
If Not EsVariableVacia(txtport) Then
    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
  
    mattxt(1) = "SELECT * FROM " & TablaPosMD & " WHERE"
    mattxt(1) = mattxt(1) & " (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN "
    mattxt(1) = mattxt(1) & " (SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM"
    mattxt(1) = mattxt(1) & " " & TablaPortPosicion & "  WHERE PORTAFOLIO = '" & txtport & "' AND FECHA_PORT = " & txtfecha & ")"

    mattxt(2) = "SELECT * FROM " & TablaPosDiv & " WHERE"
    mattxt(2) = mattxt(2) & " ((TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN "
    mattxt(2) = mattxt(2) & "(SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPortPosicion & " "
    mattxt(2) = mattxt(2) & "WHERE PORTAFOLIO = '" & txtport & "' AND FECHA_PORT =" & txtfecha & "))"
    
    mattxt(3) = "SELECT * FROM " & TablaPosSwaps & " WHERE"
    mattxt(3) = mattxt(3) & " (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN "
    mattxt(3) = mattxt(3) & " (SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM "
    mattxt(3) = mattxt(3) & "" & TablaPortPosicion & "  WHERE PORTAFOLIO = '" & txtport & "' AND FECHA_PORT = " & txtfecha & ")"

    mattxt(4) = "SELECT * FROM " & TablaPosFwd & " WHERE"
    mattxt(4) = mattxt(4) & " (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN "
    mattxt(4) = mattxt(4) & " (SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPortPosicion & " "
    mattxt(4) = mattxt(4) & " WHERE PORTAFOLIO = '" & txtport & "' AND FECHA_PORT = " & txtfecha & ")"

    mattxt(5) = "SELECT * FROM " & TablaPosDeuda & " WHERE"
    mattxt(5) = mattxt(5) & " (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN "
    mattxt(5) = mattxt(5) & " (SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM "
    mattxt(5) = mattxt(5) & "" & TablaPortPosicion & "  WHERE PORTAFOLIO = '" & txtport & "' AND FECHA_PORT = " & txtfecha & ")"
Else
   mattxt(1) = ""
   mattxt(2) = ""
   mattxt(3) = ""
   mattxt(4) = ""
   mattxt(5) = ""
End If
CrearFiltroPosPort = mattxt
Exit Function
hayerror:
MsgBox "CrearfiltroPort " & error(Err())
End Function

Function CrearFiltroPosOperPort(ByVal tipopos As Integer, ByVal fechareg As Date, ByVal txtnompos As String, horareg As String, ByVal cposicion As Integer, ByVal coperacion As String) As String()
Dim txtfecha As String
Dim mattxt(1 To 5) As String

If Not EsVariableVacia(tipopos) And Not EsVariableVacia(cposicion) And Not EsVariableVacia(coperacion) Then
    txtfecha = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    mattxt(1) = "SELECT * FROM " & TablaPosMD & " WHERE TIPOPOS = " & tipopos
    mattxt(1) = mattxt(1) & " AND FECHAREG = " & txtfecha
    mattxt(1) = mattxt(1) & " AND NOMPOS = '" & txtnompos & "'"
    mattxt(1) = mattxt(1) & " AND HORAREG = '" & horareg & "'"
    mattxt(1) = mattxt(1) & " AND CPOSICION = " & cposicion
    mattxt(1) = mattxt(1) & " AND COPERACION = '" & coperacion & "'"
    
    mattxt(2) = "SELECT * FROM " & TablaPosDiv & " WHERE TIPOPOS = " & tipopos
    mattxt(2) = mattxt(2) & " AND FECHAREG = " & txtfecha
    mattxt(2) = mattxt(2) & " AND NOMPOS = '" & txtnompos & "'"
    mattxt(2) = mattxt(2) & " AND HORAREG = '" & horareg & "'"
    mattxt(2) = mattxt(2) & " AND CPOSICION = " & cposicion
    mattxt(2) = mattxt(2) & " AND COPERACION = '" & coperacion & "'"
    
    mattxt(3) = "SELECT * FROM " & TablaPosSwaps & " WHERE TIPOPOS = " & tipopos
    mattxt(3) = mattxt(3) & " AND FECHAREG = " & txtfecha
    mattxt(3) = mattxt(3) & " AND NOMPOS = '" & txtnompos & "'"
    mattxt(3) = mattxt(3) & " AND HORAREG = '" & horareg & "'"
    mattxt(3) = mattxt(3) & " AND CPOSICION = " & cposicion
    mattxt(3) = mattxt(3) & " AND COPERACION = '" & coperacion & "'"

    mattxt(4) = "SELECT * FROM " & TablaPosFwd & " where TIPOPOS = " & tipopos
    mattxt(4) = mattxt(4) & " AND FECHAREG = " & txtfecha
    mattxt(4) = mattxt(4) & " AND NOMPOS = '" & txtnompos & "'"
    mattxt(4) = mattxt(4) & " AND HORAREG = '" & horareg & "'"
    mattxt(4) = mattxt(4) & " AND CPOSICION = " & cposicion
    mattxt(4) = mattxt(4) & " AND COPERACION = '" & coperacion & "'"


    mattxt(5) = "SELECT * FROM " & TablaPosDeuda & " WHERE TIPOPOS = " & tipopos
    mattxt(5) = mattxt(5) & " AND FECHAREG = " & txtfecha
    mattxt(5) = mattxt(5) & " AND NOMPOS = '" & txtnompos & "'"
    mattxt(5) = mattxt(5) & " AND HORAREG = '" & horareg & "'"
    mattxt(5) = mattxt(5) & " AND CPOSICION = " & cposicion
    mattxt(5) = mattxt(5) & " AND COPERACION = '" & coperacion & "'"
    
Else
       mattxt(1) = ""
       mattxt(2) = ""
       mattxt(3) = ""
       mattxt(4) = ""
       mattxt(5) = ""
End If
CrearFiltroPosOperPort = mattxt
End Function


Function CrearFiltroPosSim(ByVal txtpossim As String) As String()
Dim mattxt(1 To 5) As String
If Not EsVariableVacia(txtpossim) Then
   mattxt(1) = "SELECT * FROM " & TablaPosMD & " WHERE TIPOPOS = 2 AND NOMPOS = '" & txtpossim & "'"
   mattxt(2) = "SELECT * FROM " & TablaPosDiv & " WHERE TIPOPOS = 2 AND NOMPOS = '" & txtpossim & "'"
   mattxt(3) = "SELECT * FROM " & TablaPosSwaps & " WHERE TIPOPOS = 2 AND NOMPOS = '" & txtpossim & "'"
   mattxt(4) = "SELECT * FROM " & TablaPosFwd & " WHERE TIPOPOS = 2 AND NOMPOS = '" & txtpossim & "'"
   mattxt(5) = "SELECT * FROM " & TablaPosDeuda & " WHERE TIPOPOS = 2 AND NOMPOS = '" & txtpossim & "'"
Else
   mattxt(1) = ""
   mattxt(2) = ""
   mattxt(3) = ""
   mattxt(4) = ""
   mattxt(5) = ""
End If
CrearFiltroPosSim = mattxt
End Function

Function CrearFiltroPosID(ByVal txtpossim As String) As String()
Dim mattxt(1 To 5) As String
If Not EsVariableVacia(txtpossim) Then
   mattxt(1) = "SELECT * FROM " & TablaPosMD & " WHERE TIPOPOS = 3 AND NOMPOS = '" & txtpossim & "'"
   mattxt(2) = "SELECT * FROM " & TablaPosDiv & " WHERE TIPOPOS = 3 AND NOMPOS = '" & txtpossim & "'"
   mattxt(3) = "SELECT * FROM " & TablaPosSwaps & " WHERE TIPOPOS = 3 AND NOMPOS = '" & txtpossim & "'"
   mattxt(4) = "SELECT * FROM " & TablaPosFwd & " WHERE TIPOPOS = 3 AND NOMPOS = '" & txtpossim & "'"
   mattxt(5) = "SELECT * FROM " & TablaPosDeuda & " WHERE TIPOPOS = 3 AND NOMPOS = '" & txtpossim & "'"
Else
   mattxt(1) = ""
   mattxt(2) = ""
   mattxt(3) = ""
   mattxt(4) = ""
   mattxt(5) = ""
End If
CrearFiltroPosID = mattxt
End Function

Sub LeerPosBDatos(ByRef mattxt1() As String, _
                  ByRef matpos() As propPosRiesgo, _
                  ByRef matposmd() As propPosMD, _
                  ByRef matposdiv() As propPosDiv, _
                  ByRef matposswap() As propPosSwaps, _
                  ByRef matposfwd() As propPosFwd, _
                  ByRef matflswap() As estFlujosDeuda, _
                  ByRef matposdeuda() As propPosDeuda, _
                  ByRef matfldeuda() As estFlujosDeuda, _
                  ByRef txtmsg As String, ByRef exito As Boolean)

If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim txtmsg1 As String
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim exito4 As Boolean
exito = True
'posicion de mercado de dinero
   matposmd = LeerBaseMD(mattxt1(1))
   If Not EsArrayVacio(matposmd) Then
      Call LeerFlujosEmMD(matposmd, MatFlujosMD, txtmsg1, exito2)
      exito = exito And exito2
   End If
'posicion de cambios
   matposdiv = LeerPosDiv(mattxt1(2))
'posicion de swaps
   Call CrearPosSwaps(mattxt1(3), matposswap, matflswap, exito3)
   If Not EsArrayVacio(matposswap) Then exito = exito And exito3
'posicion de forwards
   matposfwd = LeerTablaPosFwd(mattxt1(4))
 'posiciones de deuda
   Call LeerPosDeuda(mattxt1(5), matposdeuda, matfldeuda, exito4)
   If Not EsArrayVacio(matposdeuda) Then exito = exito And exito4
'se actualiza la tasa cupon de los flujos de swaps
matpos = ConsolidarPosiciones(matposmd, matposdiv, matposswap, matposfwd, matposdeuda)
If Not exito Then
  txtmsg = txtmsg1
End If
If EsArrayVacio(matpos) Then
   exito = False
End If
On Error GoTo 0
Exit Sub
hayerror:
    MsgBox "LeerPosBDatos " & error(Err())
    exito = False
End Sub

Sub DefinirParValPos(ByRef matpos() As propPosRiesgo, _
                     ByRef matposmd() As propPosMD, _
                     ByRef matposdiv() As propPosDiv, _
                     ByRef matposswaps() As propPosSwaps, _
                     ByRef matposfwd() As propPosFwd, _
                     ByRef matposdeuda() As propPosDeuda, _
                     ByVal tval As Integer, _
                     ByRef txtmsg As String, _
                     ByRef exito As Boolean)

On Error GoTo hayerror
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim exito4 As Boolean
Dim exito5 As Boolean
Dim txtmsg1 As String, txtmsg2 As String, txtmsg3 As String, txtmsg4 As String, txtmsg5 As String
exito1 = True: exito2 = True: exito3 = True: exito4 = True: exito5 = True
'posicion de mercado de dinero
   If Not EsArrayVacio(matposmd) Then
      Call ClasTProdMD(matpos, matposmd, txtmsg1, exito1)
   End If
'posicion de cambios
   If Not EsArrayVacio(matposdiv) Then
      Call ClasTprod2(matpos, matposdiv, txtmsg2, exito2)
   End If
'posicion de swaps
   If Not EsArrayVacio(matposswaps) Then
      If tval = 1 Then
         Call AgregarPValSwaps(matpos, matposswaps, MatTValSwaps1, txtmsg3, exito3)
      Else
         Call AgregarPValSwaps(matpos, matposswaps, MatTValSwaps2, txtmsg3, exito3)
      End If
   End If
'posicion de forwards
   If Not EsArrayVacio(matposfwd) Then
      If tval = 1 Then
         Call AgregarPValFwd(matpos, matposfwd, MatTValFwdsTC1, txtmsg4, exito4)
      Else
         Call AgregarPValFwd(matpos, matposfwd, MatTValFwdsTC2, txtmsg4, exito4)
      End If
   End If
 'posiciones de deuda
   If Not EsArrayVacio(matposdeuda) Then
      Call AgregarPValDeuda(matpos, matposdeuda, MatTValDeuda, txtmsg5, exito5)
   End If
   exito = exito1 And exito2 And exito3 And exito4 And exito5
   If exito Then
      txtmsg = "El proceso finalizo correctamente"
   Else
      txtmsg = txtmsg1 & txtmsg2 & txtmsg3 & txtmsg4 & txtmsg5
   End If
   Exit Sub
hayerror:
MsgBox "DefinirParValPos " & error(Err())
End Sub

Function ListaOpPendientes(ByVal fecha As Date) As Variant()
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg1 As Long
Dim noreg2 As Long
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT COPERACION FROM " & TablaPosSwaps & " WHERE FECHAREG = " & txtfecha & " AND TIPOPOS =3 AND INTENCION ='C' AND COPERACION , HORAREG  NOT IN(SELECT CAST(COPERACION AS VARCHAR2(12)), CAST(HORAREGIKOS AS VARCHAR2(12)) FROM " & TablaOperValidada & ")"
txtfiltro1 = "SELECT COUNT (*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg1 = rmesa.Fields(0)
rmesa.Close
If noreg1 <> 0 Then
   ReDim mata1(1 To noreg1, 1 To 1) As Variant
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg1
       mata1(i, 1) = rmesa.Fields(0)
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata1(0 To 0, 0 To 0) As Variant
End If
txtfiltro2 = "SELECT COPERACION FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfecha & " AND TIPOPOS ='3' AND INTENCION ='C' AND (CAST(COPERACION AS VARCHAR2(12)), CAST(HORAPOS AS VARCHAR2(12))) NOT IN(SELECT CAST(COPERACION AS VARCHAR2(12)), CAST(HORAREGIKOS AS VARCHAR2(12)) FROM " & TablaOperValidada & ")"
txtfiltro1 = "SELECT COUNT (*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg2 = rmesa.Fields(0)
rmesa.Close
If noreg2 <> 0 Then
   ReDim mata2(1 To noreg2, 1 To 1) As Variant
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg2
       mata2(i, 1) = rmesa.Fields(0)
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata2(0 To 0, 0 To 0) As Variant
End If
ListaOpPendientes = UnirTablas(mata1, mata2, 1)
End Function

Function ConsolidarPosiciones(ByRef matpos1() As propPosMD, ByRef matpos2() As propPosDiv, ByRef matpos3() As propPosSwaps, ByRef matpos4() As propPosFwd, ByRef matpos5() As propPosDeuda)
On Error GoTo hayerror
Dim contar As Long
Dim i As Long
Dim matpos() As New propPosRiesgo

contar = 0
ReDim matpos(1 To 1)
If Not EsArrayVacio(matpos1) Then
If UBound(matpos1, 1) <> 0 Then
   For i = 1 To UBound(matpos1, 1)
       contar = contar + 1
       ReDim Preserve matpos(1 To contar)
       matpos(contar).IndPosicion = i                               'lugar en el array
       matpos(contar).No_tabla = 1                                     'tabla 1
       matpos(contar).tipopos = matpos1(i).tipopos               'tipo de la posicion
       matpos(contar).nompos = matpos1(i).nompos                  'nombre de la posicion
       matpos(contar).C_Posicion = matpos1(i).C_Posicion            'clave de la posicion
       matpos(contar).fechareg = matpos1(i).fechareg             'fecha de registro
       matpos(contar).HoraRegOp = matpos1(i).HoraRegOp             'fecha de registro
       matpos(contar).c_operacion = matpos1(i).c_operacion          'clave de operacion
       matpos(contar).intencion = matpos1(i).intencion            'intencion
       matpos(contar).Tipo_Mov = matpos1(i).Tipo_Mov                'tipo de operacion
       matpos(contar).Signo_Op = matpos1(i).Signo_Op
   Next i
End If
End If
If UBound(matpos2, 1) <> 0 Then
   For i = 1 To UBound(matpos2, 1)
       contar = contar + 1
       ReDim Preserve matpos(1 To contar)
       matpos(contar).IndPosicion = i                                'lugar en el array
       matpos(contar).No_tabla = 2                                     'tabla 2
       matpos(contar).tipopos = matpos2(i).tipopos                 'tipo de la posicion
       matpos(contar).nompos = matpos2(i).nompos                  'nombre de la posicion
       matpos(contar).C_Posicion = matpos2(i).C_Posicion             'clave de la posicion
       matpos(contar).fechareg = matpos2(i).fechareg               'fecha de registro
       matpos(contar).HoraRegOp = matpos2(i).HoraRegOp             'fecha de registro
       matpos(contar).c_operacion = matpos2(i).c_operacion           'clave de operacion
       matpos(contar).intencion = matpos2(i).intencion           'intencion
       matpos(contar).Tipo_Mov = matpos2(i).Tipo_Mov                 'tipo de operacion
       matpos(contar).Signo_Op = matpos2(i).Signo_Op
   Next i
End If
If Not EsArrayVacio(matpos3) Then
   For i = 1 To UBound(matpos3, 1)
       contar = contar + 1
       ReDim Preserve matpos(1 To contar)
       matpos(contar).IndPosicion = i                               'lugar en el array
       matpos(contar).No_tabla = 3                                    'tabla 3
       matpos(contar).tipopos = matpos3(i).tipopos               'tipo de la posicion
       matpos(contar).nompos = matpos3(i).nompos                 'nombre de la posicion
       matpos(contar).C_Posicion = matpos3(i).C_Posicion             'clave de la posicion
       matpos(contar).fechareg = matpos3(i).fechareg               'fecha de registro
       matpos(contar).HoraRegOp = matpos3(i).HoraRegOp             'fecha de registro
       matpos(contar).c_operacion = matpos3(i).c_operacion          'clave de operacion
       matpos(contar).intencion = matpos3(i).intencion            'intencion
       matpos(contar).Tipo_Mov = matpos3(i).Tipo_Mov                 'tipo de operacion
       matpos(contar).Signo_Op = matpos3(i).Signo_Op
   Next i
End If
If UBound(matpos4, 1) <> 0 Then
   For i = 1 To UBound(matpos4, 1)
       contar = contar + 1
       ReDim Preserve matpos(1 To contar)
       matpos(contar).IndPosicion = i                               'lugar en el array
       matpos(contar).No_tabla = 4                                      'tabla 4
       matpos(contar).tipopos = matpos4(i).tipopos                 'tipo de la posicion
       matpos(contar).nompos = matpos4(i).nompos                  'nombre de la posicion
       matpos(contar).C_Posicion = matpos4(i).C_Posicion            'clave de la posicion
       matpos(contar).fechareg = matpos4(i).fechareg              'fecha de registro
       matpos(contar).HoraRegOp = matpos4(i).HoraRegOp            'fecha de registro
       matpos(contar).c_operacion = matpos4(i).c_operacion           'clave de operacion
       matpos(contar).intencion = matpos4(i).intencion           'intencion
       matpos(contar).Tipo_Mov = matpos4(i).Tipo_Mov                 'tipo de operacion
       matpos(contar).Signo_Op = matpos4(i).Signo_Op
   Next i
End If
If UBound(matpos5, 1) <> 0 Then
   For i = 1 To UBound(matpos5, 1)
       contar = contar + 1
       ReDim Preserve matpos(1 To contar)
       matpos(contar).IndPosicion = i                                'lugar en el array
       matpos(contar).No_tabla = 5                                      'tabla 5
       matpos(contar).tipopos = matpos5(i).tipopos                 'tipo de la posicion
       matpos(contar).nompos = matpos5(i).nompos                  'nombre de la posicion
       matpos(contar).C_Posicion = matpos5(i).C_Posicion             'clave de la posicion
       matpos(contar).fechareg = matpos5(i).fechareg               'fecha de registro
       matpos(contar).HoraRegOp = matpos5(i).HoraRegOp            'fecha de registro
       matpos(contar).c_operacion = matpos5(i).c_operacion           'clave de operacion
       matpos(contar).intencion = matpos5(i).intencion            'intencion
       matpos(contar).Tipo_Mov = matpos5(i).Tipo_Mov                 'tipo de operacion
       matpos(contar).Signo_Op = matpos5(i).Signo_Op
   Next i
End If
If contar = 0 Then
ReDim matpos(0 To 0) As New propPosRiesgo
End If
ConsolidarPosiciones = matpos
On Error GoTo 0
Exit Function
hayerror:
  MsgBox "consolidarposiciones " & error(Err())
End Function

Sub CrearPosSwaps(ByVal txtfiltro1 As String, ByRef matpos() As propPosSwaps, ByRef matflujos() As estFlujosDeuda, ByRef exito As Boolean)
If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim exito1 As Boolean
'se lee la posicion de swaps de la fecha
exito = True
    matpos = LeerTablaSwaps(txtfiltro1)
    If UBound(matpos, 1) > 0 Then
       matflujos = FiltrarFlujosSwaps3(matpos, True, exito1)
       exito = exito And exito1

    Else
       ReDim matflujos(0 To 0)
       exito = False
    End If
On Error GoTo 0
Exit Sub
hayerror:
  MsgBox "crearposswaps " & error(Err())
End Sub

Sub LeerPosDeuda(ByVal txtfiltro1 As String, ByRef matpos() As propPosDeuda, ByRef matflujos() As estFlujosDeuda, ByRef exito As Boolean)
If ActivarControlErrores Then
On Error GoTo hayerror
End If
    matpos = LeerTablaDeuda(txtfiltro1)
    If UBound(matpos, 1) > 0 Then
       matflujos = FiltrarFlujosDeuda(matpos, exito)   'filtra en funcion de matpos
    End If
  On Error GoTo 0
Exit Sub
hayerror:
MsgBox "Leerposdeuda " & error(Err())
End Sub

Function DetPosSwapsValContrap(ByVal fecha As Date, ByRef matpos() As Variant, ByVal id_contrap As Integer)
Dim matem() As Variant
Dim noreg As Long
Dim nocampos As Long
Dim indmax As Long
Dim i As Long
Dim j As Long
Dim kk As Long
Dim contar As Long
Dim contar2 As Long

'esta rutina debe de filtrar las operaciones que componen la posicion valida
'esta es la prioridad de filtros
'identificar las operaciones
'determinar para cada operacion cuantos fecha de registro se han hecho
'filtrar las operaciones con fecha de registro valida
'filtrar las operaciones con fecha de inicio y final dentro del rango
'filtrar las operaciones de negociacion o cobertura
matem = ObtFactUnicos(matpos, 8)     'SE OBTIENEN LAS EMISIONES UNICAS DE LA POSICION
noreg = UBound(matem, 1)
nocampos = UBound(matpos, 2)
ReDim matcont(1 To noreg) As Integer
ReDim matfechas(1 To noreg, 1 To 1) As Variant
ReDim matind1(1 To noreg, 1 To 1) As Variant
ReDim matind(1 To noreg) As Integer
indmax = 0
For i = 1 To noreg
   For j = 1 To UBound(matpos, 1)
       If matpos(j, 8) = matem(i, 1) Then             'cuenta los registros de la misma emision
          matcont(i) = matcont(i) + 1
          indmax = Maximo(indmax, matcont(i))
          ReDim Preserve matfechas(1 To noreg, 1 To indmax) As Variant
          ReDim Preserve matind1(1 To noreg, 1 To indmax) As Variant
          matfechas(i, matcont(i)) = matpos(j, 2)     'FECHA DE REGISTRO
          matind1(i, matcont(i)) = j                  'registro de la posicion
       End If
   Next j
Next i
For i = 1 To noreg
      ReDim matf(1 To matcont(i), 1 To 2) As Variant
      For j = 1 To matcont(i)
          matf(j, 1) = matfechas(i, j)
          matf(j, 2) = matind1(i, j)
      Next j
      matf = RutinaOrden(matf, 1, SRutOrden)
      If fecha < matf(1, 1) Then  'menor que la primer fecha de registro, entonces no existe
         matind(i) = 0
      ElseIf fecha >= matf(matcont(i), 1) Then
         matind(i) = matf(matcont(i), 2)
      ElseIf matcont(i) <> 1 Then
         For j = 2 To matcont(i)
         If matf(j - 1, 1) <= fecha And fecha < matf(j, 1) Then
            matind(i) = matf(j - 1, 2)
            Exit For
         End If
         Next j
      End If
Next i
contar = 0
ReDim matpos1(1 To nocampos, 0 To 0) As Variant

     For i = 1 To noreg
     If matind(i) <> 0 Then
     'se valida que la fecha de la posicion sean entre la fecha de inicio y la fecha de fin del swap
        If fecha >= matpos(matind(i), 10) And fecha < matpos(matind(i), 11) Then
          contar = contar + 1
          ReDim Preserve matpos1(1 To nocampos, 0 To contar) As Variant
          For kk = 1 To nocampos
            matpos1(kk, contar) = matpos(matind(i), kk)
          Next kk
        End If
     End If
     Next i
contar2 = 0
ReDim matpos2(1 To nocampos, 0 To contar2) As Variant
For i = 1 To contar
If matpos1(25, i) = id_contrap Then
contar2 = contar2 + 1
ReDim Preserve matpos2(1 To nocampos, 0 To contar2) As Variant
   For kk = 1 To nocampos
       matpos2(kk, contar2) = matpos1(kk, i)
   Next kk
End If
Next i
If contar2 <> 0 Then
matpos2 = MTranV(matpos2)
Else
 ReDim matpos2(0 To 0, 0 To 0) As Variant
End If
DetPosSwapsValContrap = matpos2
End Function

Function DetPosSwapsPrimValida(ByVal fecha As Date, ByRef matpos() As Variant)
Dim matem() As Variant
Dim noreg As Long
Dim nocampos As Long
Dim indmax As Long
Dim i As Long
Dim j As Long
Dim contar As Long
Dim kk As Long

'esta rutina debe de filtrar las operaciones que componen la posicion valida
'esta es la prioridad de filtros
'identificar las operaciones
'determinar para cada operacion cuantos fecha de registro se han hecho
'filtrar las operaciones con fecha de registro valida
'filtrar las operaciones con fecha de inicio y final dentro del rango
'filtrar las operaciones de negociacion o cobertura
matem = ObtFactUnicos(matpos, 2)
noreg = UBound(matem, 1)
nocampos = UBound(matpos, 2)
ReDim matcont(1 To noreg) As Integer
ReDim matfechas(1 To noreg, 1 To 1) As Variant
ReDim matind1(1 To noreg, 1 To 1) As Variant
ReDim matind(1 To noreg) As Integer
indmax = 0
For i = 1 To UBound(matem, 1)
   For j = 1 To UBound(matpos, 1)
      If matpos(j, 2) = matem(i, 1) Then   'cuenta los registros de la misma emision
         matcont(i) = matcont(i) + 1
         indmax = Maximo(indmax, matcont(i))
         ReDim Preserve matfechas(1 To noreg, 1 To indmax) As Variant
         ReDim Preserve matind1(1 To noreg, 1 To indmax) As Variant
         matfechas(i, matcont(i)) = matpos(j, 1)
         matind1(i, matcont(i)) = j
         
      End If
   Next j
Next i

For i = 1 To noreg
   For j = 1 To matcont(i)
      If fecha < matfechas(i, 1) Then
         matind(i) = 0
         Exit For
      ElseIf fecha >= matfechas(i, matcont(i)) Then
         matind(i) = j
         Exit For
      ElseIf matcont(i) <> 1 Then
         If matfechas(i, j - 1) <= fecha And fecha < matfechas(i, j) Then
            matind(i) = j - 1
            Exit For
         End If
      End If
   Next j
Next i
contar = 0

ReDim matpos1(1 To nocampos, 1 To 1) As Variant
For i = 1 To noreg
     If matind(i) <> 0 Then
        If fecha >= matpos(matind1(i, matind(i)), 4) And fecha < matpos(matind1(i, matind(i)), 5) Then
          contar = contar + 1
          ReDim Preserve matpos1(1 To nocampos, 1 To contar) As Variant
          For kk = 1 To nocampos
            matpos1(kk, contar) = matpos(matind1(i, matind(i)), kk)
          Next kk
        End If
     End If
Next i
matpos1 = MTranV(matpos1)
DetPosSwapsPrimValida = matpos1
End Function

Sub ValidarActRealizadaF(ByVal fecha As Date, ByVal indice As Long, ByVal opcion As Integer, ByVal txtmsg As String, ByVal exito As Boolean)
Dim txtfecha As String
Dim txthoy As String
Dim txthora As String
Dim txtactualiza As String
Dim noreg As Long
Dim txttabla As String
txtmsg = ReemplazaCadenaTexto(txtmsg, "'", "")
    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txthoy = "to_date('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txthora = "to_date('" & Format$(Time, "hh:mm:ss") & "','HH24:MI:SS')"
    txtactualiza = "UPDATE " & DetermTablaProc(opcion) & " SET FINALIZADO = '" & ConvBolStr(exito) & "', BLOQUEADO = 'N', FFINAL = " & txthoy & ", HFINAL = " & txthora & ", COMENTARIO = '" & Left(txtmsg, 150) & "', USUARIO = '" & NomUsuario & "' WHERE FECHAP = " & txtfecha & "AND ID_TAREA = " & indice
    ConAdo.Execute txtactualiza, noreg
End Sub

Sub ValidarActRealizada(ByVal idfolio As Long, ByVal opcion As Integer, ByVal txtmsg As String, ByVal exito As Boolean)
Dim txtactualiza As String
Dim txthoy As String
Dim txthora As String
Dim noreg As Long
Dim txttabla As String
txttabla = DetermTablaProc(opcion)

    txthoy = "to_date('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txthora = "to_date('" & Format$(Time, "hh:mm:ss") & "','HH24:MI:SS')"
    txtactualiza = "UPDATE " & txttabla & " SET FINALIZADO = '" & ConvBolStr(exito) & "', BLOQUEADO = 'N', FFINAL = " & txthoy & ", HFINAL = " & txthora & ", COMENTARIO = '" & txtmsg & "', USUARIO = '" & NomUsuario & "' WHERE ID_FOLIO = " & idfolio
    ConAdo.Execute txtactualiza, noreg
End Sub

Function ValidarEjecucionProceso(ByVal id_tarea As Integer, ByVal fecha As Date, ByRef matproceso() As Variant, ByRef matssubproc() As Variant, ByVal opcion As Integer)
Dim contar1 As Integer
Dim contar2 As Integer
Dim i As Integer
Dim j As Integer
Dim txtfecha As String
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim txtfiltro As String
Dim mata() As Variant
Dim matb() As Variant
Dim nocols1 As Integer
Dim nocols2 As Integer
Dim exito As Boolean
Dim exito1 As Boolean
Dim exito2 As Boolean

nocols1 = UBound(matproceso, 2)
nocols2 = UBound(matssubproc, 2)

   contar1 = 0
   contar2 = 0
   'se lee la lista de procesos que preceden al id_tarea actual
    ReDim mata(1 To nocols1, 1 To 1) As Variant
    ReDim matb(1 To nocols2, 1 To 1) As Variant
    noreg1 = UBound(matproceso, 1)
    For i = 1 To noreg1
        If matproceso(i, 1) = id_tarea Then  'si precede al id_tarea
           contar1 = contar1 + 1
           ReDim Preserve mata(1 To nocols1, 1 To contar1) As Variant
           For j = 1 To nocols1
             mata(j, contar1) = matproceso(i, j)
           Next j
        End If
    Next i
    noreg2 = UBound(matssubproc, 1)
    For i = 1 To noreg2
        If matssubproc(i, 1) = id_tarea Then  'si precede al id_tarea
           contar2 = contar2 + 1
           ReDim Preserve matb(1 To nocols2, 1 To contar2) As Variant
           For j = 1 To nocols2
             matb(j, contar2) = matssubproc(i, j)
           Next j
        End If
    Next i
    If contar1 > 0 Then
       mata = MTranV(mata)
       exito1 = True
       For i = 1 To contar1
           exito = validarcondProc(fecha, mata, i, opcion)
           exito1 = exito1 And exito
       Next i
    Else
      exito1 = True
    End If
    If contar2 > 0 Then
       matb = MTranV(matb)
       exito2 = True
       For i = 1 To contar2
           exito = validarcondSubProc(fecha, matb, i, opcion)
           exito2 = exito2 And exito
       Next i
    Else
      exito2 = True
    End If
ValidarEjecucionProceso = exito1 And exito2
End Function

Function validarcondProc(ByVal fecha As Date, ByRef matparam() As Variant, ByVal indice As Integer, ByVal opcion As Integer)
Dim txtfecha As String
Dim txtfiltro As String
Dim i As Integer
Dim noreg As Long
Dim txttabla As String
Dim rmesa As New ADODB.recordset

    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro = "SELECT COUNT(*) FROM " & DetermTablaProc(opcion) & " WHERE FECHAP = " & txtfecha & " and ID_TAREA = " & matparam(indice, 3) & " and FINALIZADO = 'S'"
    rmesa.Open txtfiltro, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg = 0 Then 'no se ha ejecutado el proceso i
       validarcondProc = False
    Else
       validarcondProc = True
    End If
End Function

Function validarcondSubProc(ByVal fecha As Date, ByRef matparam() As Variant, ByVal indice As Integer, ByVal opcion As Integer)
Dim txtfecha As String
Dim txtfiltro As String
Dim i As Integer
Dim noreg1 As Long
Dim noreg2 As Long
Dim noreg3 As Long
Dim txttabla As String
Dim rmesa As New ADODB.recordset

txttabla = DetermTablaSubproc(opcion)

    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro = "SELECT COUNT(*) FROM " & txttabla & " WHERE FECHAP = " & txtfecha & " and ID_SUBPROCESO = " & matparam(indice, 3)
    For i = 1 To 12
        If Not EsVariableVacia(matparam(indice, i + 3)) Then
           txtfiltro = txtfiltro & " AND PARAMETRO" & i & " = '" & matparam(indice, i + 3) & "'"
        End If
    Next i
    rmesa.Open txtfiltro, ConAdo
    noreg1 = rmesa.Fields(0)
    rmesa.Close
    txtfiltro = "SELECT COUNT(*) FROM " & txttabla & " WHERE FECHAP = " & txtfecha & " and ID_SUBPROCESO = " & matparam(indice, 3) & " and FINALIZADO = 'N'"
    For i = 1 To 12
        If Not EsVariableVacia(matparam(indice, i + 3)) Then
           txtfiltro = txtfiltro & " AND PARAMETRO" & i & " = '" & matparam(indice, i + 3) & "'"
        End If
    Next i
    rmesa.Open txtfiltro, ConAdo
    noreg2 = rmesa.Fields(0)
    rmesa.Close
    txtfiltro = "SELECT COUNT(*) FROM " & txttabla & " WHERE FECHAP = " & txtfecha & " and ID_SUBPROCESO = " & matparam(indice, 3) & " and FINALIZADO = 'S' AND EXITO <> 'S'"
    For i = 1 To 12
        If Not EsVariableVacia(matparam(indice, i + 3)) Then
           txtfiltro = txtfiltro & " AND PARAMETRO" & i & " = '" & matparam(indice, i + 3) & "'"
        End If
    Next i
    rmesa.Open txtfiltro, ConAdo
    noreg3 = rmesa.Fields(0)
    rmesa.Close
    If noreg1 <> 0 And noreg2 = 0 And noreg3 = 0 Then 'se terminaron todos los subprocesos
       validarcondSubProc = True
    Else
       validarcondSubProc = False
    End If

End Function

Sub BloquearProcesoF(ByVal fecha As Date, ByVal proceso As Long, ByVal sirepbloq As Boolean, ByVal opcion As Integer, ByRef exito As Boolean)
On Error GoTo hayerror
'bloquear proceso
Dim txtfecha As String
Dim txtfiltro As String
Dim finicio As String
Dim hinicio As String
Dim txtipdir As String
Dim txtcadena As String
Dim noreg As Long
Dim txttabla As String
txttabla = DetermTablaProc(opcion)
    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    finicio = "TO_DATE('" & Format$(Date, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    hinicio = "TO_DATE('" & Format$(Time, "HH:MM:SS") & "','HH24:MI:SS')"
    txtipdir = RecuperarIP
    If sirepbloq Then
       txtcadena = "UPDATE " & txttabla & " SET BLOQUEADO = 'S', USUARIO = '" & NomUsuario & "', FINICIAL = " & finicio & ", HINICIAL = " & hinicio & ", IP_DIRECCION = '" & txtipdir & "' WHERE ID_TAREA = " & proceso & " AND FECHAP =" & txtfecha & " AND BLOQUEADO = 'N' AND FINALIZADO = 'N'"
    Else
       'no respeta el si ya se realizaron o no
       txtcadena = "UPDATE " & txttabla & " SET BLOQUEADO = 'S', USUARIO = '" & NomUsuario & "', FINICIAL = " & finicio & ", HINICIAL = " & hinicio & ", IP_DIRECCION = '" & txtipdir & "' WHERE ID_TAREA = " & proceso & " AND FECHAP =" & txtfecha & " AND BLOQUEADO = 'N'"
    End If
    ConAdo.Execute txtcadena, noreg
    If noreg <> 0 Then
       exito = True
    Else
       exito = False
    End If
On Error GoTo 0
Exit Sub
hayerror:
   exito = False
End Sub

Function ConvBolStr(ByVal entrada As Boolean)
    If entrada Then
        ConvBolStr = "S"
    Else
        ConvBolStr = "N"
    End If
End Function

Function ObtProcesosFecha(ByVal fecha As Date, ByVal opcion As Integer)
Dim i As Integer
Dim j As Integer
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfecha As String
Dim noreg As Integer
Dim nocampos As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT * FROM " & DetermTablaProc(opcion) & " WHERE FECHAP = " & txtfecha & "  ORDER BY ID_TAREA"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   For i = 1 To noreg
      For j = 1 To nocampos
          mata(i, j) = rmesa.Fields(j - 1)
      Next j
      rmesa.MoveNext
      MensajeProc = "Leyendo lista de procesos..."
   Next i
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
ObtProcesosFecha = mata
End Function

Function Obt1ProcPendFecha(ByVal fecha As Date, ByVal opcion As Integer)
Dim i As Integer
Dim j As Integer
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfecha As String
Dim noreg As Integer
Dim nocampos As Integer
Dim txttabla As String
Dim rmesa As New ADODB.recordset

If opcion = 1 Then
   txttabla = TablaProcesos1
ElseIf opcion = 2 Then
   txttabla = TablaProcesos2
End If

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT * FROM " & txttabla & " WHERE FECHAP = " & txtfecha & " AND FINALIZADO ='N' AND BLOQUEADO = 'N' ORDER BY ID_TAREA"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To nocampos) As Variant
   For j = 1 To nocampos
       mata(j) = rmesa.Fields(j - 1)
   Next j
   MensajeProc = "Leyendo lista de procesos..."
   rmesa.Close
Else
ReDim mata(0 To 0) As Variant
End If
Obt1ProcPendFecha = mata
End Function

Function ObtProcPendFecha(ByVal fecha As Date, ByVal opcion As Integer)
Dim i As Integer
Dim j As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Integer
Dim nocampos As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & DetermTablaProc(opcion) & " WHERE FECHAP = " & txtfecha & " AND FINALIZADO ='N' AND BLOQUEADO = 'N' ORDER BY ID_TAREA"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   For i = 1 To noreg
       For j = 1 To nocampos
           mata(i, j) = rmesa.Fields(j - 1)
       Next j
       rmesa.MoveNext
       MensajeProc = "Leyendo lista de procesos..."
   Next i
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
ObtProcPendFecha = mata
End Function


Function UnirParametros(ByVal fecha As Date, ByRef matpar() As Variant, ByRef matpr() As Variant, ByRef mListaP() As Integer, ByRef exito As Boolean)
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim nocampos1 As Integer
Dim nocampos2 As Integer
Dim i As Integer
Dim j As Integer
Dim contar As Long

'esta es la sefuencia final que debe quedar en la tabla de parametrizacion
'1 id del proceso
'2 parametro 1
'3 parametro 2
'4 parametro 3
'5 parametro 4
'6 parametro 5
'7 parametro 6
'8 parametro 7
'9 parametro 8
'10 parametro 9
'11 parametro 10
'12 fecha de proceso

noreg1 = UBound(matpar, 1)
noreg2 = UBound(matpr, 1)
nocampos1 = UBound(matpar, 2)
nocampos2 = UBound(matpr, 2)
If noreg1 = noreg2 And noreg1 <> 0 And noreg2 <> 0 Then
   ReDim mata(1 To noreg1, 1 To nocampos2) As Variant
   For i = 1 To noreg1
       mata(i, 1) = matpr(i, 1)  'folio
       mata(i, 2) = matpr(i, 2)  'id proceso
       mata(i, 3) = matpr(i, 3)  'descripcion
       For j = 4 To nocampos1 + 3
           If Not EsVariableVacia(matpar(i, j - 3)) Then
              mata(i, j) = matpar(i, j - 3)
           Else
              mata(i, j) = matpr(i, j)
           End If
       Next j
       For j = nocampos1 + 4 To nocampos2
           mata(i, j) = matpr(i, j)
       Next j
   Next i
   contar = 0
   ReDim matb(1 To UBound(mListaP, 1), 1 To nocampos2) As Variant
   For i = 1 To UBound(mListaP, 1)
       For j = 1 To nocampos2
           matb(i, j) = mata(mListaP(i), j)
       Next j
   Next i
   exito = True
Else
   MsgBox "No estan coincidiento los procesos manuales y la tabla automatica"
   exito = False
   ReDim matb(0 To 0, 0 To 0) As Variant
End If
UnirParametros = matb
End Function

Sub SecProcesosManual1(ByVal fecha As Date, ByVal accion1 As Boolean, ByVal accion2 As Boolean, ByVal opcion As Integer)
Dim i As Integer
Dim id_tarea As Integer
Dim id_proc As String
Dim sibloqueo As Boolean
Dim bl_exito As Boolean
Dim exito As Boolean
Dim txtmsg As String
Dim parproc() As Variant
Dim matproc() As Variant
Dim accion As Boolean

   matproc = Obt1ProcPendFecha(fecha, opcion)
   If UBound(matproc, 1) > 0 Then
      id_tarea = matproc(1)
      id_proc = matproc(2)
      accion = True
      accion = ValidarEjecucionProceso(id_tarea, fecha, MatSecProcesos, MatSecSubproc, opcion)
      If accion Or accion1 Then
         Call BloquearProcesoF(fecha, id_tarea, accion2, opcion, sibloqueo)
         If sibloqueo Then
           'se extraen los parametros para el proceso a realizar
            parproc = ObParamProc1(matproc)
            Call ListaProcesos(fecha, id_proc, parproc, opcion, txtmsg, bl_exito)
            Call ValidarActRealizadaF(fecha, id_tarea, opcion, txtmsg, bl_exito)
            DoEvents
            If bl_exito Then
               MensajeProc = "Se termino el proceso " & matproc(3) & " del " & fecha
               Call GuardaDatosBitacora(2, "Proceso", id_tarea, matproc(3), NomUsuario, fecha, MensajeProc, opcion)
            Else
               MensajeProc = "No se realizo el proceso " & matproc(3) & " del " & fecha
            End If
         End If
      End If
   End If

End Sub

Sub SecProcesosManual2(ByVal fecha As Date, ByVal accion1 As Boolean, ByVal accion2 As Boolean, ByVal id_tabla As Integer)
Dim i As Integer
Dim id_tarea As Integer
Dim id_proc As String
Dim sibloqueo As Boolean
Dim bl_exito As Boolean
Dim exito As Boolean
Dim txtmsg As String
Dim parproc() As Variant
Dim matproc() As Variant
Dim accion As Boolean

   matproc = ObtProcPendFecha(fecha, id_tabla)
   If UBound(matproc, 1) > 0 Then
      For i = 1 To UBound(matproc, 1)
          id_tarea = matproc(i, 1)
          id_proc = matproc(i, 2)
          accion = True
          accion = ValidarEjecucionProceso(id_tarea, fecha, MatSecProcesos, MatSecSubproc, id_tabla)
          If accion Or accion1 Then
             Call BloquearProcesoF(fecha, id_tarea, accion2, id_tabla, sibloqueo)
             If sibloqueo Then
               'se extraen los parametros para el proceso a realizar
                parproc = ObParamProc2(matproc, i)
                Call ListaProcesos(fecha, id_proc, parproc, id_tabla, txtmsg, bl_exito)
                Call ValidarActRealizadaF(fecha, id_tarea, id_tabla, txtmsg, bl_exito)
                DoEvents
                If bl_exito Then
                   MensajeProc = "Se termino el proceso " & matproc(i, 3) & " del " & fecha
                   Call GuardaDatosBitacora(2, "Proceso", id_tarea, matproc(i, 3), NomUsuario, fecha, MensajeProc, id_tabla)
                Else
                   MensajeProc = "No se realizo el proceso " & matproc(i, 3) & " del " & fecha
                End If
             End If
          End If
      Next i
   End If

End Sub

Function ObParamProc1(ByRef matpar() As Variant)
Dim i As Integer
ReDim mata(1 To 10) As Variant
For i = 1 To 10
   mata(i) = matpar(2 * i + 3)
Next i
ObParamProc1 = mata
End Function

Function ObParamProc2(ByRef matpar() As Variant, ByVal ind As Long)
Dim i As Integer
ReDim mata(1 To 10) As Variant
For i = 1 To 10
   mata(i) = matpar(ind, 2 * i + 3)
Next i
ObParamProc2 = mata
End Function

Sub SecProcesosAuto(ByVal accion1 As Boolean, ByVal accion2 As Boolean, ByVal opcion As Integer)
    Dim parproc() As Variant
    Dim fecha As Date
    Dim sibloqueo As Boolean
    Dim txtmsg As String
    Dim exito As Boolean
    Dim matpr() As Variant
    Dim mattareasp() As Variant
    Dim i As Long
    Dim id_tarea As Integer
    Dim id_proc As String
    Dim accion As Boolean
  
      'se obtiene la lista de procesos pendientes
      matpr = ObtenerProcesosPendientes(opcion)
      For i = 1 To UBound(matpr, 1)
          id_tarea = matpr(i, 1)      'FOLIO
          id_proc = matpr(i, 2)       'ID DE PROCESO
          fecha = matpr(i, 24)        'fecha del proceso
          'se revisa si el proceso se puede ejecutar
          If accion1 Then
             accion = True
          Else
             accion = ValidarEjecucionProceso(id_tarea, fecha, MatSecProcesos, MatSecSubproc, opcion)
          End If
          If accion Then
              Call BloquearProcesoF(fecha, id_tarea, accion2, opcion, sibloqueo)
             If sibloqueo Then
                parproc = ObParamProc2(matpr, i)
                Call ListaProcesos(fecha, id_proc, parproc, opcion, txtmsg, exito)
                Call ValidarActRealizadaF(fecha, id_tarea, opcion, txtmsg, exito)
                DoEvents
                If exito Then
                   MensajeProc = "Se termino el proceso " & matpr(i, 3) & " del " & fecha
                   Call GuardaDatosBitacora(2, "Proceso", id_tarea, matpr(i, 3), NomUsuario, fecha, MensajeProc, 2)
                Else
                   MensajeProc = "No se realizo el proceso " & matpr(i, 3) & " del " & fecha
                End If
      
             Else
                MensajeProc = "No se puede bloquear el proceso"
             End If
          Else
             MensajeProc = "No se puede ejecutar el proceso"
          End If
          DoEvents
      Next i
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ListaProcesos(ByVal fecha As Date, ByVal id_proc As String, ByRef matpar() As Variant, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
'objetivo de la funcion:
'centralizar la llamada a todos los procesos requeridos para el calculo
'de las metricas de riesgo
'desde la carga de las bases de datos
'hasta ejecucion de procesos principales de calculo

'precondiciones de los datos de entrada
'fecha  - debe de ser una fecha laborable en mexico y estar en la tabla VAR_FECHAS_VAR
'id_proc - debe de estar en la lista del select case

'entradas
'fecha  -  fecha del proceso, que es la fecha a la que corresponden los datos procesado
'id_proc - es el identificador para la subrutina que y que se encuentra mapeada en el
'matpar() es una array de tipo variant que porta todos los parametros a usar en la subrutina llamada por id_proc

'periodicidad - se ejecuta todos los dias habiles

'resultados
'txtmsg  - que indica los eventos ocurridos al ejecutar la subrutina elegida
'exito   - indica si la ejecucion de la subrutina elegida fue exitosa
'noreg   - indica el numero de registros para suburutinas que implican el procesamiento de registros
'salida del proceso: en funcion de la llamada a una subrutina determinada por id_proc



Dim noreg As Long
Dim noreg1 As Long
Dim noreg2 As Long
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim txtmsg1 As String
Dim txtmsg2 As String
Select Case id_proc
Case "scer"   'calculo de efectividad retrospectiva
   Call ProcEficRetro(fecha, txtmsg, exito)
Case "dcpp"      'descargar curvas del servidor
   Call ObtenerArchSFTP(fecha, "CURVAS", "", "XLS", matpar(1), matpar(2), matpar(3), "batchcurvas.bat", txtmsg, exito)
Case "ccci"          'CREAR CURVAS CSV IKOS
   Call CrearCurvasCSVIKOS(fecha, matpar(1), matpar(2), txtmsg, exito)
Case "ccps"
'archivo curvas completo en forma compacta
   If Not EsVariableVacia(matpar(2)) Then
      Call GuardaCurvaInd(fecha, matpar(1), CStr(matpar(2)), noreg, txtmsg, exito)
   Else
      Call GuardaCurvasTotal2(fecha, matpar(1), noreg, txtmsg, exito)
   End If
Case "cncs"
   If Not EsVariableVacia(matpar(1)) Then
      Call ImpNdCurvas(fecha, matpar(1), noreg, txtmsg, exito)
   Else
      Call ImpNdCurvas(fecha, "", noreg, txtmsg, exito)
   End If
   If exito Then SiCargaTasas = True
Case "acpp"
   Call AnalisisCurvas(fecha, matpar(1), matpar(2), txtmsg, exito)
Case "dvp1"
'bajar vector 1 del servidor
   Call TransferirVect2(fecha, "VMD", "CSV", CStr(matpar(1)), CStr(matpar(2)), CStr(matpar(3)), True, True, "batchvp2.bat", txtmsg, exito)
Case "dvp2"
'bajar vector 2 del servidor
   Call TransferirVect1(fecha, "PIP", "M", "XLS", CStr(matpar(1)), CStr(matpar(2)), CStr(matpar(3)), True, True, "batchvp1.bat", txtmsg, exito)
Case "dvp3"
'bajar vector analitico
   Call ObtVectAnalitico(fecha, "VectorAnalitico", "MD", "CSV", CStr(matpar(1)), CStr(matpar(2)), CStr(matpar(3)), True, True, "batchvpa.bat", txtmsg, exito)
Case "cvps"
'el vector de precios en forma estandar
   Call ImportarVPrecios(fecha, CStr(matpar(1)), noreg, 0, txtmsg, exito)
Case "efrvp"
'indices en el vector pip
   If Not EsVariableVacia(matpar(1)) Then
      Call ImportarFactVP(fecha, CStr(matpar(1)), noreg, txtmsg, exito)
   Else
      Call ImportarFactVP(fecha, "", noreg, txtmsg, exito)
   End If
Case "efriv"
   Call ObtTRefPIPO(fecha, fecha, noreg1, txtmsg, exito1)
   Call ObtenerTIIE(fecha, txtmsg, exito2)
   Call ObtenerLibort2(fecha, txtmsg, exito)
   Call ObtTCambioMPIP(fecha, fecha, noreg2, txtmsg, exito3)
   noreg = noreg1 + noreg2 + 1
   If exito1 And exito2 And exito3 Then
   exito = True
   Else
   exito = False
   End If
Case "eyisv"  'curvas yield bonos sobretasa
   Call ObtenerYieldsIS(fecha, txtmsg, exito)
Case "afrg"
'se hace la validacion de los factores de riesgo
   Call ValidarFactRiesgo(fecha, Val(matpar(1)), txtmsg, exito)
Case "gsee"
   Call CrearSubProcValExtremos(fecha, CDate(matpar(2)), 63, id_tabla, txtmsg, exito)
Case "ipmd"
'se lee la posicion de mercado de dinero
   Call ImportarPosMDinero(fecha, noreg, txtmsg, exito)
Case "dfpp"
'bajar flujos de servidor
   Call ObtenerArchSFTP(fecha, "BanobrasFlujos", "", "XLS", matpar(3), matpar(1), matpar(2), "batchflujos1.bat", txtmsg, exito1)
   Call ObtenerArchSFTP(fecha, "BanobrasFlujos", "_2", "XLS", matpar(3), matpar(1), matpar(2), "batchflujos2.bat", txtmsg, exito2)
   If exito1 And exito2 Then
     exito = True
   Else
     exito = False
   End If
Case "gfed"
   Call GenerarFlujosMD(fecha, noreg, txtmsg, exito)
Case "ipmc"
'posicion de divisas
   Call ImportarPosCamArch(fecha, matpar(1), noreg, txtmsg, exito)
Case "ipsw"
'LEER POS SWAPS
   Call ImpPosSwapsRed(fecha, noreg, txtmsg, exito)
   Call LlenarRelSwapEm
Case "ipfwd"
'posicion de fwds tc
   Call ImpPosFwdRed(fecha, noreg, txtmsg, exito)
Case "attsw"
 'tasas cupon de swaps y primarias
   Call ActTCFlujosSwaps(fecha, 1, matpar(1), txtmsg1, exito1)
   Call ActTCFlujosDeuda(fecha, matpar(1), txtmsg2, exito2)
   If exito1 And exito2 Then
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      txtmsg = txtmsg1 & "," & txtmsg2
      exito = False
   End If
Case "ivdi"
'importar valuacion de derivados de ikos
   Call ImpValDerIKOS(fecha, noreg, txtmsg, exito)
Case "gspp"
'generar portafolios de posicion
   Call ProcGenPortafolios(fecha, 64, 65, 66, id_tabla, txtmsg, exito)
Case "gsvp"
'generar subprocesos de valuacion
   Call GenSubProcValOper(fecha, matpar(1), "Normal", 1, 67, id_tabla, txtmsg, exito1)
   Call GenSubProcValOper(fecha, matpar(1), "Normal", 2, 67, id_tabla, txtmsg, exito2)
   exito = exito1 And exito2
Case "cvcontrap"
'valuacion por contrapartes
   Call GenSubprocValContrap(fecha, matpar(1), "Normal", 68, id_tabla, txtmsg, exito)
Case "gsvport"
'subprocesos valuacion de subportafolios
   Call GenSubProcValPosSubPort(fecha, matpar(1), "Normal", matpar(2), 68, id_tabla, txtmsg, exito)
Case "gsepgo"
  'generar subproc pyg por operacion
   Call GenSubCalcPyGOper(fecha, matpar(1), "Real", "Normal", Val(matpar(2)), Val(matpar(3)), 69, id_tabla, txtmsg, exito)
Case "gpgp"
'p y g por subportafolio
   Call GenSubprocCalcPyGSubport(fecha, matpar(1), "Normal", matpar(2), Val(matpar(3)), Val(matpar(4)), 70, id_tabla, txtmsg, exito)
   Call GenProcPyGEmPos(fecha, matpar(1), "Normal", Val(matpar(3)), Val(matpar(4)), 70, id_tabla)
   Call GenProcPyGPIDV(fecha, matpar(1), "Normal", Val(matpar(3)), Val(matpar(4)), 70, id_tabla)
   Call GenProcPyGEstructural(fecha, matpar(1), "Normal", Val(matpar(3)), Val(matpar(4)), 70, id_tabla)
Case "gseeo"
'generar subprocesos de escenarios de estres por operacion
   Call GenSubProcEscenEstres(fecha, matpar(1), "Normal", 71, id_tabla, txtmsg, exito)
Case "ceeport"
   Call GenSubProcConsolEscEstres(fecha, matpar(1), matpar(2), 72, id_tabla, txtmsg, exito)
Case "gscsensibo"
'generar subproc calc sensib x operacion
   Call GenSubCalcSensibOper(fecha, matpar(1), "Normal", 73, id_tabla, txtmsg, exito)
Case "csport"
   Call GenSubprocCalcSensibSubPort(fecha, matpar(1), "Normal", matpar(2), 74, id_tabla, txtmsg, exito)
Case "cvmark"
'var markowitz consolidado
   Call GenVaRMark(fecha, matpar(1), "Normal", matpar(2), Val(matpar(3)), Val(matpar(4)), Val(matpar(5)), txtmsg, exito)
Case "cmatchol"
   Call CalculoMatCholeski(fecha, matpar(1), matpar(2), Val(matpar(3)), Val(matpar(4)), txtmsg, exito)
Case "gsetaylor"
   Call GenSubProcEscEstresTaylor(fecha, CDate(matpar(1)), matpar(2), matpar(3), 75, id_tabla, txtmsg, exito)
Case "cvmont"
'pyg montecarlo por operacion
   Call GenSubprocPyGMontOper(fecha, matpar(1), "Normal", Val(matpar(2)), Val(matpar(3)), Val(matpar(4)), 76, id_tabla, txtmsg, exito)
Case "gpgmontport"
'generar p y g por portafolio var montecarlo
   Call GenSubprocCalcPyGMontSubport(fecha, matpar(1), "Normal", matpar(2), Val(matpar(3)), Val(matpar(4)), Val(matpar(5)), 77, id_tabla, txtmsg, exito)
Case "cpygback"
'p y g para backtesting
   Call SecuenciaBack(fecha, matpar(1), matpar(2), txtmsg, exito)
Case "crvp"
'calcular cvar por subportafolios
   Call GeneraResCVaRPos(fecha, fecha, fecha, matpar(1), "Normal", matpar(2), Val(matpar(3)), Val(matpar(4)), Val(matpar(5)), txtmsg, exito)
   Call GenProcCVaREmPos(fecha, matpar(1), "Normal", Val(matpar(3)), Val(matpar(4)), Val(matpar(5)), 23)
   Call GenProcCVaRPIDV(fecha, matpar(1), "Normal", Val(matpar(3)), Val(matpar(4)), Val(matpar(5)), 23)
   Call GenProcCVaREstructural(fecha, matpar(1), "Normal", Val(matpar(3)), Val(matpar(4)), Val(matpar(5)), 23)
Case "crvprev"
'cvar preventivo
   Call GeneraResCVaRPrevPos(fecha, matpar(1), matpar(2), Val(matpar(3)), Val(matpar(4)), 20, 0.99, txtmsg, exito)
Case "crvexp"
'cvar exponencial
   Call GeneraResCVaRExpPos(fecha, matpar(1), "Normal", matpar(2), Val(matpar(3)), Val(matpar(4)), Val(matpar(5)), Val(matpar(6)), txtmsg, exito)
Case "cvmontport"
'generar var montecarlo por subportafolios
   Call GeneraResVaRMontPort(fecha, matpar(1), "Normal", matpar(2), Val(matpar(3)), Val(matpar(4)), Val(matpar(5)), Val(matpar(6)), txtmsg, exito)
Case "ervik"
'exportar el cvar al sistema ikos derivados
   Call GuardaResVaRIKOS2(fecha, conAdoBD, txtmsg, exito)
Case "gscva"
   Call GeneraLSubprocCVA(fecha, 110, matpar(1), 500, 1, id_tabla, txtmsg, exito)
Case Else
MsgBox "No se ha definido una subrutina para este proceso"
End Select
End Sub

Sub LlenarRelSwapEm()
Dim i As Integer
Dim txtcadena As String
For i = 1 To UBound(MatRelSwapEm, 1)
    txtcadena = "UPDATE " & TablaPosSwaps & " SET SI_PIDV ='S', C_EM_PIDV ='" & MatRelSwapEm(i, 2) & "' WHERE COPERACION = '" & MatRelSwapEm(i, 1) & "' AND TIPOPOS =1"
    ConAdo.Execute txtcadena
Next i
End Sub


Sub AnalisisCurvas(ByVal fecha As Date, ByVal noesc As Long, ByVal nodesv As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim indice As Long
Dim indice0 As Long
Dim indice1 As Long
Dim matx() As Double
Dim fecha0 As Date
Dim fecha1 As Date
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim contar As Long
Dim contar1 As Long
Dim i As Long
Dim j As Long
Dim htiempo As Long
Dim matrends() As Double
Dim resp As Integer
Dim noreg As Integer
Dim txtmsg1 As String
Dim matresul() As New resCalcVol
Dim txtnomarch As String
Dim exitoarch As Boolean
noreg = 14
ReDim matc(1 To noreg) As String


matc(1) = "BAN B1"
matc(2) = "BONDES D"
matc(3) = "BPAG28"
matc(4) = "BPAG91"
matc(5) = "CCMID"
matc(6) = "CCS YEN-PESOS"
matc(7) = "CETES IMP"
matc(8) = "DESC IRS"
matc(9) = "IMP FWD EURO"
matc(10) = "IMP PESOS PIP"
matc(11) = "LIBOR"
matc(12) = "CCS UDI-TIIE"
matc(13) = "REAL IMP"
matc(14) = "REP B1"
exito2 = True
htiempo = 1
indice = BuscarValorArray(fecha, MatFechasVaR, 1)
If indice <> 0 Then
   fecha0 = MatFechasVaR(indice - noesc - htiempo, 1)
   fecha1 = MatFechasVaR(indice - htiempo, 1)   'fecha del dia anterior
   Call CrearMatFRiesgo2(fecha0, fecha, MatFactRiesgo, txtmsg1, exito1)
   indice0 = BuscarValorArray(fecha1, MatFactRiesgo, 1)
   indice1 = BuscarValorArray(fecha, MatFactRiesgo, 1)
   If indice0 <> 0 And indice1 <> 0 Then
      contar = 0
      ReDim matsigma(1 To NoFactores) As Double
      ReDim matresul(1 To 1)
      For i = 1 To UBound(MatCaracFRiesgo, 1)
          For j = 1 To noreg
            If MatCaracFRiesgo(i).nomFactor = matc(j) Then
               contar = contar + 1
               ReDim Preserve matresul(1 To contar)
               matresul(contar).nomFactor = MatCaracFRiesgo(i).descFactor
               matresul(contar).valfactt_1 = MatFactRiesgo(indice0, i + 1)
               matresul(contar).valfactt = MatFactRiesgo(indice1, i + 1)
               matx = ConvArVtDbl(ExtraeSubMatrizV(MatFactRiesgo, i + 1, i + 1, 1, UBound(MatFactRiesgo, 1) - 1))
               matresul(contar).desvest = CalcVol5(matresul(contar).valfactt_1, matx, htiempo, MatCaracFRiesgo(i).tfactor)
               matresul(contar).liminf = matresul(contar).valfactt_1 - matresul(contar).desvest * nodesv
               matresul(contar).limsup = matresul(contar).valfactt_1 + matresul(contar).desvest * nodesv
               If matresul(contar).valfactt < matresul(contar).liminf Or matresul(contar).valfactt > matresul(contar).limsup Then
                  exito2 = False
               End If
            End If
        Next j
      Next i
      If Not exito2 Then
         contar1 = 0
         txtnomarch = DirResVaR & "\Analisis curvas " & Format(fecha, "yyyy-mm-dd") & ".txt"
         frmEjecucionProc.CommonDialog1.FileName = txtnomarch
         frmEjecucionProc.CommonDialog1.ShowSave
         txtnomarch = frmEjecucionProc.CommonDialog1.FileName
         Call VerificarSalidaArchivo(txtnomarch, 1, exitoarch)
         If exitoarch Then
         Print #1, "Descripcion del factor" & Chr(9) & "valor t-1" & Chr(9) & "valor t" & Chr(9) & "Desviacion estandar" & Chr(9) & "Cota inferior" & Chr(9) & "Cota superior"
         For i = 1 To contar
             If matresul(i).valfactt < matresul(i).liminf Or matresul(i).valfactt > matresul(i).limsup Then
                contar1 = contar1 + 1
                Print #1, matresul(i).nomFactor & Chr(9) & matresul(i).valfactt_1 & Chr(9) & matresul(i).valfactt & Chr(9) & matresul(i).desvest & Chr(9) & matresul(i).liminf & Chr(9) & matresul(i).limsup
             End If
         Next i
         Close #1
         End If
      End If
      If exito2 Then
         txtmsg = "El proceso finalizo correctamente"
         exito = True
      Else
         MsgBox "En " & contar1 & " factores, la variacion excede las " & nodesv & " desviaciones estandar. Se recomienda revisar los datos"
         resp = MsgBox("Desea continuar", vbYesNo)
         If resp = 6 Then
            txtmsg = "El proceso finalizo correctamente"
            exito = True
         Else
            txtmsg = "En revision de los datos del proveedor"
            exito = False
         End If
      End If
   Else
      txtmsg = "No hay fecha"
      exito = False
   End If
Else
   txtmsg = "No hay fecha"
   exito = False
End If
End Sub

Sub ValidarFactRiesgo(ByVal fecha As Date, ByVal noesc As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim indice As Long
Dim indice0 As Long
Dim indice1 As Long
Dim matx() As Double
Dim fecha0 As Date
Dim fecha1 As Date
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim i As Long
Dim j As Long
Dim htiempo As Long
Dim matrends() As Double
Dim resp As Integer
Dim noreg As Integer
Dim txtmsg1 As String
Dim matresul() As New resCalcVol
Dim txtnomarch As String
Dim exitoarch As Boolean
Dim nodesv As Double
Dim txtcadena As String
Dim vobservado As Double

exito2 = True
htiempo = 1
nodesv = 3
indice = BuscarValorArray(fecha, MatFechasVaR, 1)
If indice <> 0 Then
   fecha0 = MatFechasVaR(indice - noesc - htiempo, 1)
   fecha1 = MatFechasVaR(indice - htiempo, 1)   'fecha del dia anterior
   Call CrearMatFRiesgo2(fecha0, fecha, MatFactRiesgo, txtmsg1, exito1)
   indice0 = BuscarValorArray(fecha1, MatFactRiesgo, 1)
   indice1 = BuscarValorArray(fecha, MatFactRiesgo, 1)
   If indice0 <> 0 And indice1 <> 0 Then
      ReDim matsigma(1 To NoFactores) As Double
      ReDim matresul(1 To NoFactores)
      For i = 1 To UBound(MatCaracFRiesgo, 1)
          matresul(i).nomFactor = MatCaracFRiesgo(i).descFactor
          matresul(i).valfactt_1 = MatFactRiesgo(indice0, i + 1)
          matresul(i).valfactt = MatFactRiesgo(indice1, i + 1)
          matx = ConvArVtDbl(ExtraeSubMatrizV(MatFactRiesgo, i + 1, i + 1, 1, UBound(MatFactRiesgo, 1) - 1))
          matresul(i).desvest = CalcVol5(matresul(i).valfactt_1, matx, htiempo, MatCaracFRiesgo(i).tfactor)
          matresul(i).liminf = matresul(i).valfactt_1 - matresul(i).desvest * nodesv
          matresul(i).limsup = matresul(i).valfactt_1 + matresul(i).desvest * nodesv
      Next i
      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtcadena = "delete from " & TablaAnalisisFRO & " WHERE FECHA = " & txtfecha
      ConAdo.Execute txtcadena
      For i = 1 To NoFactores
          txtcadena = "INSERT INTO " & TablaAnalisisFRO & " VALUES("
          txtcadena = txtcadena & txtfecha & ","                                  'fecha
          txtcadena = txtcadena & "'" & MatCaracFRiesgo(i).nomFactor & "',"       'concepto
          txtcadena = txtcadena & MatCaracFRiesgo(i).plazo & ","                  'plazo
          txtcadena = txtcadena & matresul(i).valfactt_1 & ","                    'valor t - 1
          txtcadena = txtcadena & matresul(i).valfactt & ","                      'valor en t
          vobservado = Abs(matresul(i).valfactt - matresul(i).valfactt_1)         'incremento observado
          txtcadena = txtcadena & vobservado & ","
          txtcadena = txtcadena & matresul(i).desvest & ","                       'incremento esperado
          txtcadena = txtcadena & noesc & ")"                                     'no de escenarios
          ConAdo.Execute txtcadena
          txtmsg = "El proceso finalizo correctamente"
          exito = True
      Next i
   End If
End If
End Sub


Sub ImpValDerIKOS(ByVal fecha As Date, ByRef noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim paso1 As Boolean
Dim paso2 As Boolean
Dim noreg1 As Integer
Dim matcar1() As Variant
Dim matcar2() As Variant
Dim i As Integer
Dim txtfecha As String
Dim txtborra As String
Dim txtcadena As String
Dim noreg2 As Long

paso1 = False
paso2 = False
matcar1 = LeerValSwapsIKOS(fecha, TablaInterfCarac, conAdoBD)    'las caracteristicas de los swaps
matcar2 = ImpValFwdIkosO(fecha, TablaInterfFwd, noreg2, conAdoBD)
If UBound(matcar1, 1) <> 0 And UBound(matcar1, 1) <> 0 Then
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtborra = "DELETE FROM " & TablaVDerIKOS & " WHERE FECHA = " & txtfecha
ConAdo.Execute txtborra
noreg1 = UBound(matcar1, 1)
If noreg1 > 0 Then
   paso1 = True
   For i = 1 To noreg1
       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaVDerIKOS & " VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & matcar1(i, 2) & "',"  'CLAVE DE operacion
       txtcadena = txtcadena & matcar1(i, 3) & ","         'val pata activa
       txtcadena = txtcadena & matcar1(i, 4) & ","         'val pata pasiva
       txtcadena = txtcadena & matcar1(i, 5) & ")"         'mtm
       ConAdo.Execute txtcadena
       AvanceProc = i / UBound(matcar1, 1)
       MensajeProc = "Leyendo la valuacion de derivados del " & fecha & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
Else
paso1 = False
End If

If UBound(matcar2, 1) > 0 Then
   paso2 = True
   For i = 1 To UBound(matcar2, 1)
       txtfecha = "to_date('" & Format(matcar2(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaVDerIKOS & " VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & matcar2(i, 2) & "',"
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & matcar2(i, 3) & ")"
       ConAdo.Execute txtcadena
       AvanceProc = i / UBound(matcar2, 1)
       MensajeProc = "Leyendo la valuacion de derivados del " & fecha & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
 Else
 paso2 = False
 End If
 noreg = noreg1 + noreg2
 If fecha = fechavalIKOS Then fechavalIKOS = 0
End If
If paso1 And paso2 Then
 exito = True
 txtmsg = "El proceso finalizo correctamente"
Else
 exito = False
 txtmsg = "No existen datos para esta fecha"
End If
End Sub

Function DeterminaTRef1(ByVal valor1 As String, ByVal valor2 As String, ByRef matpos1() As propPosSwaps)
Dim i As Long


For i = 1 To UBound(matpos1, 1)
    If valor1 = matpos1(i).c_operacion And valor2 = "B" Then
       DeterminaTRef1 = matpos1(i).TCActivaSwap
       Exit Function
    ElseIf valor1 = matpos1(i).c_operacion And valor2 = "C" Then
       DeterminaTRef1 = matpos1(i).TCPasivaSwap
       Exit Function
    End If
Next i
MensajeProc = "No se encontro la tasa de referencia de la operacion " & valor1
DeterminaTRef1 = ""
End Function

Function DeterminaTRef2(ByVal valor1 As String, ByRef matpos1() As propPosDeuda)
Dim i As Long

For i = 1 To UBound(matpos1, 1)
    If valor1 = matpos1(i).c_operacion Then
       DeterminaTRef2 = matpos1(i).TRefDeuda
       Exit Function
    End If
Next i
End Function

Sub ActTCFlujosSwaps(ByVal fecha As Date, ByVal tipopos As Integer, ByVal opcion As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim fecha1 As Date
Dim fecha2 As Date
Dim matpos1() As propPosSwaps
Dim coperacion As String
Dim tpata As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim txtfiltro3 As String
Dim matflujos() As New estFlujosDeuda
Dim mata() As Variant
Dim matb() As Variant
Dim matc() As Variant
Dim matd() As Variant
Dim mate() As Variant
Dim ncol1 As Integer
Dim i As Integer
Dim fechab As Date
Dim tcupon As Double
Dim tasaref As String
Dim accion As Integer
Dim indice As Integer
Dim fechax As Date
Dim txtcadena As String
Dim noreg As Long

txtmsg = "El proceso finalizo correctamente"
    fecha1 = fecha - 150
    fecha2 = fecha
    txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro1 = "SELECT * FROM " & TablaPosSwaps
    txtfiltro1 = txtfiltro1 & " where TIPOPOS = " & tipopos
    txtfiltro1 = txtfiltro1 & " AND (FECHAREG,COPERACION) IN (SELECT MAX(FECHAREG) AS FECHAREG,COPERACION FROM " & TablaPosSwaps
    txtfiltro1 = txtfiltro1 & " WHERE FECHAREG <= " & txtfecha2 & " AND TIPOPOS = 1 GROUP BY COPERACION) AND FINICIO <= " & txtfecha2
    txtfiltro1 = txtfiltro1 & " AND FVENCIMIENTO > " & txtfecha2
    txtfiltro1 = txtfiltro1 & " ORDER BY COPERACION"
    If opcion = "S" Then
       txtfiltro3 = "SELECT * FROM " & TablaFlujosSwapsO
       txtfiltro3 = txtfiltro3 & " where FINICIO >= " & txtfecha1 & " AND FINICIO <= " & txtfecha2 & " AND TIPOPOS = " & tipopos
    Else
       txtfiltro3 = "SELECT * FROM " & TablaFlujosSwapsO & " where FINICIO >= " & txtfecha1 & " AND FINICIO <= " & txtfecha2 & " AND TIPOPOS = " & tipopos & " and TASA = 0"
    End If
    matpos1 = LeerTablaSwaps(txtfiltro1)
    matflujos = LeerFlujosSwaps(txtfiltro3, True)

'se obtiene la tasa de tiie y se toma como tasa cupon para flujos de tiie
'txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
exito = True

If Not EsArrayVacio(matflujos) Then
   mata = Leer1FactorR(fecha1 - 30, fecha2, "TIIE 28", 0)
   matb = Leer1FactorR(fecha1 - 30, fecha2, "TIIE 91", 0)
   matc = Leer1FactorR(fecha1 - 30, fecha2, "LIBOR 1M PIP", 0)
   matd = Leer1FactorR(fecha1 - 30, fecha2, "LIBOR 3M PIP", 0)
   mate = Leer1FactorR(fecha1 - 30, fecha2, "LIBOR 6M PIP", 0)
   ncol1 = 4
'se actualizan las tasas tiie para los swaps de cupon de 28 dias
   noreg = UBound(matflujos, 1)
   For i = 1 To noreg
       fechab = matflujos(i).finicio                    'fecha de inicio del flujo
       coperacion = matflujos(i).coperacion             'clave de la operacion
       tpata = matflujos(i).tpata                       'pertenece a la pata activa o pasiva
       tcupon = 0
       tasaref = DeterminaTRef1(coperacion, tpata, matpos1)
'tasa cupon fija
       If coperacion = "369" Or coperacion = "505" Or coperacion = "405" Or coperacion = "386" Or coperacion = "465" Or coperacion = "370" Or coperacion = "506" Or coperacion = "404" Or coperacion = "387" Or coperacion = "466" Then
          accion = 2
       Else
          accion = 1
       End If
       If Val(ReemplazaVacioValorO(tasaref, 0, 3)) = 0 Then
          Select Case tasaref
              Case "TIIE28[0]',", "TIIE28[0]", "TIIE 28[0]", "TIIE28[0", "TIIE 28"
                  indice = BuscarValorArray(fechab, mata, 1)
                  If indice <> 0 Then
                     tcupon = mata(indice, 2)   'valor de la tasa cupon vigente
                  Else
                     MensajeProc = "Falta la TIIE 28 cupon del " & fechab & " swap " & coperacion
                     txtmsg = MensajeProc
                     exito = False
                  End If
              Case "TIIE91[0]", "TIIE91[0"
                  indice = BuscarValorArray(fechab, matb, 1)
                  If indice <> 0 Then
                     tcupon = matb(indice, 2)
                  Else
                     MensajeProc = "Falta la TIIE 91 para el cupon del " & fechab & " swap " & coperacion
                     txtmsg = MensajeProc
                     exito = False
                  End If
              Case "LIBOR 28"
                  indice = BuscarValorArray(fechab, matc, 1)
                  If indice <> 0 Then
                     tcupon = matc(indice, 2)
                  Else
                     MensajeProc = "Falta la LIBOR 28 para el " & fechab
                     txtmsg = MensajeProc
                     exito = False
                  End If
              Case "LIBOR1M[2]"
                  fechax = DescDiasHabUSUK(fechab, 2)
                  indice = BuscarValorArray(fechax, matc, 1)
                  If indice <> 0 Then
                     tcupon = matc(indice, 2)
                  Else
                     MensajeProc = "Falta la LIBOR 28(-2) para el cupon del " & fechab
                     txtmsg = MensajeProc
                     exito = False
                  End If
              Case "LIBOR3M[0]", "LIBOR 3M", "LIBOR3M[0"
                  indice = BuscarValorArray(fechab, matd, 1)
                  If indice <> 0 Then
                     tcupon = matd(indice, 2)
                  Else
                     MensajeProc = "Falta la LIBOR 91 para el cupon del " & matflujos(i, 1) & " " & fechab
                     txtmsg = MensajeProc
                     exito = False
                  End If
             Case "LIBOR3M[2]", "LIBOR3M[2"
                  fechax = DescDiasHabUSUK(fechab, 2)
                  indice = BuscarValorArray(fechax, matd, 1)
                  If indice <> 0 Then
                     tcupon = matd(indice, 2)
                  Else
                     MensajeProc = "Falta la LIBOR 91(-2) para el cupon del " & fechab
                     txtmsg = MensajeProc
                     exito = False
                  End If
             Case "LIBOR 6M", "LIBOR6M[0", "LIBOR6M[0]"
                  indice = BuscarValorArray(fechab, mate, 1)
                  If indice <> 0 Then
                     tcupon = mate(indice, 2)
                  Else
                     MensajeProc = "Falta la LIBOR 180  cupon  " & fechab & " swap " & coperacion
                     txtmsg = MensajeProc
                     exito = False
                  End If
             Case "LIBOR6M[2]", "LIBOR6M[2"
                  fechax = DescDiasHabUSUK(fechab, 2)
                  indice = BuscarValorArray(fechax, mate, 1)
                  If indice <> 0 Then
                     tcupon = mate(indice, 2)
                  Else
                     MensajeProc = "Falta la LIBOR 6M(-2) para el cupon del " & fechab & " (" & fechax & ")"
                     txtmsg = MensajeProc
                     exito = False
                  End If
             Case Else
                  MensajeProc = "Se deconoce la tasa de referencia " & tasaref & " para la operacion " & coperacion & " por lo que no se puede buscar la tasa cupon vigente"
                  txtmsg = MensajeProc
             End Select
End If
'aqui viene lo bueno, en funcion de la busqueda se procede a realizar un update en la tabla de flujos
If tcupon <> 0 Then
   txtfecha = "to_date('" & Format(fechab, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   If opcion = "S" Then
      txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & tcupon & " WHERE COPERACION = '" & coperacion & "' AND TPATA = '" & tpata & "' AND FINICIO = " & txtfecha
   Else
      txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & tcupon & " WHERE COPERACION = '" & coperacion & "' AND TPATA = '" & tpata & "' AND FINICIO = " & txtfecha & " AND TASA = 0"
   End If
   ConAdo.Execute txtcadena
End If
AvanceProc = i / noreg
MensajeProc = "Actualizando la tasa cupon de los flujos de swaps " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i

End If

End Sub

Sub ActTCFlujosSwaps3(ByVal fecha As Date, ByRef mata() As Variant, ByVal txtpos As String, ByVal txttref As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim txtcadena As String
Dim fechax As Date
Dim indice As Long
Dim valor As Double
valor = 0
indice = 0
If txttref = "TIIE28[0]" Or txttref = "TIIE91[0]" Then
   indice = BuscarValorArray(fecha, mata, 1)
   If indice <> 0 Then valor = mata(indice, 2)
ElseIf txttref = "LIBOR3M[2]" Or txttref = "LIBOR1M[2]" Or txttref = "LIBOR6M[2]" Then
   fechax = DescDiasHabUSUK(fecha, 2)
   indice = BuscarValorArray(fechax, mata, 1)
   If indice <> 0 Then valor = mata(indice, 2)
End If
If valor <> 0 Then
   txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
   txtcadena = "UPDATE " & TablaFlujosSwapsO & " SET TASA = " & valor & " "
   txtcadena = txtcadena & "WHERE TASA = 0 AND (coperacion,tpata,finicio) in "
   txtcadena = txtcadena & "(SELECT A.COPERACION, a.tpata,a.FINICIO "
   txtcadena = txtcadena & "from " & TablaFlujosSwapsO & " A JOIN " & TablaPosSwaps & " B "
   txtcadena = txtcadena & "on a.TIPOPOS = b.TIPOPOS and "
   txtcadena = txtcadena & "A.CPOSICION = b.CPOSICION AND "
   txtcadena = txtcadena & "a.coperacion = b.coperacion "
   txtcadena = txtcadena & "where a.tpata = '" & txtpos & "' and "
   If txtpos = "B" Then
      txtcadena = txtcadena & "b.TC_ACTIVA = '" & txttref & "' and "
   Else
      txtcadena = txtcadena & "b.TC_PASIVA = '" & txttref & "' and "
   End If
   txtcadena = txtcadena & "a.FINICIO = " & txtfecha & ")"
   ConAdo.Execute txtcadena
   exito = True
Else
  txtmsg = "No hay datos para esta fecha"
  exito = False
End If
End Sub

Sub ActTCFlujosDeuda(ByVal fecha As Date, ByVal opcion As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim matpos1() As propPosDeuda
Dim coperacion As String
Dim fecha0 As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim txtfiltro3 As String
Dim matflujos() As estFlujosDeuda
Dim mata() As Variant
Dim matb() As Variant
Dim matc() As Variant
Dim matd() As Variant
Dim mate() As Variant
Dim ncol1 As Integer
Dim i As Integer
Dim fechab As Date
Dim tcupon As Double
Dim tasaref As String
Dim accion As Integer
Dim indice As Integer
Dim fechax As Date
Dim txtcadena As String
Dim noreg As Long

txtmsg = "El proceso finalizo correctamente"

    fecha1 = fecha - 150
    fecha2 = fecha
    txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfiltro1 = "SELECT * FROM  " & TablaPosDeuda & " WHERE TIPOPOS = 1"
    txtfiltro1 = txtfiltro1 & " AND (FECHAREG,COPERACION) IN (SELECT MAX(FECHAREG) AS FECHAREG,COPERACION"
    txtfiltro1 = txtfiltro1 & " FROM " & TablaPosDeuda & " WHERE FECHAREG <= " & txtfecha2 & " AND TIPOPOS = 1 group by coperacion)"
    txtfiltro1 = txtfiltro1 & " AND finicio <= " & txtfecha2 & " and fvencimiento > " & txtfecha2
    txtfiltro1 = txtfiltro1 & " ORDER BY COPERACION"
    If opcion = "S" Then
       txtfiltro3 = "SELECT * FROM  " & TablaFlujosDeudaO & " where FINICIO >= " & txtfecha1 & " AND FINICIO <= " & txtfecha2 & " AND TASA <> 0 AND TIPOPOS = 1 ORDER BY COPERACION,FINICIO"
    Else
       txtfiltro3 = "SELECT * FROM  " & TablaFlujosDeudaO & " where FINICIO >= " & txtfecha1 & " AND FINICIO <= " & txtfecha2 & " AND TIPOPOS = 1 ORDER BY COPERACION,FINICIO"
    End If
    matpos1 = LeerTablaDeuda(txtfiltro1)
    matflujos = LeerFlujosDeuda(txtfiltro3, False)
'se obtiene la tasa de tiie y se toma como tasa cupon para flujos de tiie
'txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
exito = True
If Not EsArrayVacio(matflujos) Then
   fecha0 = fecha1 - 200
   mata = Leer1FactorR(fecha0, fecha2, "TIIE 28", 0)
   matb = Leer1FactorR(fecha0, fecha2, "TIIE 91", 0)
   matc = Leer1FactorR(fecha0, fecha2, "LIBOR 1M PIP", 0)
   matd = Leer1FactorR(fecha0, fecha2, "LIBOR 3M PIP", 0)
   mate = Leer1FactorR(fecha0, fecha2, "LIBOR 6M PIP", 0)
   ncol1 = 4
'se actualizan las tasas tiie para los swaps de cupon de 28 dias
   noreg = UBound(matflujos, 1)
   For i = 1 To noreg
       fechab = matflujos(i).finicio                 'fecha de inicio del flujo
       coperacion = matflujos(i).coperacion          'clave de la operacion
       tcupon = 0
       tasaref = DeterminaTRef2(coperacion, matpos1)
'tasa cupon fija
If coperacion = "369" Or coperacion = "505" Or coperacion = "405" Or coperacion = "386" Or coperacion = "465" Or coperacion = "370" Or coperacion = "506" Or coperacion = "404" Or coperacion = "387" Or coperacion = "466" Then
   accion = 2
Else
   accion = 1
End If
       If Val(ReemplazaCadenaTexto(tasaref, "%", "")) = 0 Then
          Select Case tasaref
              Case "TIIE28[0]',", "TIIE28[0]", "TIIE 28[0]", "TIIE28[0", "TIIE 28", "TIIE28"
                  indice = BuscarValorArray(fechab, mata, 1)
                  If indice <> 0 Then
                     tcupon = mata(indice, 2)   'valor de la tasa cupon vigente
                  Else
                     MensajeProc = "Falta la TIIE 28 para el cupon del " & fechab & " swap " & coperacion
                     txtmsg = MensajeProc
                     exito = False
                  End If
              Case "TIIE91[0]"
                  indice = BuscarValorArray(fechab, matb, 1)
                  If indice <> 0 Then
                     tcupon = matb(indice, 2)
                  Else
                     MensajeProc = "Falta la TIIE 91 para el cupon del " & fechab
                     txtmsg = MensajeProc
                     exito = False
                  End If
              Case "LIBOR 28"
                  indice = BuscarValorArray(fechab, matc, 1)
                  If indice <> 0 Then
                     tcupon = matc(indice, 2)
                  Else
                     MensajeProc = "Falta la LIBOR 28 para el " & fechab
                     txtmsg = MensajeProc
                     exito = False
                  End If
              Case "LIBOR3M[0]", "LIBOR 3M", "LIBOR3M[0"
                  indice = BuscarValorArray(fechab, matd, 1)
                  If indice <> 0 Then
                     tcupon = matd(indice, 2)
                  Else
                     MensajeProc = "Falta  LIBOR 91 para el cupon del " & matflujos(i, 1) & " " & fechab
                     txtmsg = MensajeProc
                     exito = False
                  End If
             Case "LIBOR3M[2]", "LIBOR3M[2"
                  If accion = 1 Then
                     fechax = DescDiasHabUSUK(FBD(fechab, "MX"), 2)
                  ElseIf accion = 2 Then
                     fechax = DescDiasHabUSUK(fechab, 2)
                  End If
                  indice = BuscarValorArray(fechax, matd, 1)
                  If indice <> 0 Then
                     tcupon = matd(indice, 2)
                  Else
                     MensajeProc = "Falta la LIBOR 91 para el cupon del " & fechab
                     txtmsg = MensajeProc
                     exito = False
                  End If
Case "LIBOR 6M", "LIBOR6M[0", "LIBOR6M"
                  indice = BuscarValorArray(fechab, mate, 1)
                  If indice <> 0 Then
                     tcupon = mate(indice, 2)
 Else
    MensajeProc = "Falta la LIBOR 180 para el cupon del " & fechab
    txtmsg = MensajeProc
    exito = False
 End If
Case "LIBOR6M[2]", "LIBOR6M[2"
 fechax = DescDiasHabUS(fechab, 2)
 indice = BuscarValorArray(fechax, mate, 1)
 If indice <> 0 Then
  tcupon = mate(indice, 2)
 Else
    MensajeProc = "Falta la LIBOR 6M(-2) para el cupon del " & fechab & " (" & fechax & ")"
    txtmsg = MensajeProc
    exito = False
 End If
Case Else
    MensajeProc = "Se deconoce la tasa de referencia " & tasaref & " para el flujo " & matflujos(i).finicio & " por lo que no se puede buscar la tasa cupon vigente"
    txtmsg = MensajeProc
End Select
End If
'aqui viene lo bueno, en funcion de la busqueda se procede a realizar un update en la tabla de flujos
If tcupon <> 0 Then
   txtfecha = "to_date('" & Format(fechab, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtcadena = "UPDATE " & TablaFlujosDeudaO & " SET TASA = " & tcupon & " WHERE COPERACION = '" & coperacion & "' AND FINICIO = " & txtfecha & " AND TASA = 0"
   ConAdo.Execute txtcadena
End If
AvanceProc = i / noreg
MensajeProc = "Actualizando la tasa cupon de los flujos de op de deuda " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i
End If

End Sub

Sub ImpPosSwapsID(ByVal fecha As Date, ByVal siarch As Boolean, ByVal txtnompos As String, ByRef nrn As Integer, ByRef nrc As Integer)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
   Call ImpPosSwapsIDRed(fecha, txtnompos, nrn, nrc)         '"S" con contraparte, "N" sin contraparte
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function TradCPosMD(coper)
 'para la posicion del IKOS se deben de cambiar las claves de la posicion
       If coper = 1 Then
          TradCPosMD = 1                       'compra en directo
       ElseIf coper = 2 Then
          TradCPosMD = 3                       'venta en reporto
       ElseIf coper = 3 Then
          TradCPosMD = 2                       'compra en reporto
       ElseIf coper = 4 Then
          TradCPosMD = 2                       'compra en reporto fecha valor
       ElseIf coper = 5 Then
          TradCPosMD = 3                       'venta en reporto fecha valor
       ElseIf coper = 6 Then
          TradCPosMD = 1                       'compra en directo fecha valor
       ElseIf coper = 7 Then
          TradCPosMD = 4                       'venta en directo fecha valor
       ElseIf coper = 10 Then
          TradCPosMD = 1                       'traspaso de la teso a la mesa
       ElseIf coper = "D" Then
          TradCPosMD = 1                      'directo
       ElseIf coper = "R" Then
          TradCPosMD = 2                      'reporto
       End If

End Function

Function VerificarPosMesaPen(ByRef mata() As Variant, ByVal id_pos As Integer) As Variant()
Dim noreg As Long
Dim nocampos As Long
Dim contar As Long
Dim i As Long

'rutina para completar algunos datos de la posicion de mesa de dinero
'esta rutina es especifica para la mesa de dinero
'rutina transitoria
'esta rutina se adecuo para trabajar con los archivos del IKOS
'no se debe usar en otro proceso
noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
If noreg <> 0 Then
   ReDim matb(1 To noreg, 1 To 20) As Variant
   For i = 1 To noreg
       matb(i, 1) = CDate(mata(i, 1))                   'fecha de posicion
       matb(i, 2) = mata(i, 2)                          'intencion
       matb(i, 3) = id_pos                              'clave de posicion
       matb(i, 4) = mata(i, 4)                          'clave de operacion
       If mata(i, 5) = "D" Then
          matb(i, 5) = 1                                'tipo de operacion
       Else
          matb(i, 5) = 2
       End If
       matb(i, 6) = mata(i, 6)                          'tipo valor
       matb(i, 7) = Trim(mata(i, 7))                    'emision
       matb(i, 8) = mata(i, 8)                          'serie
       matb(i, 9) = GeneraClaveEmision(mata(i, 6), mata(i, 7), mata(i, 8))
       matb(i, 10) = mata(i, 10)                         'no titulos
       matb(i, 11) = mata(i, 11)                        'fecha de compra
       matb(i, 12) = mata(i, 12)                        'FECHA VENCIMIENTO OPERACION
       matb(i, 13) = ConvValor(mata(i, 13))             'p asignado/compra
       matb(i, 14) = ConvValor(mata(i, 14))             'tasa reporto
       matb(i, 15) = mata(i, 15)                        'subportafolio 1
       matb(i, 16) = mata(i, 16)                        'subportafolio 2
       AvanceProc = i / noreg
       MensajeProc = "Cambiando el orden de los campos"
   Next i
Else
   ReDim matb(0 To 0, 0 To 0) As Variant
   MsgBox "No hay datos en la posicion"
End If
VerificarPosMesaPen = matb
End Function


Sub CompPosMesaVPrecios(ByVal fecha As Date, ByVal nompos As String, ByRef matpos() As Variant, ByRef MatVPrecios() As Variant, ByRef noconc As Boolean)
Dim noreg As Long
Dim i As Long
Dim claveems As String
Dim tipomov As Integer
Dim fven As Date
Dim pcupon As Integer
Dim vnominal As Double
Dim tcupon As Double
Dim indice As Long
Dim fven1 As Date
Dim pcupon1 As Integer
Dim vnominal1 As Double
Dim tcupon1 As Double
Dim tvalor As String
Dim temision As String
Dim tserie As String
Dim exitoarch As Boolean

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'el objetivo es comparar la posicion con los datos del vector de precios
'si encuentra alguna inconsistencia lo notifica y procede a
'corregir el error
noconc = False
'el objetivo es comparar la posicion con los datos del vector de precios
'si encuentra alguna inconsistencia lo notifica y procede a
'corregir el error
noreg = UBound(matpos, 1)
Call VerificarSalidaArchivo(DirResVaR & "\" & nompos & Format(fecha, "yyyymmdd") & ".txt", 1, exitoarch)
If exitoarch Then
Print #1, "Incidencias encontradas para la posicion del " & Format(fecha, "dd/mm/yyyy")
If noreg <> 0 Then
For i = 1 To noreg
claveems = matpos(i).cEmisionMD
tipomov = matpos(i).Tipo_Mov
If Not IsNull(claveems) Then
If Not IsNull(matpos(i).fVencMD) Then
fven = matpos(i).fVencMD
Else
fven = 0
End If
If Not IsNull(matpos(i).PCuponActSwap) Then
pcupon = matpos(i).PCuponActSwap
Else
pcupon = 0
End If
If Not IsNull(matpos(i).vNominalMD) Then
vnominal = matpos(i).vNominalMD
Else
vnominal = 0
End If
If Not IsNull(matpos(i).tCuponMD) Then
tcupon = matpos(i).tCuponMD
Else
tcupon = 0
End If

indice = 0
fven1 = 0
pcupon1 = 0
vnominal1 = 0
tcupon1 = 0
indice = BuscarValorArray(claveems, MatVPrecios, 17)
If indice <> 0 Then
'se encontro la emision ahora se comparan datos
fven1 = MatVPrecios(indice, 12)
vnominal1 = MatVPrecios(indice, 11)
If Not EsVariableVacia(MatVPrecios(indice, 16)) Then
pcupon1 = MatVPrecios(indice, 16)
Else
pcupon1 = 0
End If
tcupon1 = MatVPrecios(indice, 15)
tvalor = MatVPrecios(indice, 3)
temision = MatVPrecios(indice, 4)
tserie = MatVPrecios(indice, 5)
matpos(i).tValorMD = tvalor
matpos(i).emisionMD = temision
matpos(i).serieMD = tserie

If Val(fven - fven1) <> 0 And fven1 <> 0 And Not IsNull(fven1) Then
 matpos(i).fVencMD = fven1
 Print #1, "no coinciden las fechas de vencimiento para " & matpos(i).cEmisionMD
 MsgBox "no coinciden las fechas de vencimiento para " & matpos(i).cEmisionMD
End If
If Val(vnominal - vnominal1) <> 0 Then
 matpos(i).vNominalMD = vnominal1
 Print #1, "no coincide el valor nominal para " & matpos(i).cEmisionMD
  MsgBox "no coincide el valor nominal para " & matpos(i).cEmisionMD
End If

'If Val(tcupon - tcupon1) <> 0 Then
' matpos(i).tCuponMD  = tcupon1
'End If
Else
 Call MostrarMensajeSistema("Falta " & claveems & " " & i & " en el Vector de Precios PIP", frmProgreso.Label2, 1, Date, Time, NomUsuario)
' MsgBox "Falta " & claveems & " " & i & " en el Vector de Precios PIP"
 Print #1, "Falta " & claveems & " " & i & " en el Vector de Precios PIP"
' noconc =true
End If

End If
  MensajeProc = "Revision de la posicion de mesa"
  AvanceProc = i / noreg
 DoEvents
Next i
Close #1
End If
End If

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function CompFondoPenVPrecios(ByVal fecha As Date, ByRef matpos() As Variant, ByRef MatVPrecios() As Variant, ByRef sierrores As Boolean)
Dim nofilas As Long
Dim noreng As Long
Dim nn As Long
Dim nn1 As Long
Dim i As Long
Dim j As Long
Dim claveems As String
Dim fven As Date
Dim dxv As Integer
Dim toper As Integer
Dim indice As Long
Dim fven1 As Date
Dim pcupon1 As Integer
Dim vnominal1  As Double
Dim tcupon1 As Double
Dim noreg As Integer


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'1 fecha pos
'2 intencion
'3 tipo de operacion
'4 tipo valor
'5 emision
'6 serie
'7 no titulos
'8 fecha vencimiento
'9 valor nominal
'10 plazo cupon
'11 tasa cupon
'12 tasa reporto
'13 precio pactado
'14 propietario
'15 fecha compra
'16 tipo mercado
'17 mesa o posicion
'18 emision

'el objetivo es comparar la posicion con los datos del vector de precios
'si encuentra alguna inconsistencia lo notifica y procede a
'corregir el error
nofilas = UBound(matpos, 1)
noreng = UBound(matpos, 2)
nn1 = UBound(MatVPrecios, 1)
'se le agregan mas columnas a la posicion
ReDim matb(1 To nofilas, 1 To noreng + 6) As Variant
For i = 1 To nofilas
 For j = 1 To noreng
  matb(i, j) = matpos(i, j)
 Next j
Next i

If nn1 > 0 Then
noreg = UBound(matpos, 1)
For i = 1 To noreg
'cpapel = Val(matpos(i, 4))             'clave del papel

claveems = matpos(i, 18)               'clave de la emision
fven = matpos(i, 8)
dxv = fven - fecha
toper = matpos(i, 3)
If Not IsNull(claveems) Then
indice = 0
fven1 = 0
pcupon1 = 0
vnominal1 = 0
tcupon1 = 0
indice = BuscarValorArray(claveems, MatVPrecios, 17)
If toper = ClaveCDirec Then
If indice <> 0 Then
'se encontro la emision ahora se comparan datos
matb(i, 4) = MatVPrecios(indice, 3)     'tipo valor
matb(i, 5) = MatVPrecios(indice, 4)     'emision
matb(i, 6) = MatVPrecios(indice, 5)     'serie
fven1 = CDate(MatVPrecios(indice, 12))  'fecha vencimiento
vnominal1 = Val(MatVPrecios(indice, 11)) 'valor nominal
pcupon1 = Val(MatVPrecios(indice, 16))   'p cupon
tcupon1 = Val(MatVPrecios(indice, 15))   'tasa cupon
If fven <> fven1 Then
 sierrores = True
 matb(i, 8) = fven1
End If
If pcupon1 <> 0 Then
 sierrores = True
 matb(i, 10) = pcupon1
End If
If vnominal1 <> 0 Then
 sierrores = True
 matb(i, 9) = vnominal1
 End If
If tcupon1 <> 0 Then
 sierrores = True
 matb(i, 11) = tcupon1
End If
'matb(i, 14) = matvprecios(indice, 6) 'esta linea es temporal
Else
 MensajeProc = "Falta " & claveems & " de la posición del Fondo de Pensiones"
 
End If
End If
End If
Next i
End If
CompFondoPenVPrecios = matb
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CargaPTCurvas(ByVal txtcurva As String)
Dim i As Long
Dim noreg As Long

Dim base1 As DAO.Database
Dim registros1 As DAO.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

Set base1 = OpenDatabase(DirBases & "\catalogos", dbDriverNoPrompt, False, ";Pwd=" & ContraseñaCatalogos)
Set registros1 = base1.OpenRecordset("select * from [CURVAS Y FACTORES] WHERE CURVA = '" & txtcurva & "' ORDER BY PLAZO", dbOpenDynaset, dbReadOnly)

 registros1.MoveLast
 noreg = registros1.RecordCount
 ReDim mata(1 To noreg) As Long
 registros1.MoveFirst
For i = 1 To noreg
 mata(i) = LeerTAccess(registros1, "PLAZO", i)
 registros1.MoveNext
Next i
registros1.Close
CargaPTCurvas = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub LeerFlujosEmMD(ByRef matpos() As propPosMD, ByRef matfl() As estFlujosMD, ByRef txtmsg As String, ByRef exito As Boolean)
Dim fecha As Date
Dim contar As Long
Dim i As Long
Dim j As Long
Dim ll As Long
Dim noreg As Integer
Dim noreg1 As Integer
Dim claveemision As String
Dim TOPERACION As Integer
Dim tv As String
Dim fven As Date
Dim matfl1() As estFlujosMD

'se cargan los flujos de la deuda, swap o derivado de credito que se tenga
'en posicion
'esta rutina expone una forma de obtener los datos pero es posible obtenerlos de otra forma
txtmsg = ""
If IsArray(matpos) Then
   ReDim matfl(1 To 1)
'ahora se debe determinar donde empieza y termina cada flujo de cada emision
   noreg = UBound(matpos, 1)
   contar = 0
'se verifica si los flujos cargados son consistentes para su uso
   For i = 1 To noreg
       If DetermSiFlujosMD(matpos(i).tValorMD) And (matpos(i).Tipo_Mov = 1 Or matpos(i).Tipo_Mov = 4) Then
          fecha = matpos(i).fechareg
          claveemision = matpos(i).cEmisionMD
          TOPERACION = matpos(i).Tipo_Mov
          tv = matpos(i).tValorMD
          fven = matpos(i).fVencMD
          matfl1 = CFlujosEmisionesMD(fecha, claveemision, False)
          noreg1 = UBound(matfl1, 1)
          If noreg1 <> 0 Then
             matpos(i).iFlujoMD = contar + 1
             matpos(i).fFlujoMD = contar + noreg1
             ReDim Preserve matfl(1 To contar + noreg1)
             For j = 1 To noreg1
                    Set matfl(contar + j) = matfl1(j)
             Next j
             contar = contar + noreg1
          Else
             txtmsg = txtmsg & "No hay flujos para la emision " & claveemision & " en la tabla " & TablaFlujosMD & ","
             MensajeProc = txtmsg
             exito = False
          End If
       End If
       AvanceProc = i / noreg
       MensajeProc = "Cargando los flujos de las emisiones de MD " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   If Len(txtmsg) = 0 Then
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      exito = False
   End If
End If

End Sub

Function DetermSiFlujosMD(ByVal tv As String)
Dim i As Integer
Dim noreg As Integer
noreg = 28
ReDim mata(1 To noreg) As String
mata(1) = "CD"
mata(2) = "D1"
mata(3) = "D2"
mata(4) = "D8"
mata(5) = "F"
mata(6) = "IM"
mata(7) = "IP"
mata(8) = "IQ"
mata(9) = "IS"
mata(10) = "IT"
mata(11) = "JE"
mata(12) = "JI"
mata(13) = "LD"
mata(14) = "LS"
mata(15) = "M"
mata(16) = "M0"
mata(17) = "M7"
mata(18) = "PI"
mata(19) = "S"
mata(20) = "S0"
mata(21) = "XA"
mata(22) = "2U"
mata(23) = "90"
mata(24) = "91"
mata(25) = "92"
mata(26) = "93"
mata(27) = "94"
mata(28) = "95"

DetermSiFlujosMD = False
For i = 1 To noreg
    If tv = mata(i) Then
       DetermSiFlujosMD = True
       Exit Function
    End If
Next i
End Function

Function ObtenerFlujosMD(ByVal inicio As Long, ByVal final As Long, ByRef matfl() As estFlujosMD) As estFlujosMD()
Dim i As Long
Dim matfl1() As New estFlujosMD
ReDim matfl1(1 To final - inicio + 1)
For i = inicio To final
    Set matfl1(i - inicio + 1) = matfl(i)
Next i
ObtenerFlujosMD = matfl1
End Function


Function LeerFlujosPrimSwaps(ByVal txtfiltro As String)
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta funcion es para los swap

rmesa.Open "SELECT COUNT(*) FROM " & txtfiltro, ConAdo
 noreg = rmesa.Fields(0)
rmesa.Close
'se deben de ordenar los flujos en orden ascendente
If noreg <> 0 Then
    Call MostrarMensajeSistema("Cargando los flujos de los swaps: " & noreg & " registros ", frmProgreso.Label2, 0, Date, Time, NomUsuario)
    rmesa.Open "SELECT * FROM " & txtfiltro, ConAdo
    nocampos = rmesa.Fields.Count
    ReDim mata(1 To noreg, 1 To nocampos + 2) As Variant
    For i = 1 To noreg
     For j = 1 To nocampos
       mata(i, j) = rmesa.Fields(j - 1)
     Next j
     mata(i, 13) = "S " & mata(i, 2) & " P " & mata(i, 3)  'clave de la pata
     mata(i, 14) = mata(i, 13) & " " & CLng(mata(i, 5))    'clave de ordenacion de los flujos
     rmesa.MoveNext
     If mata(i, 5) > mata(i, 6) Then MsgBox "La fecha de inicio del periodo es mayor que la fecha final"
     AvanceProc = i / noreg
     MensajeProc = "Cargando flujos de Oracle: " & Format(AvanceProc, "##0.00 %")
     DoEvents
    Next i
    rmesa.Close
    MensajeProc = "Se procede a ordenar los flujos"
    'se ordena por la clave de ordenacion de los flujos
    mata = RutinaOrden(mata, UBound(mata, 2), SRutOrden)
Else
    ReDim mata(0 To 0, 0 To 0) As Variant
    
End If
LeerFlujosPrimSwaps = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerFlujosSwaps(ByVal txtfiltro As String, ByVal simav As Boolean) As estFlujosDeuda()
If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim txtfiltro1 As String
Dim noreg As Long
Dim i As Long
Dim tiempo1 As Date
Dim tiempo2 As Date
Dim tinterval2 As Date

Dim mata() As New estFlujosDeuda
'esta funcion es para los swap
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
RFlujos.Open txtfiltro1, ConAdo
noreg = RFlujos.Fields(0)
RFlujos.Close
'se deben de ordenar los flujos en orden ascendente
tiempo1 = Time
If noreg <> 0 Then
    RFlujos.Open txtfiltro, ConAdo
    ReDim mata(1 To noreg)
    For i = 1 To noreg
        mata(i).tipopos = RFlujos.Fields("TIPOPOS")
        mata(i).fechareg = RFlujos.Fields("FECHAREG")
        mata(i).nompos = RFlujos.Fields("NOMPOS")
        mata(i).horareg = RFlujos.Fields("HORAREG")
        mata(i).coperacion = RFlujos.Fields("COPERACION")
        mata(i).tpata = RFlujos.Fields("TPATA")
        mata(i).finicio = RFlujos.Fields("FINICIO")
        mata(i).ffin = RFlujos.Fields("FFINAL")
        mata(i).fpago = RFlujos.Fields("FDESC")
        mata(i).pago_int = RFlujos.Fields("PAGO_INT")
        mata(i).int_t_saldo = RFlujos.Fields("INT_S_SALDO")
        mata(i).saldo = RFlujos.Fields("SALDO")
        mata(i).amort = RFlujos.Fields("AMORTIZACION")
        mata(i).t_cupon = RFlujos.Fields("TASA")
        RFlujos.MoveNext
        If simav Then
           AvanceProc = i / noreg
           MensajeProc = "Cargando flujos de swaps: " & Format(AvanceProc, "##0.00 %")
           DoEvents
        End If
    Next i
    RFlujos.Close
Else
    ReDim mata(0 To 0)
    MensajeProc = "No hay informacion de flujos"
End If
    tiempo2 = Time
    tinterval2 = tiempo2 - tiempo1
 LeerFlujosSwaps = mata
On Error GoTo 0
Exit Function
hayerror:
MsgBox "Leerflujosswaps " & error(Err())
End Function

Function LeerFlujosDeuda(ByVal txtfiltro As String, ByVal simav As Boolean) As estFlujosDeuda()
Dim noreg As Long
Dim i As Long
Dim mata() As New estFlujosDeuda

'esta funcion es para los swap
RFlujos.Open "SELECT COUNT(*) FROM (" & txtfiltro & ")", ConAdo
noreg = RFlujos.Fields(0)
RFlujos.Close
'se deben de ordenar los flujos en orden ascendente
If noreg <> 0 Then
    RFlujos.Open txtfiltro, ConAdo
    ReDim mata(1 To noreg)
    For i = 1 To noreg
        mata(i).tipopos = RFlujos.Fields("TIPOPOS")
        mata(i).fechareg = RFlujos.Fields("FECHAREG")
        mata(i).nompos = RFlujos.Fields("NOMPOS")
        mata(i).horareg = RFlujos.Fields("HORAREG")
        mata(i).coperacion = RFlujos.Fields("COPERACION")
        mata(i).finicio = RFlujos.Fields("FINICIO")
        mata(i).ffin = RFlujos.Fields("FFINAL")
        mata(i).fpago = RFlujos.Fields("FDESC")
        mata(i).pago_int = RFlujos.Fields("PAGO_INT")
        mata(i).int_t_saldo = RFlujos.Fields("SALDO_INT")
        mata(i).saldo = RFlujos.Fields("SALDO")
        mata(i).amort = RFlujos.Fields("AMORTIZACION")
        mata(i).t_cupon = RFlujos.Fields("TASA")
        'mata(i, 14) = "S " & mata(i, 5)                          'clave de la emision
        'mata(i, 15) = mata(i, 14) & " " & CLng(mata(i, 6))       'clave de ordenacion de los flujos
        RFlujos.MoveNext
        AvanceProc = i / noreg
        If simav Then MensajeProc = "Cargando flujos de Deuda: " & Format(AvanceProc, "##0.00 %")
        DoEvents
    Next i
    RFlujos.Close
Else
    ReDim mata(0 To 0)
    MensajeProc = "No hay informacion de flujos"
End If
LeerFlujosDeuda = mata
End Function


Function CargaParamEmisiones()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta funcion es para los swap
'se carga en memoria una serie de pagos para el calculo
'de flujos del swap
'columnas
txtfiltro1 = "select count(*) from " & PrefijoBD & TablaPEmision
txtfiltro2 = "select * from " & PrefijoBD & TablaPEmision
rmesa.Open txtfiltro1, ConAdo
 noreg = rmesa.Fields(0)
rmesa.Close
'se deben de ordenar los flujos en orden ascendente
If noreg <> 0 Then
rmesa.Open txtfiltro2, ConAdo
ReDim mata(1 To noreg, 1 To 5) As Variant
rmesa.MoveFirst
For i = 1 To noreg
    mata(i, 1) = rmesa.Fields(0)
    mata(i, 2) = rmesa.Fields(1)
    mata(i, 3) = rmesa.Fields(2)
    mata(i, 4) = rmesa.Fields(3)
    mata(i, 5) = rmesa.Fields(4)
    rmesa.MoveNext
    AvanceProc = i / noreg
    
 Call MostrarMensajeSistema("Cargando caracteristicas de emisiones de MD " & Format(AvanceProc, "###0.00 %"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
 DoEvents
Next i
rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
 Call MostrarMensajeSistema("No hay datos de emisiones", frmProgreso.Label2, 2, Date, Time, NomUsuario)
End If
CargaParamEmisiones = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CFlujosEmisionesMD(ByVal fecha As Date, ByVal cemision As String, ByVal simav As Boolean)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim mata() As New estFlujosMD
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta funcion es para los swap
'se carga en memoria una serie de pagos para el calculo
'de flujos del swap
'columnas
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaFlujosMD & " WHERE EMISION = '" & cemision & "' AND FREGISTRO IN (SELECT MAX(FREGISTRO) AS FECHAX FROM " & TablaFlujosMD & " WHERE EMISION = '" & cemision & "' AND FREGISTRO <= " & txtfecha & ") ORDER BY FINICIO"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
 noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg)
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i).c_emision = rmesa.Fields("emision")
       mata(i).finicio = rmesa.Fields("finicio")
       mata(i).ffin = rmesa.Fields("ffinal")
       mata(i).saldo = rmesa.Fields("nocional")
       mata(i).amort = Val(rmesa.Fields("amortizacion"))
       mata(i).tasa = CDbl(rmesa.Fields("tasa"))
       mata(i).p_cupon = rmesa.Fields("pcupon")
       rmesa.MoveNext
       If simav Then
          AvanceProc = i / noreg
          MensajeProc = "Cargando los flujos de las emisiones de MD " & Format(AvanceProc, "###0.00 %")
          DoEvents
       End If
   Next i
   rmesa.Close
Else
 ReDim mata(0 To 0)
End If
CFlujosEmisionesMD = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerFlujoEmisionMD(ByRef flujod() As Variant, ByVal indice1 As Long, ByVal indice2 As Long) As Variant()
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta funcion es para los swap
'cuando se hace referencia a un swap en particular
'se carga en memoria una serie de pagos para el calculo
'de flujos del swap
'los flujos ya ha sido previamente leidos solo se necesita saber donde se ubican para
'cada swap o deuda
'columnas
'1   fecha de vencimiento
'2   monto nocional
'3   spread
'4   subyacente
noreg = indice2 - indice1 + 1
nocampos = UBound(flujod, 2)
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To nocampos) As Variant
For i = indice1 To indice2
 For j = 1 To nocampos
 mata(i - indice1 + 1, j) = flujod(i, j)
 Next j
Next i
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerFlujoEmisionMD = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function


Function LeerUsuariosSistema2()
Dim txtcadena As String
Dim sql_filtro1 As String
Dim sql_filtro2 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

'esta rutina se es la correspondiente para oracle
 txtcadena = "from " & TablaUsuarios & " ORDER BY IDUSUARIO"
 sql_filtro1 = "select count(*) " & txtcadena
 rmesa.Open sql_filtro1, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
 If noreg <> 0 Then
  sql_filtro2 = "select * " & txtcadena
  rmesa.Open sql_filtro2, ConAdo
  ReDim mata(1 To noreg, 1 To 13) As Variant
  For i = 1 To noreg
   mata(i, 1) = rmesa.Fields("IDUSUARIO")
   mata(i, 2) = rmesa.Fields("USUARIO")
   mata(i, 3) = rmesa.Fields("NOMBRE")
   mata(i, 4) = rmesa.Fields("GRUPO")
   mata(i, 5) = rmesa.Fields("ACCESO")
   mata(i, 6) = rmesa.Fields("ENLINEA")
   mata(i, 7) = rmesa.Fields("FENTRADA")
   mata(i, 8) = rmesa.Fields("HENTRADA")
   mata(i, 9) = rmesa.Fields("FSALIDA")
   mata(i, 10) = rmesa.Fields("HSALIDA")
   mata(i, 11) = rmesa.Fields("FUREPORTE")
   mata(i, 12) = rmesa.Fields("HUREPORTE")
   mata(i, 13) = rmesa.Fields("DIRECCION_IP")
  '7 no de intento de acceso del usuario i
   rmesa.MoveNext
  Next i
  rmesa.Close
  mata = RutinaOrden(mata, 1, SRutOrden)
 Else
  ReDim mata(0 To 0, 0 To 0) As Variant
 End If
LeerUsuariosSistema2 = mata
End Function

Function LeerCaractSwapsIKOS(ByVal fecha As String, ByVal txttabla As String, ByRef obj1 As ADODB.Connection, ByRef txtmsg As String, ByRef exito As Boolean)
' de entrada no filtra los registros de la interfase
Dim txtfecha As String
Dim txtfiltro, txtfiltro1 As String
Dim fecha1 As Date
Dim i As Integer
Dim noreg As Integer
Dim c_contrap As Long
Dim fechaa As Date
Dim fechab As Date
Dim rmesa As New ADODB.recordset
Dim RInterfIKOS As New ADODB.recordset
exito = True
txtmsg = ""
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT * FROM " & txttabla
txtfiltro = txtfiltro & " WHERE FECHA_POS = " & txtfecha & " ORDER BY NUMSEC"
txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
rmesa.Open txtfiltro1, obj1
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 RInterfIKOS.Open txtfiltro, obj1
 RInterfIKOS.MoveFirst
 fecha1 = RInterfIKOS.Fields(0)
 If fecha <> fecha1 Then
     RInterfIKOS.Close
     MensajeProc = "No coincide la fecha de la tabla de caracteristicas con la solicitada"
     ReDim mata(0 To 0) As New propPosSwaps
     noreg = 0
     LeerCaractSwapsIKOS = mata
     Exit Function
 End If
 ReDim mata(1 To noreg) As New propPosSwaps
  For i = 1 To noreg
      mata(i).fechareg = RInterfIKOS.Fields("FECHA_POS")                        'FECHA
      mata(i).c_operacion = RInterfIKOS.Fields("NUMSEC")                        'CLAVE DE OPERACION
      fechaa = RInterfIKOS.Fields("FEC_OPERAC")
      fechab = RInterfIKOS.Fields("FEC_INI")
      mata(i).FCompraSwap = Minimo(fechaa, fechab)                              'FECHA DE INICIO
      mata(i).FvencSwap = RInterfIKOS.Fields("FEC_VENC")                        'FECHA DE VENCIMIENTO
      mata(i).RIntAct = RInterfIKOS.Fields("REINVIERTE_INT_ACT")                'R INT ACTIVA
      mata(i).RIntPas = RInterfIKOS.Fields("REINVIERTE_INT_PAS")                'T INT PASIVA
      mata(i).TCActivaSwap = TradTRef(Trim(ReemplazaVacioValor(RInterfIKOS.Fields("TASA_REF_ACTIVA"), "")))               'tasa ref activa
      mata(i).TCPasivaSwap = TradTRef(Trim(ReemplazaVacioValor(RInterfIKOS.Fields("TASA_REF_PASIVA"), "")))               'tasa ref pasiva
      mata(i).STActiva = Val(RInterfIKOS.Fields("SOBRETASA_ACTIVA")) / 100         'sobretasa activa
      mata(i).STPasiva = Val(RInterfIKOS.Fields("SOBRETASA_PASIVA")) / 100        'sobretasa pasiva
      mata(i).C_Posicion = ClavePosDeriv
      mata(i).Tipo_Mov = 1
      c_contrap = DetermEquivContrap(fecha, "" & RInterfIKOS.Fields("CVECONTPP") & "", 1)
      If c_contrap <> 0 Then
         mata(i).ID_ContrapSwap = c_contrap                                          'CLAVE DE contraparte
      Else
         mata(i).ID_ContrapSwap = 0
         exito = False
         txtmsg = txtmsg & "La contraparte Banxico " & RInterfIKOS.Fields("CVECONTPP") & " no esta en el catalogo de SIVARMER,"
      End If
      mata(i).EstructuralSwap = RInterfIKOS.Fields("NEGO_ESTRUC")                        'estructural
      RInterfIKOS.MoveNext
      AvanceProc = i / noreg
      MensajeProc = "Leyendo la tabla de caracteristicas de los swaps del dia " & fecha & " " & Format(AvanceProc, "##0.00 %")
      DoEvents
 Next i
 RInterfIKOS.Close
 If EsVariableVacia(txtmsg) Then txtmsg = "El proceso finalizo correctamente"
Else
    ReDim mata(0 To 0) As New propPosSwaps
    txtmsg = "No hay registros en la tabla de caracteristicas"
End If
'se ordena por la clave de ikos
LeerCaractSwapsIKOS = mata
End Function

Function TradTRef(ByVal txttref As String)
If txttref = "TIIE 28" Then
   TradTRef = "TIIE28[0]"
ElseIf txttref = "TIIE28" Then
   TradTRef = "TIIE28[0]"
ElseIf txttref = "TIIE28[0" Then
   TradTRef = "TIIE28[0]"
ElseIf txttref = "TIIE91[0" Then
   TradTRef = "TIIE91[0]"
ElseIf txttref = "LIBOR 3M" Then
   TradTRef = "LIBOR3M[2]"
ElseIf txttref = "LIBOR3M[2" Then
   TradTRef = "LIBOR3M[2]"
ElseIf txttref = "LIBOR6M[2" Then
   TradTRef = "LIBOR6M[2]"
ElseIf txttref = "LIBOR 6M" Then
   TradTRef = "LIBOR6M[2]"
Else
   TradTRef = txttref
End If
End Function

Function LeerValSwapsIKOS(ByVal fecha As String, ByVal txttabla As String, ByRef obj1 As ADODB.Connection) As Variant()
' de entrada no filtra los registros de la interfase
Dim txtfecha As String
Dim txtfiltro, txtfiltro1 As String
Dim fecha1 As Date
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset
Dim RInterfIKOS As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT * FROM " & txttabla & " ORDER BY NUMSEC"
txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
rmesa.Open txtfiltro1, obj1
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 RInterfIKOS.Open txtfiltro, obj1
 RInterfIKOS.MoveFirst
 fecha1 = RInterfIKOS.Fields(0)
 If fecha <> fecha1 Then
     RInterfIKOS.Close
     MensajeProc = "No coincide la fecha de la tabla de caracteristicas con la solicitada"
     ReDim mata(0 To 0, 0 To 0) As Variant
     noreg = 0
     LeerValSwapsIKOS = mata
     Exit Function
 End If
 ReDim mata(1 To noreg, 1 To 5) As Variant
  For i = 1 To noreg
      mata(i, 1) = RInterfIKOS.Fields("FECHA_POS")                      'FECHA
      mata(i, 2) = RInterfIKOS.Fields("NUMSEC")                         'CLAVE DE OPERACION
      mata(i, 3) = Trim(RInterfIKOS.Fields("VAL_ACTIVA"))               'tasa ref activa
      mata(i, 4) = Trim(RInterfIKOS.Fields("VAL_PASIVA"))               'tasa ref pasiva
      mata(i, 5) = Val(RInterfIKOS.Fields("MARCA_MERCADO"))
      RInterfIKOS.MoveNext
      AvanceProc = i / noreg
      MensajeProc = "Leyendo la tabla de caracteristicas de los swaps del dia " & fecha & " " & Format(AvanceProc, "##0.00 %")
      DoEvents
 Next i
 RInterfIKOS.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
'se ordena por la clave de ikos
mata = RutinaOrden(mata, 5, SRutOrden)
LeerValSwapsIKOS = mata
End Function

Sub ImpPosSwapsRed(ByVal fecha As Date, ByRef noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Dim txtmsg1 As String
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim matcart() As Variant
Dim matcar() As New propPosSwaps
Dim matcar2() As Variant
Dim matfl() As Variant
Dim matfl2() As Variant
'se leen los flujos y las caracteristicas de la interfase de ikos
matcar = LeerCaractSwapsIKOS(fecha, TablaInterfCarac, conAdoBD, txtmsg1, exito3)   'las caracteristicas de los swaps
matcar2 = LeerCaractenFlujosSwapsIKOS(ByVal fecha, conAdoBD)
If UBound(matcar, 1) <> 0 And UBound(matcar2, 1) <> 0 And exito3 Then
   Call UnirCaractSwaps(fecha, matcar, matcar2, txtmsg, exito1)
   If UBound(matcar, 1) <> 0 And exito1 Then
      Call ActualizarPosSwaps(fecha, matcar, 1, "Real", noreg, exito2)
      exito = exito1 And exito2
      txtmsg = "El proceso finalizo correctamente"
   Else
      exito = False
   End If
Else
  exito = False
  MensajeProc = txtmsg1
  txtmsg = MensajeProc
End If

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub UnirCaractSwaps(ByVal fecha As Date, ByRef matcar1() As propPosSwaps, ByRef matcar2() As Variant, ByRef txtmsg As String, ByRef exito As Boolean)
'objetivo de la funcion: construir una tabla de caracteristicas con los datos de
Dim noreg As Long
Dim i As Long
Dim indice2 As Long
Dim indice3 As Long
   MatDerEstandar = CargaDervEstandar()
   exito = True
   noreg = UBound(matcar1, 1)
   For i = 1 To UBound(matcar1, 1)
       indice2 = BuscarValorArray("" & matcar1(i).c_operacion & " B", matcar2, 9)
       If indice2 <> 0 Then
          matcar1(i).intencion = matcar2(indice2, 3)                      'intencion
          matcar1(i).IntercIFSwap = matcar2(indice2, 4)                   'intercambio inicial de flujos
          matcar1(i).IntercFFSwap = matcar2(indice2, 5)                   'intercambio intermedio y final de flujos
          matcar1(i).STActiva = matcar2(indice2, 6)                       'sobretasa activa
          matcar1(i).ClaveProdSwap = Trim(matcar2(indice2, 8))            'clave del tipo de swap
          matcar1(i).cProdSwapGen = TraduceDerivadoEstandar2(matcar1(i).ClaveProdSwap)
       End If
       indice2 = BuscarValorArray("" & matcar1(i).c_operacion & " C", matcar2, 9)
       If indice2 <> 0 Then
          matcar1(i).STPasiva = matcar2(indice2, 6)                       'sobretasa pasiva
       End If
       indice3 = BuscarValorArray(matcar1(i).ClaveProdSwap, MatDerEstandar, 2)
       If indice3 <> 0 Then
          matcar1(i).ConvIntAct = TraduceConvCalcInt(Trim(MatDerEstandar(indice3, 15)))                     'conv activa
          matcar1(i).ConvIntPas = TraduceConvCalcInt(Trim(MatDerEstandar(indice3, 16)))                     'conv pasiva
       Else
          matcar1(i).ConvIntAct = "Actual/360"
          matcar1(i).ConvIntPas = "Actual/360"
       End If
   Next i
End Sub

Function TraduceConvCalcInt(ByVal texto As String)
If texto = "INTERES" Then
   TraduceConvCalcInt = "Actual/360"
ElseIf texto = "180/360" Then
   TraduceConvCalcInt = "180/360"
ElseIf texto = "INTERES 365" Then
   TraduceConvCalcInt = "Actual/365"
Else
   TraduceConvCalcInt = ""
End If
End Function

Function DetermEquivContrap(ByVal fecha As Date, ByVal idcontrap As String, ByVal tclas As Integer)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim rmesa As New ADODB.recordset
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
If tclas = 1 Then
   txtfiltro2 = "SELECT * FROM " & PrefijoBD & TablaEquivContrap & " WHERE"
   txtfiltro2 = txtfiltro2 & " (id_contrap_BANXICO,FECHA) IN"
   txtfiltro2 = txtfiltro2 & " (SELECT id_contrap_BANXICO,MAX(FECHA) AS FECHA FROM " & PrefijoBD & TablaEquivContrap
   txtfiltro2 = txtfiltro2 & " WHERE FECHA <= " & txtfecha & " and ID_CONTRAP_BANXICO = '" & idcontrap & "' GROUP BY ID_CONTRAP_BANXICO)"
Else
   txtfiltro2 = "SELECT *  FROM " & PrefijoBD & TablaEquivContrap & " WHERE "
   txtfiltro2 = txtfiltro2 & "  (id_contrap_IKOS,FECHA) IN"
   txtfiltro2 = txtfiltro2 & " (SELECT id_contrap_IKOS,MAX(FECHA) AS FECHA FROM " & PrefijoBD & TablaEquivContrap
   txtfiltro2 = txtfiltro2 & " WHERE FECHA <= " & txtfecha & " AND ID_CONTRAP_IKOS = '" & idcontrap & "' GROUP BY id_contrap_IKOS)"
End If
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   DetermEquivContrap = rmesa.Fields("ID_CONTRAP")
   rmesa.Close
Else
   DetermEquivContrap = 0
End If
End Function

Sub RestructurarFlujosSwaps(ByRef matfl() As Variant, ByRef matfl3() As Variant, ByRef matcar2() As Variant)
Dim matfl2() As Variant
Dim nocampos As Integer
Dim i As Long
Dim j As Long
Dim matclaves() As Variant
Dim noclaves As Long
Dim contar As Long
Dim nodatos As Long
Dim kk As Long
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim matem() As Variant

 nocampos = UBound(matfl, 2)
 ReDim matfl1(1 To UBound(matfl, 1), 1 To nocampos + 2) As Variant
 For i = 1 To UBound(matfl, 1)
  For j = 1 To nocampos
   matfl1(i, j) = matfl(i, j)
  Next j
 Next i
 For i = 1 To UBound(matfl, 1)
 'clave para identificar la pata: Clave de operación+clave pata
  matfl1(i, nocampos + 1) = "S " & matfl(i, 2) & "P " & matfl(i, 3) 'clave para identificar la pata
 'clave para ordenar los flujos: Clave de operación+clave pata+fecha inicio flujo
  matfl1(i, nocampos + 2) = "S " & matfl(i, 2) & "P " & matfl(i, 3) & " F " & Format(matfl(i, 6), "#######") 'clave para ordenar los flujos
 Next i
'se ordenan los campos de acuerdo a la segunda clave de ordenacion
matfl1 = RutinaOrden(matfl1, nocampos + 2, SRutOrden)
'se obtiene las claves unicas por pata
matclaves = ObtFactUnicos(matfl1, UBound(matfl1, 2) - 1)
noclaves = UBound(matclaves, 1)
'se debe de proceder a filtrar en la matriz solo los registros que corresponden a la pata
contar = 0
nodatos = 0
ReDim matc(1 To UBound(matfl1, 2), 1 To 1) As Variant
For i = 1 To noclaves   'esta rutina es ineficiente
For j = 1 To UBound(matfl1, 1)
 If matclaves(i, 1) = matfl1(j, UBound(matfl1, 2) - 1) Then
  contar = contar + 1
 End If
Next j

ReDim matsal(1 To contar, 1 To UBound(matfl1, 2)) As Variant
contar = 0
For j = 1 To UBound(matfl1, 1)
 If matclaves(i, 1) = matfl1(j, UBound(matfl1, 2) - 1) Then
  contar = contar + 1
  For kk = 1 To UBound(matfl1, 2)
   matsal(contar, kk) = matfl1(j, kk)
  Next kk
 End If
Next j


nodatos = nodatos + contar
ReDim Preserve matc(1 To UBound(matfl1, 2), 1 To nodatos) As Variant
 For j = 1 To contar
  For kk = 1 To UBound(matfl1, 2)
   matc(kk, nodatos - contar + j) = matsal(j, kk)
  Next kk
 Next j
 AvanceProc = i / noclaves
 MensajeProc = "Estructurando los flujos: " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i
matfl2 = MTranV(matc)
ReDim matfl3(1 To UBound(matfl2, 1), 1 To 10) As Variant
For i = 1 To UBound(matfl2, 1)
    matfl3(i, 1) = matfl2(i, 2)    'clave de la operacion
    matfl3(i, 2) = matfl2(i, 3)    'pata
    matfl3(i, 3) = matfl2(i, 6)    'fecha de inicio
    matfl3(i, 4) = matfl2(i, 7)    'fecha final
    matfl3(i, 5) = matfl2(i, 19)   'fecha de pago
    matfl3(i, 6) = matfl2(i, 21)   'paga intereses
    matfl3(i, 7) = True            'saldo*intereses
    matfl3(i, 8) = matfl2(i, 8)    'saldo
    matfl3(i, 9) = matfl2(i, 9)    'amortizacion
    If Not EsVariableVacia(matfl2(i, 10)) Then 'tasa cupon
       If Val(matfl2(i, 10)) <> 0 Then
          matfl3(i, 10) = Val(matfl2(i, 10)) / 100
       Else
          matfl3(i, 10) = 0
       End If
    Else
       matfl3(i, 10) = 0
    End If
Next i

'ahora se tiene que obtener las emisiones que hay en la matriz de flujos
matem = ObtFactUnicos(matfl2, 2)
ReDim matcar2(1 To UBound(matem, 1), 1 To 11) As Variant
For i = 1 To UBound(matem, 1)
    matcar2(i, 1) = "" & matem(i, 1) & ""
    exito1 = False
    exito2 = False
    For j = 1 To UBound(matfl2, 1)
        If matem(i, 1) = matfl2(j, 2) Then
           matcar2(i, 2) = matfl2(j, 17)   'intercambio inicial
           matcar2(i, 3) = matfl2(j, 18)   'intercambio intermedio y final
           matcar2(i, 4) = matfl2(j, 15)   'funcion de valuacion
           matcar2(i, 5) = matfl2(j, 5)    'intencion
           If matfl2(j, 3) = "B" Or matfl2(j, 3) = "A" Then
              If Not EsVariableVacia(matfl2(j, 10)) Then
                 matcar2(i, 6) = Trim(matfl2(j, 10))            'tasa cupon activa
              Else
                 matcar2(i, 6) = 0
              End If
              matcar2(i, 8) = Val(Trim(matfl2(j, 11))) / 100    'sobretasa cupon activa
              matcar2(i, 10) = matfl2(j, 13)                    'conv calculo activa
              exito1 = True
           End If
           If matfl2(j, 3) = "C" Or matfl2(j, 3) = "P" Then     'tasa cupon pasiva
              If Not EsVariableVacia(matfl2(j, 10)) Then
                 matcar2(i, 7) = Trim(matfl2(j, 10))
              Else
                 matcar2(i, 7) = 0
              End If
              matcar2(i, 9) = Val(Trim(matfl2(j, 11))) / 100    'sobretasa cupon activa
              matcar2(i, 11) = matfl2(j, 13)                    'conv calculo activa
              exito2 = True
           End If
           If exito1 And exito2 Then Exit For
        End If
    Next j
Next i
End Sub

Function Conversion1Flujos(mata)
Dim noreg As Long
Dim i As Long

noreg = UBound(mata, 1)
ReDim matx(1 To noreg, 1 To 15) As Variant
For i = 1 To noreg
 matx(i, 1) = mata(i, 1)                   'fecha posicion
 matx(i, 2) = mata(i, 3)                   'Clave de operación
 matx(i, 3) = mata(i, 4)                   'ACTIVA PASIVA O PRIMARIA
 matx(i, 4) = mata(i, 5)                   'intencion
 matx(i, 5) = mata(i, 6)                   'INICIO DE FLUJO
 matx(i, 6) = mata(i, 7)                   'FINAL DE FLUJO
 matx(i, 7) = Val(mata(i, 8))              'SALDO
 If Not EsVariableVacia(mata(i, 9)) Then        'AMORTIZACION
    matx(i, 8) = Val(Trim(mata(i, 9)))
 Else
    matx(i, 8) = 0
 End If
 matx(i, 9) = Trim(mata(i, 10))            'TASA INTERES
 If Not EsVariableVacia(mata(i, 11)) Then
    matx(i, 10) = Val(Trim(mata(i, 11))) / 100 'SPREAD
 Else
    matx(i, 10) = 0
 End If
 matx(i, 11) = mata(i, 12)                 'periodo cupon
 matx(i, 12) = mata(i, 13)                 'CONVENCION CALC INT
 matx(i, 13) = mata(i, 15)                 'clave del producto
 matx(i, 14) = mata(i, 17)                 'estructural
 matx(i, 15) = mata(i, 18)                 'hora de la captura
 Next i
Conversion1Flujos = matx
End Function

Function Conversion2Flujos(mata)
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long

noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
 ReDim matx(1 To noreg, 1 To nocampos + 3) As Variant
 For i = 1 To noreg
  For j = 1 To nocampos
   matx(i, j) = mata(i, j)
  Next j
 Next i
 For i = 1 To noreg
     matx(i, nocampos + 1) = "HR " & mata(i, 15) & "S " & mata(i, 2)                     'clave para identificar LA OPERACION
     matx(i, nocampos + 2) = "HR " & mata(i, 15) & "S " & mata(i, 2) & "P " & mata(i, 3) 'clave para identificar la pata
     matx(i, nocampos + 3) = "HR " & mata(i, 15) & "S " & mata(i, 2) & "P " & mata(i, 3) & " F " & Format(mata(i, 5), "#######") 'clave para ordenar los flujos
 Next i
Conversion2Flujos = matx
End Function

Function ObtenerCaractFlujos2(ByRef mata() As Variant) As Variant()
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim matem() As Variant
Dim noem As Long


noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
matem = ObtFactUnicos(mata, 2)
noem = UBound(matem, 1)
ReDim matb(1 To noem, 1 To 19) As Variant

For i = 1 To noem
    For j = 1 To noreg
        If matem(i, 1) = mata(j, 2) Then
           matb(i, 1) = mata(j, 1)                               'fecha de la posicion
           matb(i, 2) = mata(j, 2)                               'Clave de operación
           matb(i, 3) = 1                                        'posicion activa
           matb(i, 4) = mata(j, 4)                               'intencion
           If matb(i, 5) = 0 Then matb(i, 5) = mata(j, 5)        'fecha de inicio
           matb(i, 6) = Maximo(Val(matb(i, 6)), mata(j, 6))      'fecha final
           matb(i, 7) = True                                     'intercambio inicial de flujos
           matb(i, 8) = True                                     'intercambio final de flujos
           If mata(j, 3) = "B" Then
              matb(i, 9) = mata(j, 9)          'tasa activa
              matb(i, 11) = mata(j, 10)        'sobretasa activa
              matb(i, 13) = mata(j, 12)        'conv intereses activa
           End If
           If mata(j, 3) = "C" Then
              matb(i, 10) = mata(j, 9)         'tasa pasiva
              matb(i, 12) = mata(j, 10)        'sobretasa pasiva
              matb(i, 14) = mata(j, 12)        'conv intereses pasiva
           End If
           matb(i, 15) = False                   'dias comerciales activa
           matb(i, 16) = False                   'dias comerciales pasiva
           matb(i, 17) = mata(j, 13)           'clave del producto
           matb(i, 18) = "000000"              'contraparte
           matb(i, 19) = mata(i, 14)           'estructural
        End If
    Next j
Next i
ObtenerCaractFlujos2 = matb
End Function

Sub ImpPosSwapsIDRed(ByVal fecha As Date, ByVal txtnompos As String, ByRef noregn As Integer, ByRef noregc As Integer)
Dim mata() As Variant
Dim matx() As Variant
Dim matcar() As propPosSwaps
Dim matfl1() As Variant
Dim matfl() As Variant
Dim noreg As Long
Dim noreg1 As Long
Dim i As Long
Dim j As Long
Dim exito As Boolean
Dim horareg As String
'esta rutina debe de sobreescribir las operaciones repetidas y
'reemplazarlas con las operaciones mas recientes
noreg = 0
Call IniciarConexOracle(conAdo2, BDIKOS)
mata = LeerOperSwapsIIKOS(fecha, conAdo2)
If UBound(mata, 1) > 0 Then
   For i = 1 To UBound(mata, 1)
       matcar = LeerCaracSwapsSimIKOS(fecha, mata(i, 1), mata(i, 2), conAdo2)
       matfl = LeerFlujosSwapSimIKOS(fecha, mata(i, 1), mata(i, 2), conAdo2)
       noreg1 = UBound(matcar, 1)
       If noreg1 <> 0 Then
          matfl1 = EstructurarFlujosSwaps4(matfl)
   'se obtiene la lista de swaps a incluir en la posicion
          noreg = UBound(matcar, 1)
          For j = 1 To noreg
              If matcar(j).intencion = "N" Then
                 noregn = noregn + 1
              Else
                 noregc = noregc + 1
              End If
          Next j
          Call GuardarPosSwaps(fecha, 3, txtnompos, mata(i, 2), matcar, matfl1, exito)
       End If
   Next i
Else
   MensajeProc = "no hay posicion en la interfaz de swaps"
End If
conAdo2.Close
End Sub

Function EstructurarFlujosSwaps4(ByRef mata() As Variant) As Variant()
Dim noreg As Long
Dim i As Long
Dim contar As Long
Dim matb() As Variant
noreg = UBound(mata, 1)
ReDim matb(1 To 10, 1 To 1)
For i = 1 To noreg
    If i < noreg Then
       If mata(i, 8) <> 0 Then
          contar = contar + 1
          ReDim Preserve matb(1 To 10, 1 To contar)
          matb(1, contar) = mata(i, 1)             'clave de operacion
          matb(2, contar) = mata(i, 2)             'posicion
          matb(3, contar) = mata(i, 3)             'fecha inicial del flujo
          matb(4, contar) = mata(i, 4)             'fecha final del flujos
          matb(5, contar) = mata(i, 5)             'fecha de descuento de flujos
          matb(6, contar) = "S"
          matb(7, contar) = "S"
          matb(8, contar) = mata(i, 8)             'saldo
          matb(9, contar) = mata(i + 1, 9)         'amortizacion
          matb(10, contar) = mata(i + 1, 10)       'tasa cupon
       End If
    End If
Next i
If contar <> 0 Then
   matb = MTranV(matb)
Else
   ReDim matb(0 To 0, 0 To 0) As Variant
End If
EstructurarFlujosSwaps4 = matb
End Function

Function EstructurarFlujosSwaps(ByRef mata() As Variant) As Variant()
Dim matfl1() As Variant
Dim matb() As Variant
Dim noreg1 As Long
Dim nocampos As Integer
Dim matclaves() As Variant
Dim noclaves As Integer
Dim contar As Integer
Dim nodatos As Integer
Dim i As Long
Dim j As Long
Dim kk As Long

matb = Conversion1Flujos(mata)
noreg1 = UBound(matb, 1)
nocampos = UBound(matb, 2)
matfl1 = Conversion2Flujos(matb)
'clave 1: horareg, no de swap, posicion
'clave 2: horareg, no de swap, posicion, fecha de inicio de flujo
'se ordenan los campos de acuerdo a la ULTIMA  clave de ordenacion
matfl1 = RutinaOrden(matfl1, UBound(matfl1, 2), SRutOrden)
'a continuacion se obienen las claves de todas las patas en la posicion
matclaves = ObtFactUnicos(matfl1, nocampos + 2)
noclaves = UBound(matclaves, 1)
'se debe de proceder a filtrar en la matriz solo los registros que corresponden a la pata
contar = 0
nodatos = 0
ReDim matfl3(1 To nocampos + 3, 1 To 1) As Variant
For i = 1 To noclaves
For j = 1 To noreg1
 If matclaves(i, 1) = matfl1(j, nocampos + 2) Then
    contar = contar + 1
 End If
Next j
ReDim matfl2(1 To contar, 1 To nocampos + 3) As Variant
contar = 0
For j = 1 To noreg1
 If matclaves(i, 1) = matfl1(j, nocampos + 2) Then
  contar = contar + 1
  For kk = 1 To nocampos + 3
      matfl2(contar, kk) = matfl1(j, kk)
  Next kk
 End If
Next j
'se realiza el desplazamiento de las amortizaciones de los swaps
For j = 1 To contar - 1
 matfl2(j, 8) = matfl2(j + 1, 8)     'AMORTIZACION
Next j
nodatos = nodatos + contar - 1
ReDim Preserve matfl3(1 To nocampos + 3, 1 To nodatos) As Variant
For j = 1 To contar - 1
  For kk = 1 To nocampos + 2
   matfl3(kk, nodatos - contar + 1 + j) = matfl2(j, kk)
  Next kk
Next j
Next i
EstructurarFlujosSwaps = MTranV(matfl3)
End Function

Function LeerFlujosSwapsIKOS(ByVal fecha As Date, ByVal coperacion As String, ByRef obj1 As ADODB.Connection) As Variant()
Dim noreg As Long
Dim noreg1 As Long
Dim nocampos As Integer
' de entrada no filtra los registros de la interfase
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim tiempo As Double
Dim fecha1 As Date
Dim i As Long
Dim j As Integer
Dim txtmsg As String
Dim rmesa As New ADODB.recordset
Dim RInterfIKOS As New ADODB.recordset

'====================================================
'atencion esta conexion debe de estar viva aun cuando las bases sean de access
txtfiltro = "select * from " & TablaInterfFlujos & "  WHERE CVESWAP = " & coperacion & " ORDER BY POSICION,FEC_INI"
txtfiltro1 = "SELECT count(*) from (" & txtfiltro & ")"
rmesa.Open txtfiltro1, obj1
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   RInterfIKOS.Open txtfiltro, obj1
   ReDim mata(1 To noreg, 1 To 10) As Variant
   RInterfIKOS.MoveFirst
   fecha1 = RInterfIKOS.Fields(0)
   If fecha <> fecha1 Then
      RInterfIKOS.Close
      txtmsg = "No coincide la fecha de la tabla de flujos con la solicitada"
      ReDim mata(0 To 0, 0 To 0) As Variant
      noreg = 0
      LeerFlujosSwapsIKOS = mata
      Exit Function
   End If
   For i = 1 To noreg
     mata(i, 1) = RInterfIKOS.Fields("CVESWAP")          'clave de operacion
     mata(i, 2) = RInterfIKOS.Fields("POSICION")         'activa o pasiva
     mata(i, 3) = RInterfIKOS.Fields("FEC_INI")                    'f inicio
     mata(i, 4) = RInterfIKOS.Fields("FEC_TER")
     mata(i, 5) = RInterfIKOS.Fields("FEC_LIQ")
     mata(i, 6) = RInterfIKOS.Fields("PAGA_INT")         'paga intereses en el periodo
     mata(i, 7) = "S"
     mata(i, 8) = Val(RInterfIKOS.Fields("VALOR_NOC"))        'saldo
     mata(i, 9) = Val(RInterfIKOS.Fields("AMORTIZACION"))     'amort
     mata(i, 10) = Val(RInterfIKOS.Fields("TASA")) / 100          'tasa
     RInterfIKOS.MoveNext
     AvanceProc = i / noreg
     MensajeProc = "Leyendo los flujos de swaps de la interfase " & Format(AvanceProc, "##0.00 %")
     DoEvents
 Next i
 RInterfIKOS.Close
 LeerFlujosSwapsIKOS = mata
End If
Exit Function
hayerror:
End Function

Function LeerCaractenFlujosSwapsIKOS(ByVal fecha As Date, ByRef obj1 As ADODB.Connection) As Variant()
Dim noreg As Long
Dim noreg1 As Long
Dim nocampos As Integer
Dim txtfecha As String
' de entrada no filtra los registros de la interfase
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfiltro3 As String
Dim tiempo As Double
Dim fecha1 As Date
Dim i As Long
Dim j As Integer
Dim txtmsg As String
Dim rmesa As New ADODB.recordset
Dim RInterfIKOS As New ADODB.recordset

'====================================================
'atencion esta conexion debe de estar viva aun cuando las bases sean de access
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "select FECHA_POS,CVESWAP,POSICION,TIP_NEGOC,INTERI,INTERF,SPREAD,CONVENCION,TIP_SWAP  from " & TablaInterfFlujos
txtfiltro2 = txtfiltro2 & " WHERE FECHA_POS = " & txtfecha
txtfiltro2 = txtfiltro2 & " GROUP BY FECHA_POS,CVESWAP,POSICION,TIP_NEGOC,INTERI,INTERF,SPREAD,CONVENCION,TIP_SWAP"
txtfiltro1 = "SELECT count(*) from (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, obj1
noreg = rmesa.Fields(0)
rmesa.Close
tiempo = Time
Do While Time < (tiempo + 0.001) And noreg <> 0
Loop
rmesa.Open txtfiltro1, obj1
noreg1 = rmesa.Fields(0)
rmesa.Close
If noreg <> noreg1 Then
   ReDim mata(0 To 0, 0 To 0) As Variant
   LeerCaractenFlujosSwapsIKOS = mata
   Exit Function
End If
If noreg <> 0 Then
   RInterfIKOS.Open txtfiltro2, obj1
   RInterfIKOS.MoveFirst
   fecha1 = RInterfIKOS.Fields(0)
   If fecha <> fecha1 Then
      RInterfIKOS.Close
      txtmsg = "No coincide la fecha de la tabla de flujos con la solicitada"
      ReDim mata(0 To 0, 0 To 0) As Variant
      noreg = 0
      LeerCaractenFlujosSwapsIKOS = mata
      Exit Function
   End If
   ReDim mata(1 To noreg, 1 To 16) As Variant
   For i = 1 To noreg
       mata(i, 1) = RInterfIKOS.Fields("CVESWAP")              'clave de operacion
       mata(i, 2) = RInterfIKOS.Fields("POSICION")             'pata
       mata(i, 3) = RInterfIKOS.Fields("TIP_NEGOC")            'intencion
       mata(i, 4) = RInterfIKOS.Fields("INTERI")               'intercambio inicial de flujos
       mata(i, 5) = RInterfIKOS.Fields("INTERF")               'intercambio de flujos
       mata(i, 6) = Val(RInterfIKOS.Fields("SPREAD")) / 100    'SOBRETASA
       mata(i, 7) = RInterfIKOS.Fields("CONVENCION")           'convencion de intereses
       mata(i, 8) = RInterfIKOS.Fields("TIP_SWAP")             'clave de producto
       mata(i, 9) = mata(i, 1) & " " & mata(i, 2)
       RInterfIKOS.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo los flujos de swaps de la interfase " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   mata = RutinaOrden(mata, 9, SRutOrden)
   RInterfIKOS.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerCaractenFlujosSwapsIKOS = mata
Exit Function
hayerror:
End Function

Function LeerOperSwapsIIKOS(ByVal fecha As Date, ByRef conex As ADODB.Connection) As Variant()
' de entrada no filtra los registros de la interfase
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim j As Long
Dim ll As Long
Dim nocampos As Integer
Dim mata() As Variant
Dim RInterfIKOS As New ADODB.recordset

'====================================================
'esta txtbase fue creada por sistemas
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT CVESWAP,HORAREG "
txtfiltro2 = txtfiltro2 & " FROM " & TablaInterfSim1 & " GROUP BY CVESWAP,HORAREG"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
RInterfIKOS.Open txtfiltro1, conex
 noreg = RInterfIKOS.Fields(0)
RInterfIKOS.Close
If noreg <> 0 Then
   RInterfIKOS.Open txtfiltro2, conex
   nocampos = RInterfIKOS.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos + 2) As Variant
   RInterfIKOS.MoveFirst
   For j = 1 To noreg
       mata(j, 1) = RInterfIKOS.Fields("CVESWAP")
       mata(j, 2) = RInterfIKOS.Fields("HORAREG")
       RInterfIKOS.MoveNext
   Next j
   RInterfIKOS.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerOperSwapsIIKOS = mata
End Function

Function LeerCaracSwapsSimIKOS(ByVal fecha As Date, ByVal coperacion As String, ByVal horareg As String, ByRef conex As ADODB.Connection) As propPosSwaps()
' de entrada no filtra los registros de la interfase
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim j As Long
Dim ll As Long
Dim nocampos As Integer
Dim mata() As New propPosSwaps
Dim RInterfIKOS As New ADODB.recordset
'====================================================
'esta txtbase fue creada por sistemas
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT POSICION,TIP_NEGOC,MAX(FEC_TER) AS FVENC,TASA,SPREAD,CONVENCION,TIP_SWAP,ESTRUCTURAL"
txtfiltro2 = txtfiltro2 & " FROM " & TablaInterfSim1 & " WHERE CAST(CVESWAP AS VARCHAR2(20)) = '" & coperacion & "'"
txtfiltro2 = txtfiltro2 & " AND CAST(HORAREG AS VARCHAR2(20)) = '" & horareg & "'"
txtfiltro2 = txtfiltro2 & " GROUP BY POSICION,TIP_NEGOC,TASA,SPREAD,CONVENCION,TIP_SWAP,ESTRUCTURAL"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
RInterfIKOS.Open txtfiltro1, conex
 noreg = RInterfIKOS.Fields(0)
RInterfIKOS.Close
If noreg = 2 Then
   RInterfIKOS.Open txtfiltro2, conex
   nocampos = RInterfIKOS.Fields.Count
   ReDim mata(1 To 1) As New propPosSwaps
   RInterfIKOS.MoveFirst
   For j = 1 To noreg
       mata(1).c_operacion = coperacion
       mata(1).intencion = RInterfIKOS.Fields("TIP_NEGOC")
       mata(1).FvencSwap = RInterfIKOS.Fields("FVENC")
       mata(1).ClaveProdSwap = RInterfIKOS.Fields("TIP_SWAP")
       mata(1).cProdSwapGen = TraduceDerivadoEstandar2(mata(1).ClaveProdSwap)
       mata(1).EstructuralSwap = RInterfIKOS.Fields("ESTRUCTURAL")
       mata(1).C_Posicion = 4
       mata(1).FCompraSwap = fecha
       mata(1).ID_ContrapSwap = 0
       mata(1).IntercIFSwap = "S"
       mata(1).IntercFFSwap = "S"
       mata(1).RIntAct = "N"
       mata(1).RIntPas = "N"
       mata(1).Signo_Op = 1
       mata(1).Tipo_Mov = "1"
       If RInterfIKOS.Fields("POSICION") = "B" Then
          mata(1).TCActivaSwap = Trim(RInterfIKOS.Fields("TASA"))
          If Val(mata(1).TCActivaSwap) = 0 Then
             mata(1).STActiva = Val(Trim(ReemplazaVacioValor(RInterfIKOS.Fields("SPREAD"), 0))) / 100
          Else
             mata(1).STActiva = 0
          End If
          mata(1).ConvIntAct = RInterfIKOS.Fields("CONVENCION")
       ElseIf RInterfIKOS.Fields("POSICION") = "C" Then
          mata(1).TCPasivaSwap = Trim(RInterfIKOS.Fields("TASA"))
          If Val(mata(1).TCPasivaSwap) = 0 Then
             mata(1).STPasiva = Val(Trim(ReemplazaVacioValor(RInterfIKOS.Fields("SPREAD"), 0))) / 100
          Else
             mata(1).STPasiva = 0
          End If
          mata(1).ConvIntPas = RInterfIKOS.Fields("CONVENCION")
       End If
       RInterfIKOS.MoveNext
   Next j
   RInterfIKOS.Close
Else
   ReDim mata(0 To 0) As New propPosSwaps
End If
LeerCaracSwapsSimIKOS = mata
End Function

Function LeerFlujosSwapSimIKOS(ByVal fecha As Date, ByVal coperacion As String, ByVal horareg As String, ByRef conex As ADODB.Connection) As Variant()
' de entrada no filtra los registros de la interfase
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim j As Long
Dim ll As Long
Dim nocampos As Integer
Dim mata() As Variant
Dim RInterfIKOS As New ADODB.recordset

'====================================================
'esta txtbase fue creada por sistemas
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaInterfSim1 & " WHERE CAST(CVESWAP AS VARCHAR2(20)) = '" & coperacion & "'"
txtfiltro2 = txtfiltro2 & " AND CAST(HORAREG AS VARCHAR2(20)) = '" & horareg & "'"
txtfiltro2 = txtfiltro2 & " ORDER BY POSICION,FEC_INI"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
RInterfIKOS.Open txtfiltro1, conex
 noreg = RInterfIKOS.Fields(0)
RInterfIKOS.Close
If noreg <> 0 Then
   RInterfIKOS.Open txtfiltro2, conex
   ReDim mata(1 To noreg, 1 To 12) As Variant
   RInterfIKOS.MoveFirst
   For j = 1 To noreg
       mata(j, 1) = RInterfIKOS.Fields("CVESWAP")
       mata(j, 2) = RInterfIKOS.Fields("POSICION")
       mata(j, 3) = RInterfIKOS.Fields("FEC_INI")
       mata(j, 4) = RInterfIKOS.Fields("FEC_TER")
       mata(j, 5) = RInterfIKOS.Fields("FEC_TER")
       mata(j, 6) = "S"
       mata(j, 7) = "S"
       mata(j, 8) = RInterfIKOS.Fields("VALOR_NOC")
       mata(j, 9) = RInterfIKOS.Fields("AMORTIZACION")
       mata(j, 10) = Val(Trim(RInterfIKOS.Fields("TASA"))) / 100
       RInterfIKOS.MoveNext
   Next j
   RInterfIKOS.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerFlujosSwapSimIKOS = mata
End Function

Function LeerInterfSwapsIKOS2(ByVal fecha As Date, ByVal intencion As String, ByRef conex As ADODB.Connection) As Variant()
' de entrada no filtra los registros de la interfase
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim nocampos As Long
Dim j As Long
Dim mata() As Variant
Dim RInterfIKOS As New ADODB.recordset
'====================================================
'esta txtbase fue creada por sistemas
txtfecha = Format(fecha, "dd/mm/yy")
txtfiltro2 = "SELECT CVESWAP, TIP_NEGOC, TIP_SWAP, ESTATUS, HORAREG,ESTRUCTURAL FROM " & TablaInterfSim1 & " WHERE TIP_NEGOC = '" & intencion & "' GROUP BY CVESWAP, HORAREG, TIP_NEGOC, TIP_SWAP, ESTATUS, ESTRUCTURAL ORDER BY CVESWAP,HORAREG"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
RInterfIKOS.Open txtfiltro1, conex
  noreg = RInterfIKOS.Fields(0)
RInterfIKOS.Close
If noreg <> 0 Then
   RInterfIKOS.Open txtfiltro2, conex
   ReDim mata(1 To noreg, 1 To 7) As Variant
   RInterfIKOS.MoveFirst
   For j = 1 To noreg
       mata(j, 1) = RInterfIKOS.Fields(0)   'clave de operacion
       mata(j, 2) = RInterfIKOS.Fields(1)   'intencion
       mata(j, 3) = RInterfIKOS.Fields(2)   'tipo de operacion
       mata(j, 4) = RInterfIKOS.Fields(3)   'estado de la operacion
       mata(j, 5) = RInterfIKOS.Fields(4)   'hora de registro
       mata(j, 6) = RInterfIKOS.Fields(5)   'estructural
       mata(j, 7) = "Swap"

       RInterfIKOS.MoveNext
   Next j
   RInterfIKOS.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerInterfSwapsIKOS2 = mata
End Function

Function ImpFwdSimIkos(ByVal fecha As Date, ByRef noreg As Long) As Variant()
' de entrada no filtra los registros de la interfase
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer

'====================================================
txtfiltro2 = "select * FROM " & TablaInterfSim2
'txtfiltro1 = "select * FROM " & TablaInterfSim2 & " WHERE (CVESWAP as varchar2(12)),cast (HORAREG as varchar2(12))) NOT IN(SELECT CAST(COPERACION AS VARCHAR(12)), CAST(HORAREGIKOS AS VARCHAR2(12)) FROM " & TablaOperValidada & ")"
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"
RFlujos.Open txtfiltro1, conAdo2
noreg = RFlujos.Fields(0)
RFlujos.Close
If noreg <> 0 Then
   RFlujos.Open txtfiltro2, conAdo2
   nocampos = RFlujos.Fields.Count
   ReDim matpos(1 To noreg, 1 To nocampos) As Variant
   RFlujos.MoveFirst
   For i = 1 To noreg
       For j = 1 To nocampos
           matpos(i, j) = RFlujos.Fields(j - 1)
       Next j
       RFlujos.MoveNext
   Next i
    RFlujos.Close
 Else
    ReDim matpos(0 To 0, 0 To 0) As Variant
End If
ImpFwdSimIkos = matpos
End Function

Function ImpFwdSimIkos2(ByVal fecha As Date, ByVal intencion1 As String, ByVal intencion2 As String, ByRef conex As ADODB.Connection) As Variant()
' de entrada no filtra los registros de la interfase
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Integer
txtfecha = Format(fecha, "yyyymmdd")
'====================================================
txtfiltro2 = "select CVESWAP,INTENCION,TIPO_FORWARD,ESTATUS,HORAREG,ESTRUCTURAL FROM " & TablaInterfSim2 & " WHERE INTENCION = '" & intencion1 & "' OR INTENCION = '" & intencion2 & "' ORDER BY CVESWAP, HORAREG"
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"
RFlujos.Open txtfiltro1, conAdo2
noreg = RFlujos.Fields(0)
RFlujos.Close
If noreg <> 0 Then
   RFlujos.Open txtfiltro2, conAdo2
   ReDim matpos(1 To noreg, 1 To 7) As Variant
   RFlujos.MoveFirst
   For i = 1 To noreg
       matpos(i, 1) = RFlujos.Fields(0)      'clave de operacion
       matpos(i, 2) = RFlujos.Fields(1)      'intencion
       matpos(i, 3) = RFlujos.Fields(2)      'tipo de operacion
       matpos(i, 4) = RFlujos.Fields(3)      'estado de la operacion
       matpos(i, 5) = RFlujos.Fields(4)      'hora de registro
       matpos(i, 6) = RFlujos.Fields(5)      'ESTRUCTURAL
       matpos(i, 7) = "Fwd"
       RFlujos.MoveNext
   Next i
    RFlujos.Close
 Else
    ReDim matpos(0 To 0, 0 To 0) As Variant
End If
ImpFwdSimIkos2 = matpos
End Function

Function ImpFwdIkosO(ByVal fecha As Date, ByVal txtbase As String, ByRef noreg As Long, ByRef obj1 As ADODB.Connection, ByRef exito As Boolean)
' de entrada no filtra los registros de la interfase
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim i As Long
Dim j As Integer
Dim fecha1 As Date
Dim rmesa As New ADODB.recordset
Dim RInterfIKOS As New ADODB.recordset

'====================================================
txtfiltro = "SELECT * FROM " & txtbase
txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
rmesa.Open txtfiltro1, obj1
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   RInterfIKOS.Open txtfiltro, obj1
   RInterfIKOS.MoveFirst
   fecha1 = RInterfIKOS.Fields(0)
   If fecha <> fecha1 Then
      RInterfIKOS.Close
      MensajeProc = "No coincide la fecha de la posicion fwds en IKOS Derivados con la solicitada"
      ReDim mata(0 To 0, 0 To 0) As Variant
      noreg = 0
      exito = False
      ImpFwdIkosO = mata
      Exit Function
   End If
   ReDim mata(1 To noreg, 1 To 13) As Variant
   For i = 1 To noreg
       mata(i, 1) = RInterfIKOS.Fields("F_POSICION")
       mata(i, 2) = RInterfIKOS.Fields("INTENCION")
       mata(i, 3) = RInterfIKOS.Fields("CLAVE_OP")
       mata(i, 4) = RInterfIKOS.Fields("T_OPERACION")
       mata(i, 5) = RInterfIKOS.Fields("M_NOCIONAL")
       mata(i, 6) = RInterfIKOS.Fields("F_INICIO")
       mata(i, 7) = RInterfIKOS.Fields("F_VENCIMIENTO")
       mata(i, 8) = RInterfIKOS.Fields("F_LIQUIDACION")
       mata(i, 9) = RInterfIKOS.Fields("TC_PACTADO")
       mata(i, 10) = RInterfIKOS.Fields("CLAVE_PRODUCTO")
       mata(i, 11) = RInterfIKOS.Fields("CONTRAPARTE")
       mata(i, 12) = RInterfIKOS.Fields("CLAVE_PRODUCTO")
       mata(i, 13) = RInterfIKOS.Fields("NEGO_ESTRUC")
       RInterfIKOS.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo la posicion de fwds " & Format(AvanceProc, "#00.00 %")
   Next i
   RInterfIKOS.Close
   exito = True
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
   exito = True
   noreg = 0
End If
ImpFwdIkosO = mata
End Function

Function ImpValFwdIkosO(ByVal fecha As Date, ByVal txtbase As String, ByRef noreg As Long, ByRef obj1 As ADODB.Connection) As Variant()
' de entrada no filtra los registros de la interfase
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim i As Long
Dim j As Integer
Dim fecha1 As Date
Dim rmesa As New ADODB.recordset
Dim RInterfIKOS As New ADODB.recordset

'====================================================
txtfiltro = "SELECT * FROM " & txtbase
txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
rmesa.Open txtfiltro1, obj1
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 RInterfIKOS.Open txtfiltro, obj1
 RInterfIKOS.MoveFirst
 fecha1 = RInterfIKOS.Fields(0)
 If fecha <> fecha1 Then
     RInterfIKOS.Close
     MensajeProc = "No coincide la fecha de la posicion fwds en IKOS Derivados con la solicitada"
     ReDim mata(0 To 0, 0 To 0) As Variant
     noreg = 0
     ImpValFwdIkosO = mata
     Exit Function
 End If
 ReDim mata(1 To noreg, 1 To 3) As Variant
 For i = 1 To noreg
     mata(i, 1) = RInterfIKOS.Fields("F_POSICION")
     mata(i, 2) = RInterfIKOS.Fields("CLAVE_OP")
     mata(i, 3) = RInterfIKOS.Fields("VALUACION")
     RInterfIKOS.MoveNext
     AvanceProc = i / noreg
     MensajeProc = "Leyendo la posicion de fwds " & Format(AvanceProc, "#00.00 %")
 Next i
 RInterfIKOS.Close
Else
  ReDim mata(0 To 0, 0 To 0) As Variant
 noreg = 0
End If
 ImpValFwdIkosO = mata
End Function

Sub ImpPosPrimArc(ByVal fecha As Date, ByVal txtnompos As String, ByVal nomarch As String, ByVal tipopos As Integer, ByRef noreg1 As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim noreg2 As Long
Dim matcar() As New propPosDeuda
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim i As Long
Dim j As Long
Dim sihayarch As Boolean
Dim nocampos As Long
Dim sihaydatos As Boolean
Dim fvalua As String
Dim coper1 As String
Dim coper2 As String
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim indice As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
   Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
   Set registros1 = base1.OpenRecordset("Caract$", dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
   If registros1.RecordCount <> 0 Then
      registros1.MoveLast
      noreg1 = registros1.RecordCount
      registros1.MoveFirst
      nocampos = registros1.Fields.Count
      ReDim matcar(1 To noreg1) As New propPosDeuda
      For i = 1 To noreg1
          matcar(i).fechareg = fecha
          matcar(i).C_Posicion = LeerTAccess(registros1, 0, i)
          matcar(i).c_operacion = LeerTAccess(registros1, 1, i)
          If LeerTAccess(registros1, 2, i) = "A" Then
             matcar(i).Tipo_Mov = 1
          ElseIf LeerTAccess(registros1, 2, i) = "P" Then
             matcar(i).Tipo_Mov = 4
          Else
             matcar(i).Tipo_Mov = 1
          End If
          matcar(i).FinicioDeuda = LeerTAccess(registros1, 3, i)
          matcar(i).FVencDeuda = LeerTAccess(registros1, 4, i)
          matcar(i).InteriDeuda = LeerTAccess(registros1, 5, i)
          matcar(i).InterfDeuda = LeerTAccess(registros1, 6, i)
          matcar(i).RintDeuda = LeerTAccess(registros1, 7, i)
          matcar(i).TRefDeuda = LeerTAccess(registros1, 8, i)
          matcar(i).SpreadDeuda = LeerTAccess(registros1, 9, i)
          matcar(i).ConvIntDeuda = LeerTAccess(registros1, 10, i)
          coper2 = DetCoper(matcar(i).c_operacion)
          txtfecha = "to_date('" & Format(matcar(i).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & " WHERE COPERACION = '" & coper2 & "' AND TIPOPOS = 3 And FECHAREG = " & txtfecha & " ORDER BY HORAREG"
          txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
          rmesa.Open txtfiltro1, ConAdo
          noreg = rmesa.Fields(0)
          rmesa.Close
          If noreg <> 0 Then
             rmesa.Open txtfiltro2, ConAdo
              ReDim matoper(1 To noreg, 1 To 2) As Variant
              For j = 1 To noreg
                  matoper(j, 1) = rmesa.Fields("coperacion")
                  matoper(j, 2) = rmesa.Fields("HORAREG")
                  rmesa.MoveNext
              Next j
              rmesa.Close
              IndOperR = 0
              Do While IndOperR = 0
                 For j = 1 To noreg
                     frmListaOpR.Combo1.AddItem matoper(j, 1) & " " & matoper(j, 2)
                 Next j
                 frmListaOpR.Show 1
              Loop
             fvalua = DetFValxSwapAsociado(tipopos, matcar(i).fechareg, matoper(IndOperR, 1), matoper(IndOperR, 2), matcar(i).Tipo_Mov)
             If Len(Trim(fvalua)) = 0 Then MsgBox "no se definio la formula de valuacion de la pos primaria"
             matcar(i).fValuacion = fvalua
             registros1.MoveNext
          Else
             matcar(i).fValuacion = ""
          End If
      Next i
      registros1.Close
      base1.Close
   End If
 'la matriz de caracteristicas
   Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
   Set registros1 = base1.OpenRecordset("Flujos$", dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
   If registros1.RecordCount <> 0 Then
      registros1.MoveLast
      noreg2 = registros1.RecordCount
      registros1.MoveFirst
      nocampos = registros1.Fields.Count
      ReDim matfl(1 To noreg2, 1 To nocampos + 2) As Variant
      For i = 1 To noreg2
          For j = 1 To nocampos
              matfl(i, j) = LeerTAccess(registros1, j - 1, i)
          Next j
          matfl(i, nocampos + 1) = matfl(i, 1) & matfl(i, 2)
          matfl(i, nocampos + 2) = matfl(i, nocampos + 1) & CLng(matfl(i, 3))
          registros1.MoveNext
      Next i
      registros1.Close
      base1.Close
   End If
   For i = 1 To noreg1
      sihaydatos = False
      For j = 1 To noreg2
      If matcar(i).c_operacion = matfl(j, 1) Then
         sihaydatos = True
         Exit For
      End If
      Next j
   Next i
   If sihaydatos Then
      Call GuardarPosDeuda(tipopos, txtnompos, "000000", matcar, matfl)
      exito = True
      txtmsg = "El proceso finalizo correctamente"
   Else
      exito = False
      txtmsg = "La tabla de caracteristicas no coincide con la tabla de flujos"
   End If
Else
   MsgBox "No hay acceso a: " & nomarch
   txtmsg = "No hay acceso a: " & nomarch
   exito = False
End If

On Error GoTo 0
Exit Sub
ControlErrores:
'MsgBox Err() & " : " & error(Err())
exito = False
If Err() = 94 Then
   txtmsg = "Hay errores en el contenido de las tablas de excel. Revisar"
End If

If EsBaseAbierta(base1) Then
   base1.Close
End If
If EsRecAbierta(registros1) Then
   registros1.Close
End If
On Error GoTo 0
End Sub

Function EsBaseAbierta(ByRef Base As DAO.Database)
On Error GoTo hayerror
If Not EsVariableVacia(Base.Name) Then
EsBaseAbierta = True
Else
EsBaseAbierta = False
End If
On Error GoTo 0
Exit Function
hayerror:
EsBaseAbierta = False
End Function

Function EsRecAbierta(ByRef record As DAO.recordset)
On Error GoTo hayerror
If Not EsVariableVacia(record.Name) Then
EsRecAbierta = True
Else
EsRecAbierta = False
End If
On Error GoTo 0
Exit Function
hayerror:
EsRecAbierta = False
End Function

Function Es_R_ADODB_Op(ByRef record As ADODB.recordset)
On Error GoTo hayerror
If Not EsVariableVacia(record.PageCount) Then
Es_R_ADODB_Op = True
Else
Es_R_ADODB_Op = False
End If
On Error GoTo 0
Exit Function
hayerror:
Es_R_ADODB_Op = False
End Function

Function Es_ADODB_Op(ByRef conex As ADODB.Connection)
On Error GoTo hayerror
If Not EsVariableVacia(conex.Name) Then
   Es_ADODB_Op = True
Else
   Es_ADODB_Op = False
End If
On Error GoTo 0
Exit Function
hayerror:
Es_ADODB_Op = False
End Function

Sub CrearPosSwapsSimArch(ByVal fecha As Date, ByVal nomarch As String, ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horareg As String, ByRef noreg As Long)
On Error GoTo ControlErrores
Dim sihayarch As Boolean
Dim txtclave As String
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim matcar() As New propPosSwaps
Dim matfl() As Variant

txtclave = ""
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
  'la matriz de caracteristicas
   matcar = LeerCaractSwapExcel(nomarch)
   matfl = LeerFlujosSwapExcel(nomarch, exito1)
   noreg = UBound(matcar, 1)
   If noreg > 0 Then
      Call GuardarPosSwaps(fecha, tipopos, txtnompos, horareg, matcar, matfl, exito2)
   Else
      MsgBox "No hay registros en el archivo"
   End If
   If Not exito2 Then MsgBox "El archivo de posicion contiene errores"
Else
   MsgBox "No hay acceso a: " & nomarch
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub CrearPosSwapsIDArc(ByRef fecha As Date, ByVal nomarch As String, ByRef noregn As Long, ByRef noregc As Long)
Dim sihayarch As Boolean
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim txtnompos As String
Dim horareg As String
Dim matcar() As New propPosSwaps
Dim matfl() As Variant

sihayarch = False
Do While Not sihayarch
   frmEjecucionProc.CommonDialog1.ShowOpen
   nomarch = frmEjecucionProc.CommonDialog1.FileName
   sihayarch = VerifAccesoArch(nomarch)
Loop
 'la matriz de flujos
If sihayarch Then
    matcar = LeerCaractSwapExcel(nomarch)
    matfl = LeerFlujosSwapExcel(nomarch, exito1)
    Call GuardarPosSwaps(fecha, 3, txtnompos, horareg, matcar, matfl, exito2)
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ImportarPosFwds(ByVal fecha As Date, ByVal siarch As Boolean, ByVal nomdir As String, ByRef noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
If siarch Then
 Call ImpFwdExcel(fecha, nomdir, noreg, txtmsg, exito)
Else
 Call ImpPosFwdRed(fecha, noreg, txtmsg, exito)
End If
End Sub

Sub ImpPosFwdRed(ByVal fecha As Date, ByVal noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim nr As Long
Dim mata() As Variant
Dim i As Long
Dim ccontra As Integer
Dim matpos() As New propPosFwd
Dim matpos1() As New propPosFwd
Dim txtborra As String
Dim txtfecha As String
Dim horareg As String

mata = ImpFwdIkosO(fecha, TablaInterfFwd, noreg, conAdoBD, exito)
If exito Then
   If noreg <> 0 Then
      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtborra = "DELETE FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfecha
      txtborra = txtborra & " AND TIPOPOS  = 1"
      txtborra = txtborra & " AND FECHAREG = " & txtfecha
      ConAdo.Execute txtborra
      ReDim matpos(1 To noreg) As New propPosFwd
      For i = 1 To noreg
          horareg = "000000"
          matpos(i).C_Posicion = ClavePosDeriv           'clave de la posicion
          matpos(i).intencion = mata(i, 2)               'intencion
          matpos(i).c_operacion = mata(i, 3)             'clave de operacion
          If UCase(mata(i, 4)) = "LARGO" Then            'largo o corto
             matpos(i).Tipo_Mov = 1
          Else
             matpos(i).Tipo_Mov = 4
          End If
          matpos(i).MontoNocFwd = mata(i, 5)             'monto nocional
          matpos(i).FCompraFwd = mata(i, 6)              'fecha de inicio
          matpos(i).FVencFwd = mata(i, 7)                'fecha de vencimiento
          matpos(i).PAsignadoFwd = mata(i, 9)            'strike
          matpos(i).ClaveProdFwd = mata(i, 10)           'tipo forward
          ccontra = DetermEquivContrap(fecha, "" & mata(i, 11) & "", 2)
          If ccontra = 0 Then
             txtmsg = "No se encontro la equivalencia de la contraparte " & mata(i, 11)
             exito = False
          End If
          matpos(i).ID_ContrapFwd = ccontra
          If mata(i, 2) = "R" Then              'de negociacion por reclasificacion
             matpos(i).intencion = "N"
             matpos(i).ReclasificaFwd = "S"
             matpos(i).EstructuralFwd = "N"
          Else
             matpos(i).ReclasificaFwd = "N"
             matpos(i).EstructuralFwd = mata(i, 13)          'forwards estructural o no
          End If
          ReDim matpos1(1 To 1) As New propPosFwd
          Set matpos1(1) = matpos(i)
          Call GuardaPosFwds(fecha, 1, "Real", horareg, matpos1, exito)
          AvanceProc = i / noreg
          MensajeProc = "Estructurando la posicion de fwds " & Format(AvanceProc, "##0.00 %")
          DoEvents
      Next i
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      txtmsg = "El proceso finalizo correctamente"
      noreg = 0
      exito = True
   End If
Else
   MensajeProc = "No coincide la fecha de la posicion de fwds"
   exito = False
End If

End Sub

Sub ImpPosFwdFrutos(ByVal fecha As Date, ByRef matb() As Variant)
Dim txtfecha As String
Dim txtborra As String
Dim i As Integer
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtinserta As String

 txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtborra = "DELETE FROM " & TablaPosFwdFrutos & " WHERE NOMPOSICION = '" & Format(fecha, "dd/mm/yyyy") & "'"
 ConAdo.Execute txtborra
 For i = 1 To UBound(matb, 1)
    txtfecha1 = "to_date('" & Format(matb(i, 5), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(matb(i, 6), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha3 = "to_date('" & Format(matb(i, 7), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtinserta = "INSERT INTO " & TablaPosFwdFrutos & " VALUES("
    txtinserta = txtinserta & "'" & Format(fecha, "dd/mm/yyyy") & "',"  'nomposicion
    txtinserta = txtinserta & txtfecha & ","                            'fechapos
    txtinserta = txtinserta & matb(i, 1) & ","                          'coperacion
    txtinserta = txtinserta & matb(i, 2) & ","                          'largo o corto
    txtinserta = txtinserta & "1,"                                      'notitulos
    txtinserta = txtinserta & matb(i, 4) & ","                          'mnocional
    txtinserta = txtinserta & txtfecha1 & ","                           'fecha inicio
    txtinserta = txtinserta & txtfecha2 & ","                           'fecha vencimiento
    txtinserta = txtinserta & txtfecha3 & ","                           'fecha liquidacion
    txtinserta = txtinserta & matb(i, 8) & ","                          'ppactado
    txtinserta = txtinserta & "'" & matb(i, 9) & "',"                   'CPRODUCTO
    txtinserta = txtinserta & "'" & matb(i, 10) & "',"                  'intencion
    txtinserta = txtinserta & matb(i, 11) & ","                         'CONTRAPARTE
    txtinserta = txtinserta & matb(i, 14) & ","                         'V IKOS
    txtinserta = txtinserta & "'5')"                                    'CLAVE POSICION
    ConAdo.Execute txtinserta
    MensajeProc = "Guardando la posicion de fwds en la tabla para carlos frutos"
    DoEvents
 Next i

End Sub

Sub ImpFwdExcel(ByVal fecha As Date, ByVal nomdir As String, ByRef noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim i As Long
Dim j As Long
Dim nocampos As Long
Dim nomarch As String
Dim sihayarch As Boolean
Dim matpos() As New propPosFwd
Dim matpos1() As New propPosFwd
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
noreg = 0
 nomarch = nomdir & "\FWD TC " & Format(fecha, "yyyymmdd") & ".xls"
 sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
   Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
   Set registros1 = base1.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)
   registros1.MoveLast
   noreg = registros1.RecordCount
   registros1.MoveFirst
   ReDim matpos(1 To noreg) As New propPosFwd
   For i = 1 To noreg
       matpos(i).c_operacion = LeerTAccess(registros1, j - 1, i)
       ReDim matpos1(1 To 1) As New propPosFwd
       Set matpos1(1) = matpos(1)
       Call GuardaPosFwds(fecha, 1, "Real", "000000", matpos, exito)
       registros1.MoveNext
       Call MostrarMensajeSistema("Leyendo la posición de fwds del " & Format(fecha, "dd-mm-yyyy") & " ", frmProgreso.Label2, 0, Date, Time, NomUsuario)
   Next i
   registros1.Close
   base1.Close
End If
Exit Sub
ControlErrores:
MsgBox "ImpFwdExcel: " & error(Err())
On Error GoTo 0
End Sub

Function LeerFwdSimArch(ByVal fecha As Date, ByVal nomarch As String, ByRef noreg As Long, ByRef exito As Boolean)
Dim sihayarch As Boolean
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim i As Long
Dim j As Long
Dim nocampos As Long


noreg = 0
exito = False
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then

   Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
   Set registros1 = base1.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)
   registros1.MoveLast
   noreg = registros1.RecordCount
   registros1.MoveFirst
   nocampos = registros1.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   For i = 1 To noreg
       For j = 1 To nocampos
           mata(i, j) = LeerTAccess(registros1, j - 1, i)
       Next j
       registros1.MoveNext
       MensajeProc = "Leyendo la posición de fwds del " & Format(fecha, "dd-mm-yyyy")
   Next i
   registros1.Close
   base1.Close
   exito = True
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
   exito = False
End If
LeerFwdSimArch = mata
End Function

Sub ImpFwdSimArch(ByVal fecha As Date, ByVal nomarch As String, ByVal txtnompos As String, ByRef noreg As Long, ByRef exito As Boolean)
Dim sihayarch As Boolean
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim nocampos As Integer
Dim i As Long
Dim j As Integer
Dim nr As Long
Dim noreg0 As Long
Dim horareg As String
Dim matpos() As New propPosFwd
Dim matpos1() As New propPosFwd


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
noreg = 0
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
   Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
   Set registros1 = base1.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)
   registros1.MoveLast
   noreg = registros1.RecordCount
   registros1.MoveFirst
   ReDim matpos(1 To noreg) As New propPosFwd
   For i = 1 To noreg
       matpos(i).c_operacion = LeerTAccess(registros1, 0, i)
       matpos(i).MontoNocFwd = LeerTAccess(registros1, 2, i)
       matpos(i).FCompraFwd = LeerTAccess(registros1, 3, i)
       matpos(i).FVencFwd = LeerTAccess(registros1, 4, i)
       matpos(i).PAsignadoFwd = LeerTAccess(registros1, 5, i)
       matpos(i).ClaveProdFwd = LeerTAccess(registros1, 6, i)
       matpos(i).intencion = LeerTAccess(registros1, 7, i)
       matpos(i).ID_ContrapFwd = LeerTAccess(registros1, 8, i)
       matpos(i).C_Posicion = LeerTAccess(registros1, 9, i)
       matpos(i).ReclasificaFwd = LeerTAccess(registros1, 11, i)
       matpos(i).EstructuralFwd = LeerTAccess(registros1, 12, i)
       If LeerTAccess(registros1, 1, i) = "Largo" Then
          matpos(i).Tipo_Mov = 1
       Else
          matpos(i).Tipo_Mov = 4
       End If
       registros1.MoveNext

       ReDim matpos1(1 To 1) As New propPosFwd
       Set matpos1(1) = matpos(i)
       horareg = "000000"
       Call GuardaPosFwds(fecha, 2, txtnompos, horareg, matpos1, exito)
       MensajeProc = "Leyendo la posición de fwds del " & Format(fecha, "dd-mm-yyyy")
   Next i
   registros1.Close
   base1.Close
End If
Exit Sub
ControlErrores:
MsgBox "ImpFwdExcel: " & error(Err())
On Error GoTo 0
End Sub

Sub GuardaPosFwds(ByVal fecha As Date, ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horareg As String, ByRef matpos() As propPosFwd, ByRef exito As Boolean)
Dim txtcadena As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtborra As String
Dim contar As Long
Dim i As Long
Dim noreg As Long

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
noreg = UBound(matpos, 1)
For i = 1 To noreg
    txtborra = "DELETE FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfecha & " and TIPOPOS = " & tipopos
    txtborra = txtborra & " AND NOMPOS  = '" & txtnompos & "'"
    txtborra = txtborra & " AND HORAREG  = '" & horareg & "'"
    txtborra = txtborra & " AND CPOSICION = " & matpos(i).C_Posicion
    txtborra = txtborra & " AND COPERACION = '" & matpos(i).c_operacion & "'"
    ConAdo.Execute txtborra
    txtcadena = "INSERT INTO " & TablaPosFwd & " VALUES("
    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha1 = "to_date('" & Format(matpos(i).FCompraFwd, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(matpos(i).FVencFwd, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha3 = "to_date('" & Format(matpos(i).FVencFwd, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = txtcadena & tipopos & ","                              'TIPO DE POSICION: REAL SIM O INTRADIA
    txtcadena = txtcadena & txtfecha & ","                             'fecha de la posicion
    txtcadena = txtcadena & "'" & txtnompos & "',"                     'NOMBRE DE LA posicion simulada
    txtcadena = txtcadena & "'" & horareg & "',"                       'hora de registro
    txtcadena = txtcadena & "'" & matpos(i).intencion & "',"           'intencion
    txtcadena = txtcadena & "'" & matpos(i).ReclasificaFwd & "',"      'reclasificacion
    txtcadena = txtcadena & "'" & matpos(i).EstructuralFwd & "',"      'estructural
    txtcadena = txtcadena & "'" & matpos(i).C_Posicion & "',"          'clave de posicion
    txtcadena = txtcadena & "'" & matpos(i).c_operacion & "',"         'clave de operacion ikos
    txtcadena = txtcadena & matpos(i).Tipo_Mov & ","                   'largo o corto
    txtcadena = txtcadena & matpos(i).MontoNocFwd & ","                'monto nocional
    txtcadena = txtcadena & txtfecha1 & ","                            'fecha inicio
    txtcadena = txtcadena & txtfecha2 & ","                            'fecha de vencimiento
    txtcadena = txtcadena & txtfecha3 & ","                            'fecha de liquidacion
    txtcadena = txtcadena & matpos(i).PAsignadoFwd & ","               'precio pactado
    txtcadena = txtcadena & "'" & matpos(i).ClaveProdFwd & "',"        'tipo de forward
    txtcadena = txtcadena & matpos(i).ID_ContrapFwd & ")"              'contraparte
    ConAdo.Execute txtcadena
    AvanceProc = i / UBound(matpos, 1)
    MensajeProc = "Guardando la posición de fwds del " & Format(fecha, "dd-mm-yyyy") & " " & Format(AvanceProc, "##0.00 %")
    DoEvents
    contar = contar + 1
Next i
If contar = noreg Then
   exito = True
Else
   exito = False
End If
End Sub

Sub ImpFwdID(ByVal fecha As Date, ByVal siarch As Boolean, ByVal txtnompos As String, ByRef noregn As Integer, ByRef noregc As Integer)
Dim mata() As Variant
Dim noreg As Long
Dim nomarch As String
Dim exito As Boolean
Dim contar As Integer
Dim i As Integer
Dim horareg As String
Dim matpos() As New propPosFwd
Dim matpos1() As New propPosFwd

If siarch Then
   frmCalVar.CommonDialog1.ShowOpen
   nomarch = frmCalVar.CommonDialog1.FileName
   mata = LeerFwdSimArch(fecha, nomarch, noreg, exito)
Else
   Call IniciarConexOracle(conAdo2, BDIKOS)
   mata = ImpFwdSimIkos(fecha, noreg)
   conAdo2.Close
End If

noreg = UBound(mata, 1)
contar = 0
noregn = 0
noregc = 0
If noreg <> 0 Then
   ReDim matpos(1 To noreg) As New propPosFwd
   For i = 1 To noreg
       horareg = mata(i, 16)                                        'hora de registro
       matpos(i).C_Posicion = ClavePosDeriv                         'clave de la posicion
       If mata(i, 12) = "N" Or mata(i, 12) = "R" Then
          matpos(i).intencion = "N"                                 'intencion
       Else
          matpos(i).intencion = "C"                                 'intencion
       End If
       matpos(i).c_operacion = mata(i, 2)                           'clave ikos
       If mata(i, 3) = "C" Then                                     'tipo de operacion
          matpos(i).Tipo_Mov = 1
       ElseIf mata(i, 3) = "V" Then
          matpos(i).Tipo_Mov = 4
       Else
       MsgBox "No es una clave permitida en la tabla " & mata(i, 3)
          matpos(i).Tipo_Mov = 1
       End If
       matpos(i).MontoNocFwd = mata(i, 5)                           'monto nocional
       matpos(i).FCompraFwd = ConvertirTextoFecha(mata(i, 6), 0)    'fecha inicio
       matpos(i).FVencFwd = ConvertirTextoFecha(mata(i, 7), 0)      'fecha vencimiento
       matpos(i).fechareg = fecha                                   'fecha de registro
       matpos(i).PAsignadoFwd = Val(mata(i, 13))                    'precio pactado
       matpos(i).ClaveProdFwd = mata(i, 9)                          'clave de producto
   
       If mata(i, 12) = "R" Then
          matpos(i).ReclasificaFwd = "S"                            'reclasificacion
       Else
          matpos(i).ReclasificaFwd = "N"                            'reclasificacion
       End If
       
       matpos(i).EstructuralFwd = mata(i, 14)                       'estructural
       If matpos(i).intencion = "N" Then
          noregn = noregn + 1
       Else
          noregc = noregc + 1
       End If
       matpos(i).ID_ContrapFwd = 0
       ReDim matpos1(1 To 1) As New propPosFwd
       Set matpos1(1) = matpos(i)
       Call GuardaPosFwds(fecha, 3, txtnompos, horareg, matpos1, exito)
   Next i
DoEvents
End If
Exit Sub
End Sub

Sub ActCurvasRE(ByRef mcurvas() As Variant, ByVal txttabla As String, ByRef conex As ADODB.Connection, rbase)
On Error GoTo hayerror
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Long
Dim j As Long
Dim mata() As String
Dim nocampos As Integer
'se procede a actualizar el VaR
noreg = UBound(mcurvas, 1)
txtfiltro = "SELECT * FROM [" & txttabla & "]"
rbase.Open txtfiltro, conex
nocampos = rbase.Fields.Count()
ReDim mata(1 To nocampos - 1) As String
For i = 1 To nocampos - 1
    mata(i) = rbase.Fields(i).Name
Next i
rbase.Close
If noreg > 0 Then
'se verifica si el registro se encuentra en la hoja de excel
   For i = 1 To noreg
       txtcadena = "UPDATE [" & txttabla & "] SET "
       For j = 1 To nocampos - 2
           txtcadena = txtcadena & "[" & mata(j) & "] = " & mcurvas(i, j) & ","
       Next j
       txtcadena = txtcadena & "[" & mata(nocampos - 1) & "] = " & mcurvas(i, nocampos - 1) & " "
       txtcadena = txtcadena & "WHERE [Plazo] = " & i
       conex.Execute txtcadena
       AvanceProc = i / noreg
       MensajeProc = "Avance del proceso " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
End If
Exit Sub
hayerror:
MsgBox error(Err())
End Sub

Sub ActVaRRE(ByVal fecha As Date, ByRef matb() As Double, ByVal nomtabla As String, ByVal txtport As String, ByRef conex As ADODB.Connection, rbase)
On Error GoTo hayerror
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Integer
Dim noreg1 As Integer
'se procede a actualizar el VaR
noreg = UBound(matb, 1)
If noreg > 0 Then
'se verifica si el registro se encuentra en la hoja de excel
'si se encuentra se actualiza, si no se crea
   txtfiltro = "SELECT COUNT(*) FROM [" & nomtabla & "] WHERE FECHA =  " & CLng(fecha)
   rbase.Open txtfiltro, conex
   noreg1 = rbase.Fields(0)
   rbase.Close
If noreg1 <> 0 Then
   'txtcadena = "UPDATE [" & nomtabla & "] SET "
   'txtcadena = txtcadena & "CAMPO1 = " & matb(i, 2) & " "
   'txtcadena = txtcadena & "F3 = " & matb(i, 3) & ","
   'txtcadena = txtcadena & "F4 = " & matb(i, 4) & ","
   'txtcadena = txtcadena & "F5 = " & matb(i, 5) & ","
   'txtcadena = txtcadena & "F6 = " & matb(i, 6) & ","
   'txtcadena = txtcadena & "F7 = '" & matb(i, 7) & "' "
   'txtcadena = txtcadena & "WHERE FECHA = " & CLng(matb(i, 1))
   'conex.Execute txtcadena
Else
   txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
   txtcadena = txtcadena & CLng(fecha) & ","
   txtcadena = txtcadena & matb(1) & ","
   txtcadena = txtcadena & matb(2) & ","
   txtcadena = txtcadena & matb(3) & ","
   txtcadena = txtcadena & matb(4) & ","
   txtcadena = txtcadena & matb(5) & ","
   txtcadena = txtcadena & matb(6) & ")"
   conex.Execute txtcadena
End If
MensajeProc = "Se actualiza el VaR del portafolio de " & txtport & " de " & fecha
End If
Exit Sub
hayerror:
MsgBox error(Err())
End Sub

Sub ActEstressRE(ByRef matb() As Variant, ByVal txttabla As String, ByVal txtport As String, conex, rbase, orden)
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim txtfiltro As String
Dim noreg1 As Long
Dim txtcadena As String
Dim txtclave As String

'actualizando laqs pruebas de estres
noreg = UBound(matb, 1)
If noreg <> 0 Then

   For i = 1 To UBound(matb, 1)
       txtfiltro = "SELECT COUNT(*) FROM [" & txttabla & "] WHERE FECHA = " & CLng(matb(i, 1)) & " AND [mesa escenario] = '" & txtport & "'"
       rbase.Open txtfiltro, conex
       noreg1 = rbase.Fields(0)
       rbase.Close
       If noreg1 <> 0 Then
       
       Else
          txtcadena = "INSERT INTO [" & txttabla & "] VALUES("
          txtclave = CLng(matb(i, 1)) & txtport
          txtcadena = txtcadena & "'" & txtclave & "',"
          txtcadena = txtcadena & CLng(matb(i, 1)) & ","
          txtcadena = txtcadena & i & ","
          txtcadena = txtcadena & "'" & txtport & "',"
          For j = 1 To 11
              txtcadena = txtcadena & Val(matb(i, j + 7)) & ","
          Next j
          txtcadena = txtcadena & Val(matb(i, 19)) & ")"
          conex.Execute txtcadena
       End If
   Next i
End If

End Sub

Function LeerHistBack(ByVal fecha As Date) As Double()
Dim noreg As Integer
Dim i As Integer
Dim mata() As Double

noreg = UBound(MatPortSegRiesgo, 1)
ReDim matv(1 To noreg) As Double
For i = 1 To noreg
    mata = LeerBackPort(fecha, txtportCalc2, MatPortSegRiesgo(i, 1))
    matv(i) = mata(3)
Next i
   LeerHistBack = matv
End Function

Function LeerBackPort(ByVal fecha As Date, ByVal txtport As String, ByVal txtsubport As String)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

'una vez abierta la tabla de datos se procede a buscar la informacion
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaBackPort & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & txtsubport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To 3) As Double
   mata(1) = rmesa.Fields("valor1")
   mata(2) = rmesa.Fields("valor2")
   mata(3) = rmesa.Fields("diferencia")
   rmesa.Close
Else
   ReDim mata(1 To 3) As Double
End If
LeerBackPort = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LecturaResVaRO(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal dvar As Integer, ByVal horiz As Integer, ByVal nconf As Double) As Variant()
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro As String
Dim i As Long
Dim ll As Long
Dim txtcadena As String
Dim txtclave1 As String
Dim rmesa As New ADODB.recordset


'esta rutina es para obtener los resultados de var para un tamaño de muestra de
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "select count(distinct FECHA) FROM " & TablaResVaR & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <=" & txtfecha2 & " AND HORIZ = " & horiz & " AND NIV_CONF = " & nconf & " ORDER BY FECHA"
rmesa.Open txtfiltro, ConAdo
NoFechas = rmesa.Fields(0)
rmesa.Close
'primero se obtienen las fechas
If NoFechas <> 0 Then
ReDim mata(1 To NoFechas, 1 To NoPortafolios + 1) As Variant
txtfiltro = "select FECHA FROM " & TablaResVaR & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <=" & txtfecha2 & " AND HORIZ = " & horiz & " AND NIV_CONF = " & nconf & "  GROUP BY FECHA ORDER BY FECHA"
rmesa.Open txtfiltro, ConAdo
For i = 1 To NoFechas
mata(i, 1) = rmesa.Fields(0)
rmesa.MoveNext
Next i
rmesa.Close

For i = 1 To NoFechas
For ll = 1 To NoPortafolios
txtfecha = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
'VAR CONSOLIDADO
txtcadena = "FECHA = " & txtfecha & " AND MPOSICION = '" & txtportBanobras & "' AND PORTAFOLIO = '" & MatPortafolios(ll, 2) & "' AND TVAR = 'CVAR VOL CONST P' AND DIAS_VOL = " & dvar & " AND HORIZ = " & horiz & " AND NIV_CONF = " & nconf
txtfiltro = "SELECT * FROM " & TablaResVaR & " WHERE " & txtcadena
rmesa.Open txtfiltro, ConAdo
If Not rmesa.EOF() Then
   mata(i, ll + 1) = Truncar(Val(rmesa.Fields("LIM_INF_VAR")), 4)
   If Val(mata(i, ll + 1)) = 0 And txtclave1 <> "3" Then MsgBox "No hay var histórico vol const " & txtclave1 & " " & MatPortafolios(ll, 2)
Else
   If txtclave1 <> "3" Then Call MostrarMensajeSistema("No hay var histórico vol const " & txtclave1 & " " & MatPortafolios(ll, 2) & " " & mata(i, 1), frmProgreso.Label2, 1, Date, Time, NomUsuario)
End If
rmesa.Close
Next ll
Next i

Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
LecturaResVaRO = mata
End Function

Function LeerCVaRPortPos(ByVal fecha As Date, ByVal txtport As String, ByVal noesc As Integer, htiempo As Integer, ByVal nosim As Integer, ByVal nconf As Double, ByVal lambda As Double, ByVal txttvar As String)
'esta rutina es para obtener los resultados de var para un tamaño de muestra de
Dim txtportfr As String
Dim exito As Boolean
Dim i As Integer
Dim noreg As Long


txtportfr = "Normal"
   noreg = UBound(MatPortSegRiesgo, 1)
   ReDim mata(1 To noreg) As Double
   For i = 1 To noreg
   mata(i) = LeerResVaR(fecha, txtport, txtportfr, MatPortSegRiesgo(i, 1), noesc, htiempo, nosim, nconf, lambda, txttvar, exito)
   Next i
LeerCVaRPortPos = mata
End Function

Function LeerResCVaRPrev(ByVal fecha As Date, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtport As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset
noreg1 = UBound(MatPortSegRiesgo, 1)

txtport = txtportCalc1
txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
ReDim mata(1 To noreg1) As Double
For i = 1 To noreg1
    txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & MatPortSegRiesgo(i, 1) & "'"
    txtfiltro2 = txtfiltro2 & " AND TVAR = 'CVARPrev'"
    txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc
    txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo
    txtfiltro2 = txtfiltro2 & " AND NCONF = 0.01"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       mata(i) = rmesa.Fields("VALOR")
       rmesa.Close
    Else
    mata(i) = 0
    End If
   
Next i
LeerResCVaRPrev = mata
End Function

Function LeerResCVaRExp(ByVal fecha As Date, ByVal txtgrupoport As String, ByVal noesc As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtport As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim mats() As String
Dim rmesa As New ADODB.recordset

noreg1 = UBound(MatPortSegRiesgo, 1)
txtport = txtportCalc1
txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
ReDim mata(1 To noreg1) As Double
For i = 1 To noreg1
    txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & MatPortSegRiesgo(i, 1) & "'"
    txtfiltro2 = txtfiltro2 & " AND TVAR = 'VARExp'"
    txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc
    txtfiltro2 = txtfiltro2 & " AND HTIEMPO = 1"
    txtfiltro2 = txtfiltro2 & " AND LAMBDA = 0.97"
    txtfiltro2 = txtfiltro2 & " AND NCONF = 0.01"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       mata(i) = rmesa.Fields("VALOR")
       rmesa.Close
    Else
       mata(i) = 0
    End If
Next i
LeerResCVaRExp = mata
End Function

Sub GuardarResVaRPrev(ByVal fecha As Date, ByRef mata() As Double, ByVal txttabla As String, conex, rbase)
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim j As Integer

noreg = UBound(mata, 1)
If noreg > 0 Then
       txtfiltro = "SELECT COUNT(*) FROM [" & txttabla & "] WHERE FECHA = " & CLng(fecha)
       rbase.Open txtfiltro, conex
       noreg1 = rbase.Fields(0)
       rbase.Close
       If noreg1 <> 0 Then
       
       Else
          txtcadena = "INSERT INTO [" & txttabla & "] VALUES("
          txtcadena = txtcadena & CLng(fecha) & ","
          For j = 1 To noreg - 1
              txtcadena = txtcadena & Val(mata(j)) & ","
          Next j
          txtcadena = txtcadena & Val(mata(noreg)) & ")"
          conex.Execute txtcadena
       End If
End If
End Sub

Sub GuardarResVaRExp(ByVal fecha As Date, ByRef mata() As Double, ByVal txttabla As String, conex, rbase)
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim j As Integer

noreg = UBound(mata, 1)
If noreg > 0 Then
       txtfiltro = "SELECT COUNT(*) FROM [" & txttabla & "] WHERE FECHA = " & CLng(fecha)
       rbase.Open txtfiltro, conex
       noreg1 = rbase.Fields(0)
       rbase.Close
       If noreg1 <> 0 Then
       
       Else
          txtcadena = "INSERT INTO [" & txttabla & "] VALUES("
          txtcadena = txtcadena & CLng(fecha) & ","
          For j = 1 To noreg - 1
              txtcadena = txtcadena & Val(mata(j)) & ","
          Next j
          txtcadena = txtcadena & Val(mata(noreg)) & ")"
          conex.Execute txtcadena
       End If
End If
End Sub


Sub GuardarResVaRAcum(ByVal fecha As Date, ByRef mata() As Double, ByVal txttabla As String, conex, rbase)
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim j As Integer

noreg = UBound(mata, 1)
If noreg > 0 Then
       txtfiltro = "SELECT COUNT(*) FROM [" & txttabla & "] WHERE FECHA = " & CLng(fecha)
       rbase.Open txtfiltro, conex
       noreg1 = rbase.Fields(0)
       rbase.Close
       If noreg1 <> 0 Then
       
       Else
          txtcadena = "INSERT INTO [" & txttabla & "] VALUES("
          txtcadena = txtcadena & CLng(fecha) & ","
          For j = 1 To noreg - 1
              txtcadena = txtcadena & Val(mata(j)) & ","
          Next j
          txtcadena = txtcadena & Val(mata(noreg)) & ")"
          conex.Execute txtcadena
       End If
End If
End Sub

Function LeerResVarPort(ByVal fecha As Date, ByVal txtportfr As String, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Long)
Dim exito As Boolean
Dim mata(1 To 6) As Double
mata(1) = LeerResVaR(fecha, txtportCalc1, txtportfr, txtsubport, noesc, htiempo, 0, 0.03, 0, "CVARH", exito)
mata(2) = LeerResVaR(fecha, txtportCalc1, txtportfr, txtsubport, noesc, htiempo, 0, 0.97, 0, "CVARH", exito)
mata(3) = LeerResVaR(fecha, txtportCalc2, txtportfr, txtsubport, noesc, htiempo, 0, 0.01, 0, "VARMark", exito)
mata(4) = LeerResVaR(fecha, txtportCalc2, txtportfr, txtsubport, noesc, htiempo, 0, 0.99, 0, "VARMark", exito)
mata(5) = LeerResVaR(fecha, txtportCalc2, txtportfr, txtsubport, noesc, htiempo, nosim, 0.01, 0, "VARMont", exito)
mata(6) = LeerResVaR(fecha, txtportCalc2, txtportfr, txtsubport, noesc, htiempo, nosim, 0.99, 0, "VARMont", exito)
LeerResVarPort = mata
End Function

Function LeerResVaR(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Integer, ByVal nconf As Double, ByVal lambda As Double, ByVal txttvar As String, ByRef exito As Boolean) As Double
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim valor As Double
Dim noreg As Long
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & txtsubport & "'"
txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo
txtfiltro2 = txtfiltro2 & " AND NOSIM = " & nosim
txtfiltro2 = txtfiltro2 & " AND NCONF = " & nconf
txtfiltro2 = txtfiltro2 & " AND LAMBDA = " & lambda
txtfiltro2 = txtfiltro2 & " AND TVAR = '" & txttvar & "'"
txtfiltro1 = "SELECT COUNT (*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   valor = rmesa.Fields("VALOR")
   rmesa.Close
   LeerResVaR = valor
   exito = True
Else
   LeerResVaR = 0
   exito = False
End If

End Function


Function CargaPortPos()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Dim txtfiltro1 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

'====================================================
txtfiltro1 = "select count(*) from " & PrefijoBD & TablaCatPortPos
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 txtfiltro1 = "select * from " & PrefijoBD & TablaCatPortPos & " order BY ID_PORT"
 rmesa.Open txtfiltro1, ConAdo
 ReDim mata(1 To noreg, 1 To 3) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
  mata(i, 1) = rmesa.Fields("ID_PORT")
  mata(i, 2) = rmesa.Fields("NOMBRE")
  mata(i, 3) = rmesa.Fields("DESCRIPCION")
  rmesa.MoveNext
  AvanceProc = i / noreg
  MensajeProc = "Cargando los portafolios de posiciones " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
 End If
CargaPortPos = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CargaPortCurvas()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se cargan los distinto portafolios de curvas
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset
'====================================================
txtfiltro1 = "select count(DISTINCT PORTAFOLIO) from " & PrefijoBD & TablaPortFR
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Variant
txtfiltro1 = "select PORTAFOLIO from " & PrefijoBD & TablaPortFR & " GROUP BY PORTAFOLIO"
rmesa.Open txtfiltro1, ConAdo
ReDim mata(1 To noreg, 1 To 2) As Variant
rmesa.MoveFirst
For i = 1 To noreg
 mata(i, 1) = rmesa.Fields("PORTAFOLIO")
 rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Cargando los portafolios de factores de riesgo" & Format(AvanceProc, "###0.00 %")
Next i
rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaPortCurvas = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CargaRelSwapPrim()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se cargan el catalogo que determina como se va a definir la clave de la emision
Dim txtfiltro1 As String, txtfiltro2 As String
Dim noreg As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim rmesa As New ADODB.recordset
Dim mata() As New propRelSwapPrim
'====================================================
txtfiltro1 = "select count(*) from " & PrefijoBD & TablaPosPrimarias
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtfiltro1 = "select * from " & PrefijoBD & TablaPosPrimarias
   rmesa.Open txtfiltro1, ConAdo
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To noreg)
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i).coperacion = rmesa.Fields("COPERACION")
       mata(i).finicio = rmesa.Fields("FINICIO")
       mata(i).ffin = rmesa.Fields("FFINAL")
       mata(i).c_ppactiva = ReemplazaVacioValor(rmesa.Fields("PACTIVA"), "")
       mata(i).c_pppasiva = ReemplazaVacioValor(rmesa.Fields("PPASIVA"), "")
       mata(i).c_pswap = ReemplazaVacioValor(rmesa.Fields("PSWAP"), "")
       mata(i).t_efect = rmesa.Fields("ID_EFECTIVIDAD")
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Cargando la relacion de posiciones primarias " & Format(AvanceProc, "##0.00 %")
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0)
End If
CargaRelSwapPrim = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerPortPosEstruc() As String()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim mata() As String
Dim rmesa As New ADODB.recordset
'====================================================
txtfiltro2 = "select PORTAFOLIO FROM " & PrefijoBD & TablaPortPosEstructural & " GROUP BY PORTAFOLIO ORDER BY PORTAFOLIO"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mata(1 To noreg)
'se borran los registros que pueda haber con esta fecha
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       mata(i) = rmesa.Fields(0)
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Cargando la relacion de posiciones primarias " & Format(AvanceProc, "##0.00 %")
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0)
End If
LeerPortPosEstruc = mata
End Function

Function CargaClaveEmision()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se cargan el catalogo que determina como se va a definir la clave de la emision

Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim nocampos As Integer
Dim rmesa As New ADODB.recordset
'====================================================
txtfiltro1 = "select count(*) from " & PrefijoBD & TablaCEmision
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Variant
'se borran los registros que pueda haber con esta fecha
txtfiltro1 = "select * from " & PrefijoBD & TablaCEmision
rmesa.Open txtfiltro1, ConAdo
nocampos = rmesa.Fields.Count
ReDim mata(1 To noreg, 1 To nocampos) As Variant
rmesa.MoveFirst
For i = 1 To noreg
 For j = 1 To nocampos
 mata(i, j) = rmesa.Fields(j - 1)
 Next j
 rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Cargando el catalogos de claves de emisión " & Format(AvanceProc, "##0.00 %")
Next i
rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaClaveEmision = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerSensibNuevo(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String) As Variant()
'rutina para exportar las sensibilidades mas importantes
'a excel, desde oracle
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha1 As String
Dim txtdesc As String
Dim i As Integer
Dim noreg As Long
Dim rmesa As New ADODB.recordset

'====================================================
txtfecha1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "select * from " & TablaSensibPort & " WHERE FECHA = " & txtfecha1 & " AND PORTAFOLIO ='" & txtport & "' "
txtfiltro2 = txtfiltro2 & "AND PORT_FR = '" & txtportfr & "' AND SUBPORT = '" & txtsubport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
RnMesa.Open txtfiltro1, ConAdo
noreg = RnMesa.Fields(0)
RnMesa.Close
If noreg <> 0 Then
   ReDim mata(1 To noreg, 1 To 11) As Variant
'se borran los registros que pueda haber con esta fecha
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       txtdesc = rmesa.Fields("Descripcion")
       txtdesc = ReemplazaCadenaTexto(ReemplazaCadenaTexto(ReemplazaCadenaTexto(txtdesc, "dma", "día"), "dslar", "dólar"), "dia", "día")
       mata(i, 1) = ReemplazaCadenaTexto(ReemplazaCadenaTexto(ReemplazaCadenaTexto(txtdesc, "d?a", "día"), "d?lar", "dólar"), "dia", "día")
       mata(i, 2) = rmesa.Fields("FACTOR")                       'FACTOR
       mata(i, 3) = rmesa.Fields("CURVA")                        'curva a la que pertenece el factor
       mata(i, 4) = rmesa.Fields("PLAZO")                        'plazo del factor
       mata(i, 5) = rmesa.Fields("TVALOR")                       'tipo de factor
       mata(i, 6) = rmesa.Fields("VALOR")                        'valor del factor
       mata(i, 7) = rmesa.Fields("DERIVADA")                     'derivada
       mata(i, 8) = rmesa.Fields("VOLATIL")                      'volatilidad porcentual
       mata(i, 9) = Abs(mata(i, 6) * mata(i, 7) * mata(i, 8))    'var
       mata(i, 10) = mata(i, 7) / 1000000 / 100                  'derivada ajustada
       If mata(i, 5) = "TASA" Or mata(i, 5) = "YIELD" Or mata(i, 5) = "TASA EXT" Or mata(i, 5) = "TASA REAL" Or mata(i, 5) = "SOBRETASA" Or mata(i, 5) = "TASA REF" Or mata(i, 5) = "YIELD IS" Or mata(i, 5) = "TASA REF EXT" Then
          mata(i, 11) = mata(i, 8) * mata(i, 6) * 100 * 100      'VOLATILIDAD AJUSTADA
       ElseIf mata(i, 5) = "T CAMBIO" Or mata(i, 5) = "UDI" Or mata(i, 5) = "INDICE" Then
           mata(i, 11) = mata(i, 8) * mata(i, 6) * 100
       Else
          MsgBox "no se clasifico " & mata(i, 2)
       End If
       rmesa.MoveNext
   Next i
   rmesa.Close
   mata = RutinaOrden(mata, 9, SRutOrden)
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerSensibNuevo = mata
MensajeProc = "Se leyeron " & noreg & " registros de la tabla de sensibilidades para el dia " & fecha
End Function

Sub PegarSensResEx(ByVal fecha As Date, ByRef matb() As Variant, ByVal txtport As String, ByVal txtbase As String, conex, rbase)
Dim noreg As Integer
Dim i As Integer
Dim txtclave As String
Dim txtcadena As String

'se guardan las sensibilidades mas importantes en
'la hoja de sensibilidades
noreg = UBound(matb, 1)
If noreg <> 0 Then
 For i = 1 To Minimo(noreg, 20)
  txtclave = Format(i, "000") & CLng(fecha) & txtport
  txtcadena = "INSERT INTO [" & txtbase & "] VALUES("
  txtcadena = txtcadena & "'" & txtclave & "',"                        'clave de identificacion
  txtcadena = txtcadena & "'" & txtport & "',"                         'nombre del portafolio
  txtcadena = txtcadena & CLng(fecha) & ","                            'fecha
  txtcadena = txtcadena & "'" & matb(noreg - i + 1, 1) & "',"          'nombre del factor
  txtcadena = txtcadena & "'" & matb(noreg - i + 1, 3) & "',"          'curva
  txtcadena = txtcadena & matb(noreg - i + 1, 4) & ","                 'plazo
  txtcadena = txtcadena & matb(noreg - i + 1, 6) & ","                 'valor
  txtcadena = txtcadena & matb(noreg - i + 1, 11) & ","                'volatilidad absoluta
  txtcadena = txtcadena & matb(noreg - i + 1, 10) & ")"                'derivada
  conex.Execute txtcadena
  AvanceProc = i / noreg
  MensajeProc = "Exp sensib pos " & txtport & " de " & fecha & " " & Format(AvanceProc, "##0.00 %")
 Next i
End If
End Sub


Function CargaContrapartes()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

 txtfiltro = "select * from " & PrefijoBD & TablaContrapartes & " ORDER BY ID_CONTRAP"
 txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
 rmesa.Open txtfiltro1, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
 If noreg <> 0 Then
 rmesa.Open txtfiltro, ConAdo
 ReDim mata(1 To noreg, 1 To 9) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("ID_CONTRAP")
     mata(i, 2) = rmesa.Fields("NOMBRE")
     mata(i, 3) = rmesa.Fields("NCORTO")
     mata(i, 4) = rmesa.Fields("NOMBREIKOS")
     mata(i, 5) = rmesa.Fields("NCORTOIKOS")
     mata(i, 6) = rmesa.Fields("SECTOR")
     mata(i, 7) = rmesa.Fields("SECTOR2")
     mata(i, 8) = rmesa.Fields("USUARIO")
     mata(i, 9) = rmesa.Fields("FECHA")
     rmesa.MoveNext
     AvanceProc = i / noreg
     MensajeProc = "Cargando las contrapartes " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaContrapartes = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function


Function CargaTresholdContrap()
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

 txtfiltro = "select * from " & PrefijoBD & TablaTreshCont & " ORDER BY ID_CONTRAP, FECHA"
 txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
 rmesa.Open txtfiltro1, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
 If noreg <> 0 Then
    rmesa.Open txtfiltro, ConAdo
 ReDim mata(1 To noreg, 1 To 9) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("ID_CONTRAP")
     mata(i, 2) = rmesa.Fields("FECHA")
     mata(i, 3) = rmesa.Fields("TRESHOLD_C")
     mata(i, 4) = rmesa.Fields("TRESHOLD_F")
     mata(i, 5) = rmesa.Fields("MMTRANSFER")
     mata(i, 6) = rmesa.Fields("MONEDA")
  rmesa.MoveNext
  AvanceProc = i / noreg
  MensajeProc = "Cargando las contrapartes " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaTresholdContrap = mata
End Function

Function CargaClavesContrap()
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

 txtfiltro = "select * from " & PrefijoBD & TablaEquivContrap & " ORDER BY FECHA,ID_CONTRAP"
 txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
 rmesa.Open txtfiltro1, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
 If noreg <> 0 Then
    rmesa.Open txtfiltro, ConAdo
 ReDim mata(1 To noreg, 1 To 9) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("FECHA")        'fecha
     mata(i, 2) = rmesa.Fields("ID_CONTRAP")   'id sivarmer
     If Not EsVariableVacia(rmesa.Fields(2)) Then
     mata(i, 3) = rmesa.Fields(2)              'id ikos
     Else
     mata(i, 3) = ""
     End If
     If Not EsVariableVacia(rmesa.Fields(3)) Then
     mata(i, 4) = rmesa.Fields(3)              'id banxico
     Else
     mata(i, 4) = ""
     End If
  rmesa.MoveNext
  AvanceProc = i / noreg
  MensajeProc = "Cargando las contrapartes " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaClavesContrap = mata
End Function

Function CargaDerivSinLMargen(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfiltro = "select * from " & PrefijoBD & TablaBLTresh & " WHERE FECHA IN(SELECT MAX(FECHA) FROM " & TablaBLTresh & " WHERE FECHA <= " & txtfecha & ")"
 txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
 rmesa.Open txtfiltro1, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
 If noreg <> 0 Then
 rmesa.Open txtfiltro, ConAdo
 ReDim mata(1 To noreg, 1 To 1) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("COPERACION")
     rmesa.MoveNext
     AvanceProc = i / noreg
     MensajeProc = "Cargando las contrapartes " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaDerivSinLMargen = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function CargaRecuperacion(ByVal fecha As Date, ByVal txtbase As String) As Variant()
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "select * from " & txtbase & " where FECHA IN"
txtfiltro2 = txtfiltro2 & "(SELECT MAX(FECHA) FROM " & txtbase & " WHERE FECHA <=" & txtfecha & ") ORDER BY CALIFICACION"
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"
 rmesa.Open txtfiltro1, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
 If noreg <> 0 Then
 rmesa.Open txtfiltro2, ConAdo
 ReDim mata(1 To noreg, 1 To 6) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("CALIFICACION")
     mata(i, 2) = rmesa.Fields("RECUPERA1")
     mata(i, 3) = rmesa.Fields("RECUPERA2")
     mata(i, 4) = rmesa.Fields("RECUPERA3")
     mata(i, 5) = rmesa.Fields("RECUPERA4")
     mata(i, 6) = rmesa.Fields("RECUPERA5")
     rmesa.MoveNext
  AvanceProc = i / noreg
  MensajeProc = "Cargando las contrapartes " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
CargaRecuperacion = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function BuscarFactorFecha(ByVal fecha As Date, ByVal nomfac As String)
Dim i As Long
Dim indice As Long
Dim fechax As Date
Dim ind1 As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'busca y devuelve el valor del factor de riesgo para la fecha
'especificada
For i = 1 To NoFactores
    If MatCaracFRiesgo(i).nomFactor = nomfac Then
       ind1 = i
       Exit For
    End If
Next i
indice = 0
fechax = fecha
Do
indice = BuscarValorArray(fechax, MatFactRiesgo, 1)
fechax = fechax + 1
Loop Until indice <> 0
BuscarFactorFecha = MatFactRiesgo(indice, ind1)
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub GuardaDatosBitacora(ByVal id_t_evento As Integer, ByVal tevento As String, ByVal idproc As Long, ByVal txtproc As String, ByVal usuario As String, ByVal fecha As Date, ByVal txtmsg As String, ByVal opcion As Integer)
Dim finicio As Date
Dim ffinal As Date
Dim hinicio As Date
Dim hfinal As Date
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txthora1 As String
Dim txthora2 As String
Dim txtcadena As String
Dim ipdirecc As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim txttabla As String
Dim rmesa As New ADODB.recordset
If opcion = 1 Then
   txttabla = TablaProcesos1
ElseIf opcion = 2 Then
   txttabla = TablaProcesos2
End If
'el archivo de bitacora puede llegar a crecer mucho ya
'que aqui se registran todas las incidencias en las
'operaciones que el sistema tiene en todo el día
txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
If opcion = 1 Or opcion = 2 Then
txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
If idproc <> 0 Then
   txtfiltro2 = "SELECT * FROM " & txttabla & " WHERE FECHAP = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND ID_TAREA = " & idproc
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0).value
   rmesa.Close
   If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      finicio = rmesa.Fields("FINICIAL").value
      hinicio = rmesa.Fields("HINICIAL").value
      ffinal = rmesa.Fields("FFINAL").value
      hfinal = rmesa.Fields("HFINAL").value
      rmesa.Close
   End If
Else
   finicio = Date
   hinicio = Time
   ffinal = Date
   hfinal = Time
End If
End If
 ipdirecc = RecuperarIP
 txtfecha1 = "TO_DATE('" & Format(finicio, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfecha2 = "TO_DATE('" & Format(ffinal, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txthora1 = "to_date('" & Format$(hinicio, "hh:mm:ss") & "','HH24:MI:SS')"
 txthora2 = "to_date('" & Format$(hfinal, "hh:mm:ss") & "','HH24:MI:SS')"
 
 txtcadena = "INSERT INTO " & TablaBitacora & " VALUES("
 txtcadena = txtcadena & id_t_evento & ","
 txtcadena = txtcadena & "'" & tevento & "',"
 txtcadena = txtcadena & idproc & ","
 txtcadena = txtcadena & "'" & txtproc & "',"
 txtcadena = txtcadena & "'" & usuario & "',"
 txtcadena = txtcadena & "'" & ipdirecc & "',"
 txtcadena = txtcadena & txtfecha & ","
 txtcadena = txtcadena & txtfecha1 & ","
 txtcadena = txtcadena & txthora1 & ","
 txtcadena = txtcadena & txtfecha2 & ","
 txtcadena = txtcadena & txthora2 & ","
 txtcadena = txtcadena & "'" & Left(txtmsg, 300) & "')"
 ConAdo.Execute txtcadena
 txtcadena = "UPDATE " & TablaUsuarios & " SET FUREPORTE= " & txtfecha2 & ", HUREPORTE = " & txthora2 & " WHERE USUARIO = '" & usuario & "'"
 ConAdo.Execute txtcadena

End Sub

Sub GuardaAccesoBitacora(ByVal id_t_evento As Integer, ByVal tevento As String, ByVal idproc As Long, ByVal txtproc As String, ByVal usuario As String, ByVal fecha As Date, ByVal txtmsg As String)
Dim finicio As Date
Dim ffinal As Date
Dim hinicio As Date
Dim hfinal As Date
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txthora1 As String
Dim txthora2 As String
Dim txtcadena As String
Dim ipdirecc As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer


'el archivo de bitacora puede llegar a crecer mucho ya
'que aqui se registran todas las incidencias en las
'operaciones que el sistema tiene en todo el día
 txtfecha = "TO_DATE('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 ipdirecc = RecuperarIP
 txtfecha1 = "TO_DATE('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfecha2 = "TO_DATE('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txthora1 = "to_date('" & Format$(Time, "hh:mm:ss") & "','HH24:MI:SS')"
 txthora2 = "to_date('" & Format$(Time, "hh:mm:ss") & "','HH24:MI:SS')"
 
 txtcadena = "INSERT INTO " & TablaBitacora & " VALUES("
 txtcadena = txtcadena & id_t_evento & ","
 txtcadena = txtcadena & "'" & tevento & "',"
 txtcadena = txtcadena & idproc & ","
 txtcadena = txtcadena & "'" & txtproc & "',"
 txtcadena = txtcadena & "'" & usuario & "',"
 txtcadena = txtcadena & "'" & ipdirecc & "',"
 txtcadena = txtcadena & txtfecha & ","
 txtcadena = txtcadena & txtfecha1 & ","
 txtcadena = txtcadena & txthora1 & ","
 txtcadena = txtcadena & txtfecha2 & ","
 txtcadena = txtcadena & txthora2 & ","
 txtcadena = txtcadena & "'" & Left(txtmsg, 300) & "')"
 ConAdo.Execute txtcadena
 txtcadena = "UPDATE " & TablaUsuarios & " SET FUREPORTE= " & txtfecha2 & ", HUREPORTE = " & txthora2 & " WHERE USUARIO = '" & usuario & "'"
 ConAdo.Execute txtcadena

End Sub

Sub ImportarVPrecios(ByVal fecha As Date, ByVal direc As String, ByRef noreg As Long, ByVal tfecha As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
  'se construye el nombre del archivo
  Dim txtnomarch1 As String
  Dim txtnomarch2 As String
  Dim txtnomarch3 As String
  Dim colorden As Integer
  Dim mfriesgo() As Variant
  Dim matvec3() As Variant
  Dim sihayarch1 As Boolean
  Dim sihayarch2 As Boolean
  Dim sihayarch3 As Boolean
  Dim matvec2() As Variant
  Dim i As Long
  Dim indice As Long
  Dim nreg1 As Long
  Dim matvp() As New propVecPrecios
  
  
  exito = True
  colorden = 32
    noreg = 0
    txtnomarch1 = direc & "\" & "VMD" & Format(fecha, "yyyymmdd") & ".CSV"
    txtnomarch2 = direc & "\" & "PIP" & Format(fecha, "yyyymmdd") & "M.XLS"
    txtnomarch3 = direc & "\" & "VectorAnalitico" & Format(fecha, "yyyymmdd") & "MD.csv"
    sihayarch1 = VerifAccesoArch(txtnomarch1)
    sihayarch2 = VerifAccesoArch(txtnomarch2)
    sihayarch3 = VerifAccesoArch(txtnomarch3)
'  se lee desde un archivo de texto
    If sihayarch1 And sihayarch2 And sihayarch3 Then
       mfriesgo = LeerArchTexto(txtnomarch1, ",", "Leyendo el vector de precios VMD")
       mfriesgo = VerifVPrecios1(mfriesgo, tfecha)
       mfriesgo = AnexarClaveEmision(mfriesgo, 3, 4, 5, colorden)
       mfriesgo = RutinaOrden(mfriesgo, colorden, SRutOrden)
       sihayarch2 = VerifAccesoArch(txtnomarch2)
       If sihayarch2 Then
          matvec2 = LeerVectPreciosMD(txtnomarch2)
       ' SE procede a buscar la yield y a guardar el datos en una sola tabla
          If UBound(matvec2, 1) <> 0 Then
             For i = 1 To UBound(matvec2, 1)
                 If Not EsVariableVacia(matvec2(i, 10)) And matvec2(i, 9) <> 0 Then
                       indice = BuscarValorArray(matvec2(i, 10), mfriesgo, colorden)
                       If indice <> 0 Then
                          mfriesgo(indice, 17) = matvec2(i, 9)   'yield
                       End If
                 End If
             Next i
          End If
       Else
          txtmsg = "No hay acceso al archivo PIP"
          exito = False
       End If
       sihayarch3 = VerifAccesoArch(txtnomarch3)
       If sihayarch3 Then
          matvec3 = LeerVectAnalitico(txtnomarch3)
       'SE procede a buscar la yield y a guardar el datos en una sola tabla
          If UBound(matvec3, 1) <> 0 Then
             For i = 1 To UBound(matvec3, 1)
                 If Not EsVariableVacia(matvec3(i, 1)) Then
                    indice = BuscarValorArray(matvec3(i, 1), mfriesgo, colorden)
                    If indice <> 0 Then
                       mfriesgo(indice, 17) = ConvValor(matvec3(i, 10))   'yield
                       mfriesgo(indice, 18) = matvec3(i, 2)               'sp
                       mfriesgo(indice, 19) = matvec3(i, 3)               'fitch
                       mfriesgo(indice, 20) = matvec3(i, 4)               'moodys
                       mfriesgo(indice, 21) = matvec3(i, 5)               'hr
                       mfriesgo(indice, 22) = matvec3(i, 6)               'verum
                       mfriesgo(indice, 23) = "-"                         'am
                       mfriesgo(indice, 24) = matvec3(i, 7)               'dbrs
                       mfriesgo(indice, 25) = matvec3(i, 8)               'frecuencia cupon
                       mfriesgo(indice, 26) = matvec3(i, 9)               'regla cupon
                       mfriesgo(indice, 27) = ConvValor(matvec3(i, 11))   'sobretasa de colocacion
                       mfriesgo(indice, 28) = ConvValor(matvec3(i, 12))   'monto emitido
                       mfriesgo(indice, 29) = ConvValor(matvec3(i, 13))   'monto en circulacion
                       mfriesgo(indice, 30) = ReemplazaVacioValor(matvec3(i, 14), "")             'sector
                       mfriesgo(indice, 31) = ReemplazaVacioValor(matvec3(i, 15), "")             'isin
                    End If
                 End If
             Next i
          End If
       End If
       Call GuardarVPrecios(fecha, TablaVecPrecios, mfriesgo, nreg1, exito)
       If exito Then
          txtmsg = "El proceso finalizo correctamente"
       End If
       noreg = noreg + nreg1
    Else
       exito = False
       txtmsg = "Falta alguno de los vectores del " & Format(fecha, "dd/mm/yyyy")
    End If
End Sub

Function LeerVectAnalitico(ByVal txtnomarch As String)
Dim matv() As Variant
Dim noreg As Integer
Dim i As Integer

matv = LeerArchTexto(txtnomarch, ",", "Leyendo el vector analítico")
noreg = UBound(matv, 1)
ReDim mata(1 To noreg, 1 To 15) As Variant
For i = 1 To noreg
    mata(i, 1) = GeneraClaveEmision(matv(i, 2), matv(i, 3), matv(i, 4))
    mata(i, 2) = matv(i, 38)    'calificacion sp
    mata(i, 3) = matv(i, 54)    'calificacion fitchs
    mata(i, 4) = matv(i, 37)    'calificacion moodys
    mata(i, 5) = matv(i, 60)    'calificacion hr
    mata(i, 6) = matv(i, 63)    'calificacion verum
    mata(i, 7) = matv(i, 64)    'calificacion dbrs
    mata(i, 8) = matv(i, 22)    'frecuencia cupon
    mata(i, 9) = matv(i, 25)    'regla cupon
    mata(i, 10) = matv(i, 59)   'yield
    mata(i, 11) = matv(i, 21)   'sobretasa de colocacion
    mata(i, 12) = matv(i, 12)   'monto emitido
    mata(i, 13) = matv(i, 13)   'monto en circulacion
    mata(i, 14) = matv(i, 11)   'SECTOR
    mata(i, 15) = matv(i, 62)   'ISIN
Next i
LeerVectAnalitico = mata
End Function

Sub ImpVectorPrecios22(ByVal fecha As Date, ByVal nomarch As String, ByRef nr As Long, ByRef exito As Boolean)
Dim mata() As Variant
Dim contar As Long
Dim i As Long
Dim claveem As String
Dim txtfecha As String
Dim txtcadena As String
Dim noreg As Long

mata = LeerVectPreciosMD(nomarch)
If UBound(mata, 1) > 0 Then
   contar = 0
   noreg = UBound(mata, 1)
   For i = 1 To noreg
       If mata(i, 9) <> 0 Then
          claveem = mata(i, 10)
          txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtcadena = "UPDATE " & TablaVecPrecios & " SET YIELD = " & mata(i, 9) & " WHERE FECHA = " & txtfecha & " AND CLAVE_EMISION = '" & claveem & "'"
          ConAdo.Execute txtcadena
          AvanceProc = i / noreg
          MensajeProc = "Guardando la informacion del segundo vector de precios del dia " & fecha & " " & Format(AvanceProc, "##0.00 %")
          contar = contar + 1
          DoEvents
       End If
   Next i
   nr = contar
   exito = True
Else
   exito = False
   MensajeProc = "No hay datos del segundo vector de precios del dia " & fecha
End If
nr = contar
End Sub

Sub ImpVectPreciosMD(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal noreg As Long)
Dim fecha As Date
Dim nomarch As String
Dim sihayarch As Boolean
Dim base1 As DAO.Database
Dim base2 As DAO.Database
Dim registros1 As DAO.recordset
Dim registros2 As DAO.recordset
Dim cont1 As Long
Dim notablas As Long
Dim inbucle As TableDef
Dim i As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
 noreg = 0
 fecha = fecha1
  Do While fecha <= fecha2
   'se construye el nombre del archivo
    nomarch = DirVPrecios & "\" & "PIP" & Format(fecha, "YYYYMMDD") & "M.xls"
    sihayarch = VerifAccesoArch(nomarch)
'  se lee desde un archivo de texto
    If sihayarch Then
       Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
       notablas = base1.TableDefs.Count
       ReDim mattabla(1 To notablas) As String
       cont1 = 1
       For Each inbucle In base1.TableDefs
      mattabla(cont1) = inbucle.Name
      cont1 = cont1 + 1
     Next inbucle
     Set registros1 = base1.OpenRecordset(mattabla(1), dbOpenDynaset)
      If registros1.RecordCount <> 0 Then
      registros1.MoveLast
      noreg = registros1.RecordCount
      registros1.MoveFirst
      Set base2 = OpenDatabase(DirBases & "\vector precios 1", dbDriverNoPrompt, "N", ";")
      Set registros2 = base2.OpenRecordset("VECTOR PIP MD", dbOpenDynaset)
      base2.Execute "DELETE FROM [VECTOR PIP MD] WHERE FECHA = " & CDbl(fecha)
      For i = 1 To noreg
      ReDim mata(0 To 8) As Variant
      mata(0) = LeerTAccess(registros1, 0, i)
      mata(1) = LeerTAccess(registros1, 1, i)
      mata(2) = LeerTAccess(registros1, 2, i)
      mata(3) = LeerTAccess(registros1, 3, i)
      mata(4) = LeerTAccess(registros1, 4, i)
      mata(5) = LeerTAccess(registros1, 5, i)
      mata(6) = LeerTAccess(registros1, 6, i)
      mata(7) = LeerTAccess(registros1, 7, i)
      mata(8) = LeerTAccess(registros1, 8, i)
      If Not IsNull(mata(0)) Then
        registros2.AddNew
        Call GrabarTAccess(registros2, 0, fecha, i)
        Call GrabarTAccess(registros2, 1, mata(1), i)
        Call GrabarTAccess(registros2, 2, mata(2), i)
        Call GrabarTAccess(registros2, 3, mata(3), i)
        Call GrabarTAccess(registros2, 4, Val(mata(4)), i)
        Call GrabarTAccess(registros2, 5, Val(mata(5)), i)
        Call GrabarTAccess(registros2, 6, Val(mata(6)), i)
        Call GrabarTAccess(registros2, 7, Val(mata(7)), i)
        Call GrabarTAccess(registros2, 8, Val(mata(8)), i)
        registros2.Update
       End If
       registros1.MoveNext
      Next i
      registros2.Close
      base2.Close
      End If
      registros1.Close
      base1.Close
      End If
    fecha = fecha + 1
  Loop
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function LeerVectPreciosMD(ByVal txtnomarch As String)
Dim sihayarch As Boolean
Dim notablas As Integer
Dim cont1 As Integer
Dim inbucle As TableDef
Dim registros1 As DAO.recordset
Dim noreg As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer

'On Error GoTo hayerror
 noreg = 0
    'se construye el nombre del archivo
    sihayarch = VerifAccesoArch(txtnomarch)
'  se lee desde un archivo de texto
    If sihayarch Then
       Dim base1 As DAO.Database
       Set base1 = OpenDatabase(txtnomarch, dbDriverNoPrompt, False, VersExcel)
       notablas = base1.TableDefs.Count
       ReDim mattabla(1 To notablas) As String
       cont1 = 1
       For Each inbucle In base1.TableDefs
           mattabla(cont1) = inbucle.Name
           cont1 = cont1 + 1
       Next inbucle
       Set registros1 = base1.OpenRecordset(mattabla(1), dbOpenDynaset)
       If registros1.RecordCount <> 0 Then
          registros1.MoveLast
          noreg = registros1.RecordCount
          nocampos = registros1.Fields.Count
          ReDim mata(1 To noreg, 1 To nocampos + 1) As Variant
          registros1.MoveFirst
          For i = 1 To noreg
              For j = 1 To nocampos
                  mata(i, j) = registros1.Fields(j - 1)
              Next j
              If EsVariableVacia(mata(i, 2)) Then mata(i, 2) = ""
              If EsVariableVacia(mata(i, 3)) Then mata(i, 3) = ""
              If EsVariableVacia(mata(i, 4)) Then mata(i, 4) = ""
              mata(i, nocampos + 1) = GeneraClaveEmision(mata(i, 2), mata(i, 3), mata(i, 4))
              registros1.MoveNext
              AvanceProc = i / noreg
              MensajeProc = "Leyendo el vector de precios MD: " & Format(AvanceProc, "##0.00 %")
              DoEvents
          Next i
          registros1.Close
          base1.Close
       Else
          ReDim mata(0 To 0, 0 To 0) As Variant
       End If
    Else
       ReDim mata(0 To 0, 0 To 0) As Variant
    End If
LeerVectPreciosMD = mata
'Exit Function
'hayerror:
'MsgBox error(Err())
End Function

Sub ImpPosFid(ByRef mata() As Variant, ByVal nofid As Integer, ByRef nreg As Long, ByRef exito As Boolean)
Dim matf() As Variant
Dim matvp() As Variant
Dim sihayarch As Boolean
Dim i As Long
Dim indice As Long
Dim txtfiltro As String
Dim noreg As Long

'rutina para leer la posicion del fondo de pensiones
exito = False
    If nofid = 2065 Then
       mata = VerificarPosMesaPen(mata, ClavePosPension1)
    Else
       mata = VerificarPosMesaPen(mata, ClavePosPension2)
    End If
    matf = ObtFactUnicos(mata, 1)
    If nofid = 2065 Then
       Call GuardaPosFPMD(mata, matf, 1, "Real", "000000", ClavePosPension1, nreg, exito)
       Call GuardaPosFPDiv(mata, matf, 1, "Real", "000000", ClavePosPension1, nreg, exito)
    Else
       Call GuardaPosFPMD(mata, matf, 1, "Real", "000000", ClavePosPension2, nreg, exito)
       Call GuardaPosFPDiv(mata, matf, 1, "Real", "000000", ClavePosPension2, nreg, exito)
    End If
  
End Sub


Function LeerArchPosPen(ByVal nomarch As String) As Variant()
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
 
 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
 Set registros1 = base1.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)
 
 'se revisa si hay registros en la tabla
If registros1.RecordCount <> 0 Then
 registros1.MoveLast
 noreg = registros1.RecordCount
 registros1.MoveFirst
 nocampos = registros1.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos) As Variant
 For i = 1 To noreg
  For j = 1 To nocampos
  mata(i, j) = LeerTAccess(registros1, j - 1, i)
  Next j
  registros1.MoveNext
  AvanceProc = i / noreg
  MensajeProc = "Leyendo el archivo de fondo " & Format(AvanceProc, "##0.00 %")
  DoEvents
 Next i
 registros1.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerArchPosPen = mata
End Function

Function LeerArchCalificaciones(ByVal nomarch As String) As Variant()
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
 
 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
 Set registros1 = base1.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)
 
 'se revisa si hay registros en la tabla
If registros1.RecordCount <> 0 Then
 registros1.MoveLast
 noreg = registros1.RecordCount
 registros1.MoveFirst
 nocampos = registros1.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos) As Variant
 For i = 1 To noreg
  For j = 1 To nocampos
  mata(i, j) = LeerTAccess(registros1, j - 1, i)
  Next j
  registros1.MoveNext
  AvanceProc = i / noreg
  MensajeProc = "Leyendo el archivo de fondo " & Format(AvanceProc, "##0.00 %")
  DoEvents
 Next i
 registros1.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerArchCalificaciones = mata
End Function






Sub GuardaFactROracle(ByVal fecha1 As Date, ByRef mata() As Variant)
Dim nnr As Long
Dim OraSession As Object
Dim OraDatabase As Object
Dim OraDynaset As Object
Dim txtfecha As String
Dim txtborra As String
Dim txtfiltro1 As String
Dim i As Long
Dim noreg As Long

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'rutina para guardar las tasas y demas factores de riesgo a una
'tabla de datos de oracle
nnr = UBound(mata, 1)
If nnr >= 1 Then

'se realiza el filtrado
txtfecha = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtborra = "delete FROM " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha
 ConAdo.Execute txtborra
 txtfiltro1 = "SELECT * FROM " & TablaFRiesgoO
 Set OraSession = CreateObject("OracleInProcServer.XOraSession")
 Set OraDatabase = OraSession.DbOpenDatabase("driesgos", "riesgo/riesgo", 0&)
 Set OraDynaset = OraDatabase.DbCreateDynaset(txtfiltro1, 0&)
  noreg = UBound(mata, 1)
 OraDynaset.DbMovelast
  For i = 1 To noreg
   OraDynaset.AddNew
   OraDynaset.Fields("FECHA").value = mata(i, 1)
   OraDynaset.Fields("CONCEPTO").value = Left(mata(i, 3), 15)
   OraDynaset.Fields("PLAZO").value = mata(i, 4)
   OraDynaset.Fields("VALOR").value = mata(i, 5)
   OraDynaset.Fields("INDICE").value = CLng(mata(i, 1)) & Trim(mata(i, 3)) & Format(mata(i, 4), "0000000")
   OraDynaset.Update
  Next i
   OraDynaset.Database.Close
   OraDatabase.Close
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function CargaIndVPrecios()
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim nocampos As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se cargan los indices que se necesitan del vector de precios para su
'importacion a un archivo de factores de riesgo

'====================================================
txtfiltro1 = "select count(*) from " & PrefijoBD & TablaIndVecPreciosO
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
 If noreg <> 0 Then
  txtfiltro1 = "select * from " & PrefijoBD & TablaIndVecPreciosO
  rmesa.Open txtfiltro1, ConAdo
  nocampos = rmesa.Fields.Count
  ReDim mata(1 To noreg, 1 To nocampos) As Variant
  rmesa.MoveFirst
  For i = 1 To noreg
   For j = 1 To nocampos
    mata(i, j) = rmesa.Fields(j - 1)
   Next j
   rmesa.MoveNext
   AvanceProc = i / noreg
   Call MostrarMensajeSistema("Cargando la lista de Indices en el VP: " & Format(AvanceProc, "##0.00 %"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
  Next i
  rmesa.Close
 Else
  ReDim mata(0 To 0, 0 To 0) As Variant
 End If
 CargaIndVPrecios = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub LeerTablaFactR(ByRef mata() As Variant, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtmens As String)
'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim noreg As Long
Dim tiempo1 As Date
Dim tiempo2 As Date
Dim i As Long
Dim rmesa1 As New ADODB.recordset
Dim rmesa2 As New ADODB.recordset


'====================================================
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "select * from " & TablaFRiesgoO & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <=" & txtfecha2
txtfiltro2 = txtfiltro2 & " AND (CONCEPTO,PLAZO) IN (SELECT CONCEPTO,PLAZO FROM " & PrefijoBD & TablaPortFR
txtfiltro2 = txtfiltro2 & " WHERE PORTAFOLIO = '" & NombrePortFR & "') ORDER BY INDICE"
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"
rmesa1.Open txtfiltro1, ConAdo
noreg = rmesa1.Fields(0).value
rmesa1.Close
tiempo1 = Time
If noreg <> 0 Then
   ReDim mata(1 To noreg, 1 To 5) As Variant
   rmesa2.Open txtfiltro2, ConAdo, adOpenUnspecified, adLockReadOnly
   rmesa2.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa2.Fields("FECHA").value
       mata(i, 2) = Trim(rmesa2.Fields("CONCEPTO").value)
       mata(i, 3) = Val(rmesa2.Fields("PLAZO").value)
       mata(i, 4) = CDbl(rmesa2.Fields("VALOR").value)
       mata(i, 5) = CLng(mata(i, 1)) & mata(i, 2) & Format(mata(i, 3), "0000000")
       rmesa2.MoveNext
       AvanceProc = i / noreg
       MensajeProc = txtmens & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   rmesa2.Close
   tiempo2 = Time
   mata = RutinaOrden(mata, 5, SRutOrden)
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function LeerBaseValExtO(ByVal fecha As Date, ByVal inc As String) As Variant()
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
'====================================================
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "select * from " & TablaValExtO & " WHERE FECHA = " & txtfecha & " AND INCDEC = '" & inc & "' ORDER BY CONCEPTO, PLAZO"
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 6) As Variant
'se seleccionan los registros que pueda haber con esta fecha
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("FECHA")
       mata(i, 2) = Trim(rmesa.Fields("CONCEPTO"))
       mata(i, 3) = Val(rmesa.Fields("PLAZO"))
       mata(i, 4) = rmesa.Fields("FVALOR")
       mata(i, 5) = CDbl(rmesa.Fields("VALOR"))
       mata(i, 6) = mata(i, 2) & " " & mata(i, 3)
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo los factores de riesgo: " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   mata = RutinaOrden(mata, 6, SRutOrden)
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerBaseValExtO = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub CrearEscExtFR(ByVal fecha As Date, ByVal txtfactor As String, ByVal plazo As Long, ByVal finicio As Date, ByRef txtmsg As String, ByRef exito As Boolean)
Dim exito1 As Boolean
Dim valormax As Double
Dim valormin As Double
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim mata() As Variant
Dim noreg As Long
Dim i As Long
Dim rend As Double
Dim txtcadena As String
Dim fechamin As Date
Dim fechamax As Date
mata = Leer1FRiesgoxVaR(finicio, fecha, txtfactor, plazo, exito)
noreg = UBound(mata, 1)
valormax = 0
valormin = 0
fechamin = 0
fechamax = 0
For i = 1 To noreg - 1
    rend = CalcRend2(mata(i, 2), mata(i + 1, 2), "TASA")
    If rend > valormax Then
       valormax = rend
       fechamax = mata(i + 1, 1)
    End If
    If rend < valormin Then
       valormin = rend
       fechamin = mata(i + 1, 1)
    End If
Next i
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha1 = "to_date('" & Format(fechamin, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fechamax, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtcadena = "INSERT INTO " & TablaValExtO & " VALUES("
txtcadena = txtcadena & txtfecha & ","            'FECHA
txtcadena = txtcadena & "'min',"                  'esc
txtcadena = txtcadena & "'" & txtfactor & "',"    'concepto
txtcadena = txtcadena & plazo & ","               'plazo
txtcadena = txtcadena & valormin & ","            'valor
txtcadena = txtcadena & txtfecha1 & ")"            'fecha esc
ConAdo.Execute txtcadena
txtcadena = "INSERT INTO " & TablaValExtO & " VALUES("
txtcadena = txtcadena & txtfecha & ","            'FECHA
txtcadena = txtcadena & "'max',"                  'esc
txtcadena = txtcadena & "'" & txtfactor & "',"    'concepto
txtcadena = txtcadena & plazo & ","               'plazo
txtcadena = txtcadena & valormax & ","            'valor
txtcadena = txtcadena & txtfecha2 & ")"            'fecha esc
ConAdo.Execute txtcadena
txtmsg = "Proceso finalizado correctamente"
exito = True
End Sub

Function Leer1FactorR(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtfactor As String, ByVal plazo As Integer) As Variant()
'la tabla de datos de Oracle
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset

'====================================================
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "select * from " & TablaFRiesgoO & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <=" & txtfecha2 & "AND CONCEPTO = '" & txtfactor & "' AND PLAZO = " & plazo & " ORDER BY FECHA"
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 2) As Variant
'se borran los registros que pueda haber con esta fecha
 rmesa.Open txtfiltro2, ConAdo
 rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("FECHA")
     mata(i, 2) = CDbl(rmesa.Fields("VALOR"))
     rmesa.MoveNext
     MensajeProc = "Leyendo el factor de riesgo " & txtfactor
     DoEvents
 Next i
rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
Leer1FactorR = mata
End Function

Function Leer1FactorRC(ByVal txtfactor As String, ByVal plazo As Integer) As Variant()
'la tabla de datos de Oracle
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset


'====================================================

txtfiltro2 = "select * from " & TablaFRiesgoO & " WHERE CONCEPTO = '" & txtfactor & "' AND PLAZO = " & plazo & " ORDER BY FECHA"
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 2) As Variant
'se borran los registros que pueda haber con esta fecha
 rmesa.Open txtfiltro2, ConAdo
 rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("FECHA")
     mata(i, 2) = CDbl(rmesa.Fields("VALOR"))
     rmesa.MoveNext
     AvanceProc = i / noreg
     MensajeProc = "Leyendo el factor de riesgo " & txtfactor & " : " & Format(AvanceProc, "###0.00 %")
     DoEvents
 Next i
rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
Leer1FactorRC = mata
End Function

Function ExtraerFactorMatFR(ByVal txtfactor As String, ByRef matfr() As Variant, ByRef matd() As Variant) As Variant()
'esta rutina obtiene la informacion del factor de riesgo de la matriz MatFactRiesgo
Dim noreg As Long
Dim indice As Long
Dim i As Long

noreg = UBound(matfr, 1)
indice = 0
For i = 1 To NoFactores
If matd(i, 1) = txtfactor Then
 indice = i
 Exit For
End If
Next i
If indice <> 0 Then
 ReDim matf(1 To noreg, 1 To 2) As Variant
 For i = 1 To noreg
  matf(i, 1) = matfr(i, 1)
  matf(i, 2) = matfr(i, indice + 1)
 Next i
Else
 MsgBox "no existe el factor en la tabla de datos"
 ReDim matf(0 To 0, 0 To 0) As Variant
End If
ExtraerFactorMatFR = matf
End Function

Function LlenarVaciosSerie(ByRef mata() As Variant)
Dim noreg As Long
Dim noreg1 As Long
Dim i As Long
Dim indice As Long

'esta rutina debe de llenar los huecos faltantes en de la TPFB

noreg = UBound(mata, 1)
If noreg <> 0 Then
noreg1 = mata(noreg, 1) - mata(1, 1) + 1 'son los dias efectivos que hay en el vector
If noreg1 >= 1 Then
ReDim matb(1 To noreg1, 1 To 2) As Variant
matb(1, 1) = mata(1, 1)
matb(1, 2) = mata(1, 2)
For i = 1 To noreg1
If i > 1 Then
matb(i, 1) = matb(i - 1, 1) + 1
indice = BuscarValorArray(matb(i, 1), mata, 1)
If indice <> 0 Then
 matb(i, 2) = mata(indice, 2)
Else
 matb(i, 2) = matb(i - 1, 2)
End If
End If
Next i
Else
ReDim matb(1 To 1, 1 To 2) As Variant
End If
Else
ReDim matb(1 To 1, 1 To 2) As Variant
End If
LlenarVaciosSerie = matb
End Function

Sub CrearMatFRiesgo2(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef mfriesgo() As Variant, ByVal txtmsg As String, ByRef exito As Boolean)
If ActivarControlErrores Then
 On Error GoTo hayerror
End If
Dim mattas1() As Variant
Dim mfechas() As Date
Dim i As Long
Dim j As Long
Dim indice As Long
Dim clave As String
Dim nodays As Long

  'se leen las fechas para las cuales hay factores de riesgo
exito = False
mfechas = LeerFechasVaR(fecha1, fecha2)
nodays = UBound(mfechas, 1)
If nodays <> 0 Then
   ReDim mfriesgo(1 To nodays, 1 To NoFactores + 1) As Variant
   For i = 1 To nodays
       mfriesgo(i, 1) = mfechas(i, 1)
   Next i
    Call LeerTablaFactR(mattas1, fecha1, fecha2, "Leyendo factores de riesgo")
  'se crea la estructura de la matriz con el arreglo
 'se procede a poner la informacion en la matriz de factores de riesgo
    For i = 1 To NoFactores
         For j = 1 To nodays
             clave = CLng(mfriesgo(j, 1)) & Trim(MatCaracFRiesgo(i).nomFactor) & Format(MatCaracFRiesgo(i).plazo, "0000000")
             indice = BuscarValorArray(clave, mattas1, 5)
             If indice <> 0 Then
                mfriesgo(j, i + 1) = mattas1(indice, 4)
             Else
                mfriesgo(j, i + 1) = 0
             End If
         Next j
         AvanceProc = i / NoFactores
         MensajeProc = "Procesando la información " & Format(AvanceProc, "#00.00 %")
    Next i
    txtmsg = "El proceso finalizo correctamente"
    exito = True
Else
   ReDim mfriesgo(0 To 0, 0 To 0) As Variant
   txtmsg = "No hay datos"
   exito = False
End If
On Error GoTo 0
Exit Sub
hayerror:
MsgBox "CrearMatFriesgo2: " & error(Err())
On Error GoTo 0
End Sub

Sub CrearMatFRiesgo3(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef mfriesgo() As Variant, ByVal txtmsg As String, ByRef exito As Boolean)
If ActivarControlErrores Then
 On Error GoTo hayerror
End If
Dim matfact() As Variant
Dim mfechas() As Date
Dim i As Long
Dim j As Long
Dim indice As Long
Dim clave As String
Dim nodays As Long

  'se leen las fechas para las cuales hay factores de riesgo
exito = True
mfechas = LeerFechasVaR(fecha1, fecha2)
nodays = UBound(mfechas, 1)
If nodays <> 0 Then
   ReDim mfriesgo(1 To nodays, 1 To NoFactores + 1) As Variant
   For i = 1 To nodays
       mfriesgo(i, 1) = mfechas(i, 1)
   Next i
 'se procede a poner la informacion en la matriz de factores de riesgo
    For i = 1 To NoFactores
         matfact = Leer1FactorR(fecha1, fecha2, MatCaracFRiesgo(i).nomFactor, MatCaracFRiesgo(i).plazo)
         If UBound(matfact, 1) <> 0 Then
         For j = 1 To nodays
             indice = BuscarValorArray(mfechas(j, 1), matfact, 1)
             If indice <> 0 Then
                mfriesgo(j, i + 1) = matfact(indice, 2)
             Else
                mfriesgo(j, i + 1) = 0
                exito = False
             End If
         Next j
         End If
         AvanceProc = i / NoFactores
         MensajeProc = "Procesando la información " & Format(AvanceProc, "#00.00 %")
         DoEvents
    Next i
    txtmsg = "El proceso finalizo correctamente"
    exito = True
Else
   ReDim mfriesgo(0 To 0, 0 To 0) As Variant
   txtmsg = "No hay datos"
   exito = False
End If
On Error GoTo 0
Exit Sub
hayerror:
MsgBox "CrearMatFriesgo2: " & error(Err())
On Error GoTo 0
End Sub


Function CargaTProd(ByVal nombase As String, ByVal txtbase As String) As Variant()
Dim txtfiltro As String
Dim noreg As Integer
Dim i As Integer
Dim nocampos As Integer
Dim j As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
txtfiltro = "select count(*) from " & nombase
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtfiltro = "select * from " & nombase
   rmesa.Open txtfiltro, ConAdo
   rmesa.MoveFirst
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   For i = 1 To noreg
       For j = 1 To nocampos
           mata(i, j) = rmesa.Fields(j - 1)
       Next j
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Cargando la matriz de Productos " & txtbase & " " & Format(AvanceProc, "###0.00")
   Next i
   rmesa.Close
 Else
   ReDim mata(0 To 0, 0 To 0) As Variant
 End If
 CargaTProd = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub GenerarFechasTareas(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef exito As Boolean)
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro As String
Dim fecha As Date
Dim txtfecha As String
Dim txtcadena As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
exito = False
'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
'====================================================
If IsDate(fecha1) And IsDate(fecha2) Then
 txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfiltro = "delete from " & TablaFechasTareas1 & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2
 ConAdo.Execute txtfiltro
 fecha = fecha1
 Do While fecha <= fecha2
    If Not NoLabMX(fecha) Then
       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaFechasTareas1 & " VALUES("
       txtcadena = txtcadena & txtfecha & ")"
       ConAdo.Execute txtcadena
    End If
  fecha = fecha + 1
  If fecha1 <> fecha2 Then
     AvanceProc = (fecha - fecha1) / (fecha2 - fecha1)
     MensajeProc = "Generando las fechas para generacion de tareas " & Format(100 * (fecha - fecha1) / (fecha2 - fecha1), "###0.00")
  Else
     MensajeProc = "Generando las fechas para generacion de tareas"
  End If
  DoEvents
 Loop
 If fecha = fecha2 + 1 Then exito = True
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox "GenerarFechasVaRO: " & error(Err())
On Error GoTo 0
End Sub

Sub GenerarFechasTareasPos(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef exito As Boolean)
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim fecha As Date
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtcadena As String
Dim txtfiltro As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
exito = False
'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
'====================================================
If IsDate(fecha1) And IsDate(fecha2) Then
 txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfiltro = "delete from " & TablaFechasTareas1 & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2
 ConAdo.Execute txtfiltro
 fecha = fecha1
 Do While fecha <= fecha2
    If Not NoLabMX(fecha) Then
       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaFechasTareas1 & " VALUES("
       txtcadena = txtcadena & txtfecha & ")"
       ConAdo.Execute txtcadena
    End If
  fecha = fecha + 1
  If fecha1 <> fecha2 Then
     AvanceProc = (fecha - fecha1) / (fecha2 - fecha1)
     MensajeProc = "Generando las fechas para generacion de tareas " & Format(100 * (fecha - fecha1) / (fecha2 - fecha1), "###0.00")
  Else
     MensajeProc = "Generando las fechas para generacion de tareas"
  End If
  DoEvents
 Loop
 If fecha = fecha2 + 1 Then exito = True
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox "GenerarFechasVaRO: " & error(Err())
On Error GoTo 0
End Sub

Sub GenerarFechasVaR(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txttabla As String
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txtfiltro As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim fecha As Date
Dim txtcadena As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
exito = False
'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
txttabla = TablaFechasVaR
'====================================================
If IsDate(fecha1) And IsDate(fecha2) Then
 txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfiltro = "delete from " & txttabla & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2
 ConAdo.Execute txtfiltro
 fecha = fecha1
 Do While fecha <= fecha2 And fecha <= Date
    If Not NoLabMX(fecha) Then
       txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtcadena = "INSERT INTO " & TablaFechasVaR & " VALUES("
       txtcadena = txtcadena & txtfecha & ")"
       ConAdo.Execute txtcadena
    End If
    fecha = fecha + 1
    If fecha1 <> fecha2 Then
       AvanceProc = (fecha - fecha1) / (fecha2 - fecha1)
       MensajeProc = "Generando las fechas para VaR " & Format(100 * (fecha - fecha1) / (fecha2 - fecha1), "###0.00")
    Else
       MensajeProc = "Generando las fechas para VaR"
    End If
 Loop
 If fecha = fecha2 + 1 Then
    exito = True
    txtmsg = "El proceso finalizo correctamente"
 End If
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox "GenerarFechasVaRO: " & error(Err())
On Error GoTo 0
End Sub

Sub GenerarFechasFR(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef exito As Boolean)
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro As String
Dim fecha As Date
Dim txtcadena As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
exito = False
'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
'====================================================
If IsDate(fecha1) And IsDate(fecha2) Then
   txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtfiltro = "delete from " & TablaFechasFR & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2
 ConAdo.Execute txtfiltro
 fecha = fecha1
 Do While fecha <= fecha2
  If Not NoLabMX(fecha) Or Not NolabUS(fecha) Or (Month(fecha + 1) <> Month(fecha)) Then
    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "INSERT INTO " & TablaFechasFR & " VALUES("
    txtcadena = txtcadena & txtfecha & ")"
    ConAdo.Execute txtcadena
  End If
  fecha = fecha + 1
  If fecha1 <> fecha2 Then
     AvanceProc = (fecha - fecha1) / (fecha2 - fecha1)
     MensajeProc = "Generando las fechas para FR " & Format(100 * (fecha - fecha1) / (fecha2 - fecha1), "###0.00")
  Else
     MensajeProc = "Generando las fechas para FR"
  End If
  Loop
 If fecha = fecha2 + 1 Then exito = True
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox "LeerFechasFRO: " & error(Err())
On Error GoTo 0
End Sub

Function LeerFechasTareas(ByVal fecha1 As Date, ByVal fecha2 As Date) As Date()
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim noreg As Long
Dim i As Long
Dim txtfiltro1 As String
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
'====================================================
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
sql_num_mesa = "select count(distinct fecha) from " & TablaFechasTareas1 & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2
txtfiltro1 = "select FECHA from " & TablaFechasTareas1 & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & "  GROUP BY FECHA ORDER BY FECHA"
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Date
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("FECHA")
     rmesa.MoveNext
     AvanceProc = i / noreg
     MensajeProc = "Carga las fechas con datos de factores de riesgo " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Date
End If
LeerFechasTareas = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox "LeerFechasTareas: " & error(Err())
On Error GoTo 0
End Function

Function LeerFechasTareasT() As Date()
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txtfiltro1 As String
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
'====================================================
sql_num_mesa = "select count(distinct fecha) from " & TablaFechasTareas1
txtfiltro1 = "select FECHA from " & TablaFechasTareas1 & " GROUP BY FECHA ORDER BY FECHA"
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Date
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
 For i = 1 To noreg
  mata(i, 1) = rmesa.Fields("FECHA")
  rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Carga las fechas con datos de factores de riesgo " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Date
End If
LeerFechasTareasT = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox "LeerFechasTareas: " & error(Err())
On Error GoTo 0
End Function

Function LeerFechasTareasPos(ByVal fecha1 As Date, ByVal fecha2 As Date) As Date()
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
'====================================================

txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
sql_num_mesa = "select count(distinct fecha) from " & TablaFechasTareas1 & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2
txtfiltro1 = "select FECHA from " & TablaFechasTareas1 & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & "  GROUP BY FECHA ORDER BY FECHA"
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Date
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
 For i = 1 To noreg
  mata(i, 1) = rmesa.Fields("FECHA")
  rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Carga las fechas con datos de factores de riesgo " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Date
End If
LeerFechasTareasPos = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox "LeerFechasTareas: " & error(Err())
On Error GoTo 0
End Function

Function LeerFechasTareasPosT() As Date()
Dim txttabla As String
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim i As Integer
Dim noreg As Integer
Dim txtfiltro1 As String
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle

'====================================================
sql_num_mesa = "select count(distinct fecha) from " & TablaFechasTareas1
txtfiltro1 = "select FECHA from " & TablaFechasTareas1 & " GROUP BY FECHA ORDER BY FECHA"
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Date
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
 For i = 1 To noreg
  mata(i, 1) = rmesa.Fields("FECHA")
  rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Carga las fechas con datos de factores de riesgo " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Date
End If
LeerFechasTareasPosT = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox "LeerFechasTareas: " & error(Err())
On Error GoTo 0
End Function

Function LeerFechasVaR(ByVal fecha1 As Date, ByVal fecha2 As Date) As Date()

Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle

'====================================================

txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
sql_num_mesa = "select count(distinct fecha) from " & TablaFechasVaR & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 '& " GROUP BY FECHA"
txtfiltro1 = "select FECHA from " & TablaFechasVaR & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & "  GROUP BY FECHA ORDER BY FECHA"
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Date
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("FECHA")
     rmesa.MoveNext
     AvanceProc = i / noreg
     MensajeProc = "Carga las fechas con datos de factores de riesgo " & Format(AvanceProc, "###0.00")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Date
End If
LeerFechasVaR = mata
End Function

Function LeerFechasVaRT() As Date()
Dim txttabla As String
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txtfiltro1 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
txttabla = TablaFechasVaR
'====================================================
sql_num_mesa = "select count(distinct fecha) from " & txttabla
txtfiltro1 = "select FECHA from " & txttabla & " GROUP BY FECHA ORDER BY FECHA"
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Date
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("FECHA")
     rmesa.MoveNext
     AvanceProc = i / noreg
     MensajeProc = "Carga las fechas con datos de factores de riesgo " & Format(AvanceProc, "###0.00")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Date
End If
LeerFechasVaRT = mata
End Function

Function LeerRangoFechas(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef matf() As Date) As Date()
Dim indice1 As Long
Dim indice2 As Long
Dim i As Long

Do While True
   indice1 = BuscarValorArray(fecha1, matf, 1)
   If indice1 <> 0 Then
      Exit Do
   End If
   fecha1 = fecha1 + 1
Loop

Do While True
   indice2 = BuscarValorArray(fecha2, matf, 1)
   If indice2 <> 0 Then
      Exit Do
   End If
   fecha2 = fecha2 - 1
Loop

ReDim matfs(indice2 - indice1 + 1, 1 To 1) As Date
For i = 1 To indice2 - indice1 + 1
matfs(i, 1) = matf(indice1 - 1 + i, 1)
Next i
LeerRangoFechas = matfs
End Function


Function LeerFechasFR(ByVal fecha1 As Date, ByVal fecha2 As Date) As Date()
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle
'====================================================
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
sql_num_mesa = "select count(distinct fecha) from " & TablaFechasFR & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 '& " GROUP BY FECHA"
txtfiltro1 = "select FECHA from " & TablaFechasFR & " WHERE FECHA >= " & TablaFechasFR & " AND FECHA <= " & txtfecha2 & "  GROUP BY FECHA ORDER BY FECHA"
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Date
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
 For i = 1 To noreg
   mata(i, 1) = rmesa.Fields("FECHA")
   rmesa.MoveNext
   AvanceProc = i / noreg
   MensajeProc = "Carga las fechas con datos de factores de riesgo " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close

Else
 ReDim mata(0 To 0, 0 To 0) As Date
End If
LeerFechasFR = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox "LeerFechasFRO: " & error(Err())
On Error GoTo 0
End Function

Function LeerFechasFRT() As Date()
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txttabla  As String
Dim txtfiltro1 As String
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle

'====================================================
sql_num_mesa = "select count(distinct fecha) from " & TablaFechasFR
txtfiltro1 = "select FECHA from " & TablaFechasFR & " GROUP BY FECHA ORDER BY FECHA"
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Date
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
 For i = 1 To noreg
   mata(i, 1) = rmesa.Fields("FECHA")
   rmesa.MoveNext
   AvanceProc = i / noreg
   MensajeProc = "Carga las fechas con datos de factores de riesgo " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close

Else
 ReDim mata(0 To 0, 0 To 0) As Date
End If
LeerFechasFRT = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox "LeerFechasFRO: " & error(Err())
On Error GoTo 0
End Function

Function LeerFechasFR2O(ByVal fecha1 As Date, ByVal fecha2 As Date) As Date()
Dim sql_mesa As String
Dim sql_num_mesa As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'rutina para leer la posicion del mesa desde
'la tabla de datos de Oracle

'====================================================
If IsDate(fecha1) And IsDate(fecha2) Then
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 sql_num_mesa = "select count(distinct fecha) from " & TablaFRiesgoO & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 '& " GROUP BY FECHA"
 txtfiltro1 = "select FECHA from " & TablaFRiesgoO & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & "  GROUP BY FECHA ORDER BY FECHA"
Else
 sql_num_mesa = "select count(distinct fecha) from " & TablaFRiesgoO
 txtfiltro1 = "select FECHA from " & TablaFRiesgoO & " GROUP BY FECHA ORDER BY FECHA"
End If
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Date
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
 For i = 1 To noreg
  mata(i, 1) = rmesa.Fields("FECHA")
  rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Carga las fechas con datos de factores de riesgo " & Format(AvanceProc, "###0.00")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Date
End If
LeerFechasFR2O = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox "LeerFechasFRO: " & error(Err())
On Error GoTo 0
End Function

Function LeerBlackList()
     LeerBlackList = CargaTablaD(PrefijoBD & TablaBlackList, "Cargando la lista negra", 1)
End Function

Function LeerOpValidada(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfiltro As String
  txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtfiltro = "SELECT * FROM " & TablaOperValidada & " WHERE FVALIDACION = " & txtfecha
  LeerOpValidada = CargaTablaD(txtfiltro, "Cargando la operaciones validadas", 1)
End Function

Function LeerSecuenciaProcesos()
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim nocampos As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'====================================================
txtfiltro1 = "select count(*) from " & PrefijoBD & TablaSecProcesos
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 txtfiltro1 = "select * from " & PrefijoBD & TablaSecProcesos
 rmesa.Open txtfiltro1, ConAdo
 nocampos = rmesa.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
  For j = 1 To nocampos
   mata(i, j) = rmesa.Fields(j - 1)
  Next j
 rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Cargando la secuencia de procesos " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerSecuenciaProcesos = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerSecuenciaSubproc()
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim nocampos As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'====================================================
txtfiltro1 = "select count(*) from " & PrefijoBD & TablaSecSubProc
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 txtfiltro1 = "select * from " & PrefijoBD & TablaSecSubProc
 rmesa.Open txtfiltro1, ConAdo
 nocampos = rmesa.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
  For j = 1 To nocampos
   mata(i, j) = rmesa.Fields(j - 1)
  Next j
 rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Cargando la secuencia de procesos " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerSecuenciaSubproc = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function


Function LeerCatProcesos()
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim nocampos As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'====================================================
txtfiltro = "select * from " & PrefijoBD & TablaCatProcesos & " ORDER BY ID_TAREA"
txtfiltro1 = "select count(*) from (" & PrefijoBD & TablaCatProcesos & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 rmesa.Open txtfiltro, ConAdo
 nocampos = rmesa.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos) As Variant
 rmesa.MoveFirst
 For i = 1 To noreg
     For j = 1 To nocampos
         mata(i, j) = rmesa.Fields(j - 1)
     Next j
 rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Cargando el catalogo de tareas a realizar " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerCatProcesos = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerProcesosF(ByVal fecha As Date, ByVal opcion As Integer) As Variant()
Dim txtfecha As String
Dim txtcadena As String
Dim txtfiltro1 As String
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim nocampos As Long
Dim contar0 As Long
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'====================================================
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtcadena = "from " & DetermTablaProc(opcion) & " WHERE FECHAP = " & txtfecha & " ORDER BY IDTAREA"
txtfiltro1 = "select count(*) " & txtcadena
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 txtfiltro1 = "select * " & txtcadena
 rmesa.Open txtfiltro1, ConAdo
 nocampos = rmesa.Fields.Count
 ReDim mata(1 To nocampos, 1 To 1) As Variant
 rmesa.MoveFirst
 contar0 = 0
 For i = 1 To noreg
 If rmesa.Fields(2) <> "LEER POS PRIM SWAPS" And rmesa.Fields(2) <> "LEER POS PRIM FWDS" And rmesa.Fields(2) <> "POS FONDO PENSIONES 1" And rmesa.Fields(2) <> "POS FONDO PENSIONES 2" And rmesa.Fields(2) <> "SINCRONIZAR BASES ORACLE ACCESS" And rmesa.Fields(2) <> "CALCULAR VAR" Then
  contar0 = contar0 + 1
  ReDim Preserve mata(1 To nocampos, 1 To contar0) As Variant
  For j = 1 To nocampos
   mata(j, contar0) = rmesa.Fields(j - 1)
  Next j
 End If
 rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Cargando las tareas pendientes de realizar " & Format(AvanceProc, "###0.00 %")
 Next i
 rmesa.Close
 If contar0 <> 0 Then
  mata = MTranV(mata)
 Else
 ReDim mata(0 To 0, 0 To 0) As Variant
 End If
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerProcesosF = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerListaPortPos()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim sql_num_mesa As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se cargan las estructuras de portafolio

'====================================================
sql_num_mesa = "select count(distinct GRUPO) from " & PrefijoBD & TablaGruposPortPos
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 1) As Variant
'se borran los registros que pueda haber con esta fecha
txtfiltro2 = "select GRUPO FROM " & PrefijoBD & TablaGruposPortPos & " GROUP BY GRUPO"
rmesa.Open txtfiltro2, ConAdo
rmesa.MoveFirst
 For i = 1 To noreg
     mata(i, 1) = rmesa.Fields("GRUPO")
     rmesa.MoveNext
     MensajeProc = "Cargando las estructuras de portafolios existentes"
 Next i
rmesa.Close
mata = RutinaOrden(mata, 1, SRutOrden)
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerListaPortPos = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerPortPos(ByVal fecha As Date, ByVal txtport As String) As Variant()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg As Long
Dim txtfecha As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mata(1 To noreg, 1 To 2) As Variant
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("CPOSICION")  'clave de posicion
       mata(i, 2) = rmesa.Fields("COPERACION")  'clave de operacion
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerPortPos = mata
End Function

Function CreartxtRVaRR(ByVal txtport As String, ByRef matrvar() As Variant, ByRef matrvart() As Variant, ByVal txtdesc As String, ByVal ncol1 As Long, ByVal ncol2 As Long, ByVal dv As Long, ByVal horiz As Long, ByVal nc As Long)
Dim txttexto As String
Dim noreg As Long
Dim i As Long

txttexto = ""
noreg = UBound(MatGruposPortPos, 1)
If ncol2 <> 0 Then
   For i = 1 To noreg
       If matrvar(i, ncol1) <> 0 Or matrvar(i, ncol2) <> 0 Then
          txttexto = txttexto & CreartxtResVaR3(txtport, MatGruposPortPos(i, 3), txtdesc, dv, horiz, nc, matrvar(i, ncol1), matrvar(i, ncol2))
       End If
       MensajeProc = "Creando cadena texto de CVaR " & txtport & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   If matrvart(ncol1) <> 0 Or matrvart(ncol2) <> 0 Then
      txttexto = txttexto & CreartxtResVaR3(txtport, MatPortafolios(1, 2), txtdesc, dv, horiz, nc, matrvart(ncol1), matrvart(ncol2))
   End If
Else
   For i = 1 To noreg
       If matrvar(i, ncol1) <> 0 Then
          txttexto = txttexto & CreartxtResVaR3(txtport, MatGruposPortPos(i, 3), txtdesc, dv, horiz, nc, matrvar(i, ncol1), 0)
       End If
       MensajeProc = "Guardando res CVaR " & txtport & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   If matrvart(ncol1) <> 0 Then
      txttexto = txttexto & CreartxtResVaR3(txtport, MatPortafolios(1, 2), txtdesc, dv, horiz, nc, matrvart(ncol1), 0)
      DoEvents
   End If
End If
CreartxtRVaRR = txttexto
End Function

Function CreartxtResVaR3(ByVal txtport As String, ByVal tipotitulos As Integer, ByVal txttvar As String, ByVal dvol As Long, ByVal horiz As Long, ByVal nconf As Double, ByVal liminf As Double, ByVal limsup As Double)
Dim txttexto As String
txttexto = ""
 txttexto = txttexto & txtport & "|"
 txttexto = txttexto & tipotitulos & "|"
 txttexto = txttexto & txttvar & "|"
 txttexto = txttexto & dvol & "|"
 txttexto = txttexto & horiz & "|"
 txttexto = txttexto & nconf & "|"
 txttexto = txttexto & liminf & "|"
 txttexto = txttexto & limsup & "/"
 CreartxtResVaR3 = txttexto
End Function


Sub GuardarResVaR2(ByVal fecha As Date, ByVal txtport As String, ByRef matrvar(), ByRef matrvart() As Double, ByVal txtdesc As String, ByVal ncol1 As Long, ByVal ncol2 As Long, ByVal dv As Long, ByVal horiz As Long, ByVal nc As Double, ByVal tp As String)
Dim noreg As Long
Dim txtcadena As String
Dim i As Long

'se guardan los resultados de var por tipo de agrupacion de instrumentos
noreg = UBound(MatGruposPortPos, 1)
If Not EsVariableVacia(tp) Then
   txtcadena = txtdesc & " " & tp
Else
   txtcadena = txtdesc
End If
If ncol2 <> 0 Then
   For i = 1 To noreg
       If matrvar(i, ncol1) <> 0 Or matrvar(i, ncol2) <> 0 Then
          Call GuardaDatosVAR(fecha, txtport, MatGruposPortPos(i, 3), txtcadena, dv, horiz, nc, matrvar(i, ncol1), matrvar(i, ncol2))
       End If
       MensajeProc = "Guardando res CVaR " & txtport & " " & fecha & " " & Format(AvanceProc, "###0.00")
       DoEvents
   Next i
   If matrvart(ncol1) <> 0 Or matrvart(ncol2) <> 0 Then
      Call GuardaDatosVAR(fecha, txtport, MatPortafolios(1, 2), txtcadena, dv, horiz, nc, matrvart(ncol1), matrvart(ncol2))
   End If
Else
   For i = 1 To noreg
       If matrvar(i, ncol1) <> 0 Then
          Call GuardaDatosVAR(fecha, txtport, MatGruposPortPos(i, 3), txtcadena, dv, horiz, nc, matrvar(i, ncol1), 0)
       End If
       MensajeProc = "Guardando res CVaR " & txtport & " " & fecha & " " & Format(AvanceProc, "###0.00")
       DoEvents
   Next i
   If matrvart(ncol1) <> 0 Then
      Call GuardaDatosVAR(fecha, txtport, MatPortafolios(1, 2), txtcadena, dv, horiz, nc, matrvart(ncol1), 0)
      DoEvents
   End If
End If

End Sub

Sub ImpRepVaRMDT(ByVal fecha As Date)
Dim matfilas() As Integer
Dim matcolumnas() As Integer
Dim filasimp() As Integer
Dim colsimp() As Integer
Dim vallim As Double
Dim valcons As Double
Dim varmd As Double
Dim fechacn As Date
Dim mattabla() As Variant
Dim nodim1 As Integer
Dim nodim2 As Integer
Dim nofilas As Integer
Dim nocols As Integer
Dim i As Integer
Dim salida As Object
Dim contar As Integer
Dim posy As Integer
Dim noesc As Integer
noesc = 500

'capital neto banobras
   CapitalNeto = DevLimitesVaR(CDate(fecha), MatCapitalSist, "CAPITAL NETO B") * 1000000
   fechacn = DevFechaLimite(CDate(fecha), MatCapitalSist, "CAPITAL NETO B")
   vallim = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR MD")
   varmd = LeerResVaRTabla(fecha, txtportCalc1, "MERCADO DE DINERO", "CVARH", noesc, 1, 0, 0.03, 0)
   If vallim * CapitalNeto <> 0 Then valcons = -varmd / (vallim * CapitalNeto)
   mattabla = GenRepVaRMD(fecha, "REPORTE MDT", CapitalNeto, varmd, vallim, valcons)
   nodim1 = UBound(mattabla, 1)
   nodim2 = UBound(mattabla, 2)
   nofilas = nodim1 + 1
   nocols = nodim2 + 1
   ReDim matfilas(1 To nofilas) As Integer
   ReDim matcolumnas(1 To nocols) As Integer
   matfilas(1) = 700
   For i = 2 To nofilas
       matfilas(i) = matfilas(i - 1) + 600
   Next i
   matcolumnas(1) = 1300
   matcolumnas(2) = 1300 + 3000
   For i = 3 To nocols
   matcolumnas(i) = matcolumnas(i - 1) + 1700
   Next i
'se definen las columnas y filas a imprimir
ReDim filasimp(1 To nofilas) As Integer, colsimp(1 To nocols) As Integer
For i = 1 To nofilas
 filasimp(i) = i
Next i

For i = 1 To nocols
 colsimp(i) = i
Next i


'SE deben buscar los consolidados de var de mesa de dinero que se generaron en el var
'global y pegarlos aqui
Set salida = Printer
contar = 2
salida.Orientation = 2
matfilas(1) = 2700
matfilas(2) = matfilas(1) + 300
Call ImpCabRepMDT(fecha, salida)
Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, 1)
Do While contar <= nodim1
posy = matfilas(contar)
If matfilas(contar) > 11000 Then
   matfilas(contar) = 2700 + 300
   salida.NewPage
   Call ImpCabRepMD(fecha, salida)
   Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, 1)
End If
'se imprime el cuadro con los resultados del VaR
   salida.FontSize = 6
   salida.FontBold = False
   If contar <= nodim1 Then matfilas(contar + 1) = matfilas(contar) + 270
   Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, contar)
 contar = contar + 1
Loop
salida.CurrentX = 0
salida.Print "             Nota: La posición de Mercado de Dinero comprende la posición de la Mesa de Dinero, Tesorería y Portafolio de Inversion Disponible para la venta"

salida.EndDoc
End Sub

Sub ImpRepVaRMD(ByVal fecha As Date)
Dim matfilas() As Integer
Dim matcolumnas() As Integer
Dim filasimp() As Integer
Dim colsimp() As Integer
Dim vallim As Double
Dim valcons As Double
Dim varmd As Double
Dim fechacn As Date
Dim mattabla() As Variant
Dim nodim1 As Integer
Dim nodim2 As Integer
Dim nofilas As Integer
Dim nocols As Integer
Dim i As Integer
Dim salida As Object
Dim contar As Integer
Dim posy As Integer

'capital neto banobras
   CapitalNeto = DevLimitesVaR(CDate(fecha), MatCapitalSist, "CAPITAL NETO B") * 1000000
   fechacn = DevFechaLimite(CDate(fecha), MatCapitalSist, "CAPITAL NETO B")
   vallim = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR MD1")
   varmd = LeerResVaRTabla(fecha, txtportCalc1, "MESA DE DINERO", "CVARH", 500, 1, 0, 0.03, 0)
   valcons = -varmd / (vallim * CapitalNeto)
   mattabla = GenRepVaRMD(fecha, "REPORTE MD", CapitalNeto, varmd, vallim, valcons)
   nodim1 = UBound(mattabla, 1)
   nodim2 = UBound(mattabla, 2)
   nofilas = nodim1 + 1
   nocols = nodim2 + 1
   ReDim matfilas(1 To nofilas) As Integer
   ReDim matcolumnas(1 To nocols) As Integer
   matfilas(1) = 700
   For i = 2 To nofilas
       matfilas(i) = matfilas(i - 1) + 600
   Next i
   matcolumnas(1) = 1300
   matcolumnas(2) = 1300 + 3000
   For i = 3 To nocols
   matcolumnas(i) = matcolumnas(i - 1) + 1700
   Next i
'se definen las columnas y filas a imprimir
ReDim filasimp(1 To nofilas) As Integer, colsimp(1 To nocols) As Integer
For i = 1 To nofilas
 filasimp(i) = i
Next i

For i = 1 To nocols
 colsimp(i) = i
Next i


'SE deben buscar los consolidados de var de mesa de dinero que se generaron en el var
'global y pegarlos aqui
Set salida = Printer
contar = 2
salida.Orientation = 2
matfilas(1) = 2700
matfilas(2) = matfilas(1) + 300
Call ImpCabRepMD(fecha, salida)
Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, 1)
Do While contar <= nodim1
posy = matfilas(contar)
If matfilas(contar) > 11000 Then
   matfilas(contar) = 2700 + 300
   salida.NewPage
   Call ImpCabRepMD(fecha, salida)
   Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, 1)
End If
'se imprime el cuadro con los resultados del VaR
   salida.FontSize = 6
   salida.FontBold = False
   If contar <= nodim1 Then matfilas(contar + 1) = matfilas(contar) + 270
   Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, contar)
 contar = contar + 1
Loop
salida.CurrentX = 0

salida.EndDoc

End Sub

Sub ImpRepVaRTeso(ByVal fecha As Date)
Dim matfilas() As Integer
Dim matcolumnas() As Integer
Dim filasimp() As Integer
Dim colsimp() As Integer
Dim vallim As Double
Dim valcons As Double
Dim varmd As Double
Dim fechacn As Date
Dim mattabla() As Variant
Dim nodim1 As Integer
Dim nodim2 As Integer
Dim nofilas As Integer
Dim nocols As Integer
Dim i As Integer
Dim salida As Object
Dim contar As Integer
Dim posy As Integer

'capital neto banobras
   CapitalNeto = DevLimitesVaR(CDate(fecha), MatCapitalSist, "CAPITAL NETO B") * 1000000
   fechacn = DevFechaLimite(CDate(fecha), MatCapitalSist, "CAPITAL NETO B")
   vallim = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR TESO")
   varmd = LeerResVaRTabla(fecha, txtportCalc1, "TESORERIA", "CVARH", 500, 1, 0, 0.03, 0)
   valcons = -varmd / (vallim * CapitalNeto)
   mattabla = GenRepVaRMD(fecha, "REPORTE TESO", CapitalNeto, varmd, vallim, valcons)
   nodim1 = UBound(mattabla, 1)
   nodim2 = UBound(mattabla, 2)
   nofilas = nodim1 + 1
   nocols = nodim2 + 1
   ReDim matfilas(1 To nofilas) As Integer
   ReDim matcolumnas(1 To nocols) As Integer
   matfilas(1) = 700
   For i = 2 To nofilas
       matfilas(i) = matfilas(i - 1) + 600
   Next i
   matcolumnas(1) = 1300
   matcolumnas(2) = 1300 + 3000
   For i = 3 To nocols
   matcolumnas(i) = matcolumnas(i - 1) + 1700
   Next i
'se definen las columnas y filas a imprimir
ReDim filasimp(1 To nofilas) As Integer, colsimp(1 To nocols) As Integer
For i = 1 To nofilas
 filasimp(i) = i
Next i

For i = 1 To nocols
 colsimp(i) = i
Next i


'SE deben buscar los consolidados de var de mesa de dinero que se generaron en el var
'global y pegarlos aqui
Set salida = Printer
contar = 2
salida.Orientation = 2
matfilas(1) = 2700
matfilas(2) = matfilas(1) + 300
Call ImpCabRepTeso(fecha, salida)
Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, 1)
Do While contar <= nodim1
posy = matfilas(contar)
If matfilas(contar) > 11000 Then
   matfilas(contar) = 2700 + 300
   salida.NewPage
   Call ImpCabRepMD(fecha, salida)
   Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, 1)
End If
'se imprime el cuadro con los resultados del VaR
   salida.FontSize = 6
   salida.FontBold = False
   If contar <= nodim1 Then matfilas(contar + 1) = matfilas(contar) + 270
   Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, contar)
 contar = contar + 1
Loop
salida.CurrentX = 0

salida.EndDoc

End Sub


Sub ImpRepDeriv(ByVal fecha As Date)
Dim matfilas() As Integer
Dim matcolumnas() As Integer
Dim filasimp() As Integer
Dim colsimp() As Integer
Dim vallim As Double
Dim valcons As Double
Dim varmd As Double
Dim fechacn  As Date
Dim vallim1 As Double
Dim vallim2 As Double
Dim vallim3 As Double
Dim varderiv1 As Double
Dim varderiv2 As Double
Dim varderiv3 As Double
Dim valcons1 As Double
Dim valcons2 As Double
Dim valcons3 As Double
Dim nofilas As Long
Dim nocols As Long
Dim i As Long
Dim mattabla() As Variant
Dim nodim1 As Long
Dim nodim2 As Long
Dim salida As Object
Dim contar As Long
Dim posy As Long

   CapitalNeto = DevLimitesVaR(CDate(fecha), MatCapitalSist, "CAPITAL NETO B") * 1000000
   fechacn = DevFechaLimite(CDate(fecha), MatCapitalSist, "CAPITAL NETO B")
   vallim1 = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV")
   vallim2 = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV EST")
   vallim3 = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV10")
   varderiv1 = LeerResVaRTabla(fecha, txtportCalc1, "DERIVADOS DE NEGOCIACION", "CVARH", 500, 1, 0, 0.03, 0)
   varderiv2 = LeerResVaRTabla(fecha, txtportCalc1, "DERIVADOS ESTRUCTURALES", "CVARH", 500, 1, 0, 0.03, 0)
   varderiv3 = LeerResVaRTabla(fecha, txtportCalc1, "DERIVADOS NEGOCIACION RECLASIFICACION", "CVARH", 500, 1, 0, 0.03, 0)
   valcons1 = -varderiv1 / (vallim1 * CapitalNeto)
   valcons2 = -varderiv2 / (vallim2 * CapitalNeto)
   valcons3 = -varderiv3 / (vallim3 * CapitalNeto)
   mattabla = GenRepVaRDeriv(fecha, vallim1, valcons1, vallim2, valcons2, vallim3, valcons3)
   nodim1 = UBound(mattabla, 1)
   nodim2 = UBound(mattabla, 2)
   nofilas = nodim1 + 1
   nocols = nodim2 + 1
   ReDim matfilas(1 To nofilas) As Integer
   ReDim matcolumnas(1 To nocols) As Integer
   matfilas(1) = 700
   For i = 2 To nofilas
       matfilas(i) = matfilas(i - 1) + 600
   Next i
   matcolumnas(1) = 1300
   matcolumnas(2) = 1300 + 3000
   For i = 3 To nocols
   matcolumnas(i) = matcolumnas(i - 1) + 1700
   Next i
'se definen las columnas y filas a imprimir
ReDim filasimp(1 To nofilas) As Integer, colsimp(1 To nocols) As Integer
For i = 1 To nofilas
 filasimp(i) = i
Next i



'SE deben buscar los consolidados de var de mesa de dinero que se generaron en el var
'global y pegarlos aqui
Set salida = Printer
contar = 2
salida.Orientation = 2
matfilas(1) = 2700
matfilas(2) = matfilas(1) + 300
Call ImpCabRepDeriv(fecha, salida)
Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, 1)
Do While contar <= nodim1
posy = matfilas(contar)
If matfilas(contar) > 11000 Then
   matfilas(contar) = 2700 + 300
   salida.NewPage
   Call ImpCabRepDeriv(fecha, salida)
   Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, 1)
End If
'se imprime el cuadro con los resultados del VaR
   salida.FontSize = 6
   salida.FontBold = False
   If contar <= nodim1 Then matfilas(contar + 1) = matfilas(contar) + 270
   Call ImprimirLineaReporte(mattabla, Printer, matfilas, matcolumnas, filasimp, colsimp, contar)
 contar = contar + 1
Loop
salida.EndDoc

End Sub

Function LeerResVaRTabla(ByVal fecha As Date, ByVal txtport As String, ByVal txtsubport As String, ByVal tvar As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Long, ByVal nconf As Double, ByVal lambda As Double)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
   txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & txtsubport & "'"
   txtfiltro2 = txtfiltro2 & " AND TVAR = '" & tvar & "'"
   txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc
   txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo
   txtfiltro2 = txtfiltro2 & " AND NOSIM = " & nosim
   txtfiltro2 = txtfiltro2 & " AND NCONF = " & nconf
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      LeerResVaRTabla = rmesa.Fields("VALOR")
      rmesa.Close
   Else
      LeerResVaRTabla = 0
   End If
End Function

Sub ImpCabRepDeriv(ByVal fecha As Date, ByRef salida As Object)
Dim texto1 As String
Dim largo1 As Long
Dim texto2 As String
Dim largo2 As Long
Dim texto3 As String
Dim largo3 As Long

'el logo de banobras
salida.PaintPicture frmCalVar.Picture1.Picture, 50, 50, 3500, 1000
salida.CurrentY = 300
salida.CurrentX = 7100
salida.FontSize = 9
salida.FontBold = True
salida.Print "DIRECCION GENERAL ADJUNTA DE ADMINISTRACION DE RIESGOS"
salida.CurrentX = 7100
salida.Print "DIRECCION DE ADMINISTRACION DE RIESGOS"
 texto1 = "VALOR EN RIESGO (CVAR) DE LA POSICION DE DERIVADOS"
 largo1 = Printer.TextWidth(texto1)
 salida.CurrentY = 1000
 salida.CurrentX = (Printer.ScaleWidth - largo1) / 2
 salida.Print texto1
 texto2 = "POR SWAP Y CONSOLIDADO DE LA POSICION DE RIESGO"
 largo2 = Printer.TextWidth(texto2)
 salida.CurrentX = (Printer.ScaleWidth - largo2) / 2
 salida.Print texto2
 texto3 = "(Cifras en pesos)"
 largo3 = Printer.TextWidth(texto3)
 salida.CurrentX = (Printer.ScaleWidth - largo3) / 2
 salida.Print texto3

salida.Print ""
salida.FontSize = 8
salida.Print "             Fecha de la Posición:" & fecha
salida.Print "             Fecha de Evaluación:" & fecha
salida.Print "             Capital Neto BANOBRAS:" & Format(CapitalNeto, "###,###,###,###,###,###,##0")
salida.Print "             No de escenarios históricos: 500"

End Sub

Sub ImpCabRepMDT(ByVal fecha As Date, ByRef salida As Object)
Dim texto1 As String
Dim texto2 As String
Dim texto3 As String
Dim largo1 As Integer
Dim largo2 As Integer
Dim largo3 As Integer

'el logo de banobras
salida.PaintPicture frmCalVar.Picture1.Picture, 50, 50, 3500, 1000
salida.CurrentY = 300
salida.CurrentX = 7100
salida.FontSize = 9
salida.FontBold = True
salida.Print "DIRECCION GENERAL ADJUNTA DE ADMINISTRACION DE RIESGOS"
salida.CurrentX = 7100
salida.Print "DIRECCION DE ADMINISTRACION DE RIESGOS"
 texto1 = "VALOR EN RIESGO (CVAR)"
 largo1 = Printer.TextWidth(texto1)
 salida.CurrentY = 1000
 salida.CurrentX = (Printer.ScaleWidth - largo1) / 2
 salida.Print texto1
 texto2 = "POR TIPO DE INSTRUMENTO Y CONSOLIDADO DE LA POSICION DE MERCADO DE DINERO"
 largo2 = Printer.TextWidth(texto2)
 salida.CurrentX = (Printer.ScaleWidth - largo2) / 2
 salida.Print texto2
 texto3 = "(Cifras en pesos)"
 largo3 = Printer.TextWidth(texto3)
 salida.CurrentX = (Printer.ScaleWidth - largo3) / 2
 salida.Print texto3

salida.Print ""
salida.FontSize = 8
salida.Print "             Fecha de la Posición:" & fecha
salida.Print "             Fecha de Evaluación:" & fecha
salida.Print "             Capital Neto BANOBRAS:" & Format(CapitalNeto, "###,###,###,###,###,###,##0")
salida.Print "             No de escenarios históricos: 500"

End Sub

Sub ImpCabRepMD(ByVal fecha As Date, ByRef salida As Object)
Dim texto1 As String
Dim texto2 As String
Dim texto3 As String
Dim largo1 As Integer
Dim largo2 As Integer
Dim largo3 As Integer

'el logo de banobras
salida.PaintPicture frmCalVar.Picture1.Picture, 50, 50, 3500, 1000
salida.CurrentY = 300
salida.CurrentX = 7100
salida.FontSize = 9
salida.FontBold = True
salida.Print "DIRECCION GENERAL ADJUNTA DE ADMINISTRACION DE RIESGOS"
salida.CurrentX = 7100
salida.Print "DIRECCION DE ADMINISTRACION DE RIESGOS"
 texto1 = "VALOR EN RIESGO (CVAR)"
 largo1 = Printer.TextWidth(texto1)
 salida.CurrentY = 1000
 salida.CurrentX = (Printer.ScaleWidth - largo1) / 2
 salida.Print texto1
 texto2 = "POR TIPO DE INSTRUMENTO Y CONSOLIDADO DE LA POSICION DE LA MESA DE DINERO"
 largo2 = Printer.TextWidth(texto2)
 salida.CurrentX = (Printer.ScaleWidth - largo2) / 2
 salida.Print texto2
 texto3 = "(Cifras en pesos)"
 largo3 = Printer.TextWidth(texto3)
 salida.CurrentX = (Printer.ScaleWidth - largo3) / 2
 salida.Print texto3

salida.Print ""
salida.FontSize = 8
salida.Print "             Fecha de la Posición:" & fecha
salida.Print "             Fecha de Evaluación:" & fecha
salida.Print "             Capital Neto BANOBRAS:" & Format(CapitalNeto, "###,###,###,###,###,###,##0")
salida.Print "             No de escenarios históricos: 500"

End Sub


Sub ImpCabRepTeso(ByVal fecha As Date, ByRef salida As Object)
Dim texto1 As String
Dim texto2 As String
Dim texto3 As String
Dim largo1 As Integer
Dim largo2 As Integer
Dim largo3 As Integer



'el logo de banobras
salida.PaintPicture frmCalVar.Picture1.Picture, 50, 50, 3500, 1000
salida.CurrentY = 300
salida.CurrentX = 7100
salida.FontSize = 9
salida.FontBold = True
salida.Print "DIRECCION GENERAL ADJUNTA DE ADMINISTRACION DE RIESGOS"
salida.CurrentX = 7100
salida.Print "DIRECCION DE ADMINISTRACION DE RIESGOS"
 texto1 = "VALOR EN RIESGO (CVAR)"
 largo1 = Printer.TextWidth(texto1)
 salida.CurrentY = 1000
 salida.CurrentX = (Printer.ScaleWidth - largo1) / 2
 salida.Print texto1
 texto2 = "POR TIPO DE INSTRUMENTO Y CONSOLIDADO DE LA POSICION DE LA TESORERIA"
 largo2 = Printer.TextWidth(texto2)
 salida.CurrentX = (Printer.ScaleWidth - largo2) / 2
 salida.Print texto2
 texto3 = "(Cifras en pesos)"
 largo3 = Printer.TextWidth(texto3)
 salida.CurrentX = (Printer.ScaleWidth - largo3) / 2
 salida.Print texto3

salida.Print ""
salida.FontSize = 8
salida.Print "             Fecha de la Posición:" & fecha
salida.Print "             Fecha de Evaluación:" & fecha
salida.Print "             Capital Neto BANOBRAS:" & Format(CapitalNeto, "###,###,###,###,###,###,##0")
salida.Print "             No de escenarios históricos: 500"

End Sub

Function DeterminaLimVaR(ByVal fecha As Date) As Double()
Dim i As Integer

CapitalNeto = DevLimitesVaR(fecha, MatCapitalSist, "CAPITAL NETO B") * 1000000
CapitalBase = DevLimitesVaR(fecha, MatCapitalSist, "CAPITAL BASE B") * 1000000
ReDim mata(1 To NoPortafolios, 1 To 2) As Double
mata(1, 1) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR CON")
mata(2, 1) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR MD")
mata(3, 1) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR MC")
mata(4, 1) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV")
mata(5, 1) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV EST")
mata(6, 1) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV10")
For i = 1 To NoPortafolios
    mata(i, 2) = mata(i, 1) * CapitalNeto
Next i
DeterminaLimVaR = mata
End Function


Function NomPosMD(ByVal opcion As Integer)
Dim txtfiltro As String

txtfiltro = TablaPosMD & " WHERE TIPOPOS = 2"
Select Case opcion
Case 1    'solo mesa de dinero
   txtfiltro = txtfiltro & " AND CPOSICION = 1"
Case 2    'tesoreria
   txtfiltro = txtfiltro & " AND CPOSICION = 2"
Case 4    'mercado de dinero
   txtfiltro = txtfiltro & " AND CPOSICION = 1 OR CPOSICION = 2"
End Select
NomPosMD = NombresPosicion(txtfiltro)
End Function

Function FechasPosDiv(ByVal opcpos As Integer) As Date()
Dim txtfiltro As String
If opcpos = 1 Then   'reales
   txtfiltro = TablaPosDiv & " WHERE TIPOPOS = 1"
ElseIf opcpos = 3 Then   'intradia
   txtfiltro = TablaPosDiv & " WHERE TIPOPOS = 3"
ElseIf opcpos = 5 Or opcpos = 6 Then 'reales + intradia
  txtfiltro = TablaPosDiv & " WHERE TIPOPOS = 1 OR TIPOPOS = 3"
End If
FechasPosDiv = FechasPosicion(txtfiltro)
End Function

Function NomPosDiv()
Dim txtfiltro As String
txtfiltro = TablaPosDiv & " WHERE TIPOPOS = 2"
NomPosDiv = NombresPosicion(txtfiltro)
End Function

Function NomPosSwaps()
Dim txtfiltro As String
txtfiltro = TablaPosSwaps
NomPosSwaps = NombresPosicion(txtfiltro)
End Function

Function FechasPosFwd(ByVal opcpos As Integer)
Dim txtfiltro As String
If opcpos = 1 Then
   txtfiltro = TablaPosFwd & " WHERE TIPOPOS = 1"    'real
ElseIf opcpos = 3 Then
   txtfiltro = TablaPosFwd & " WHERE TIPOPOS = 3"    'intradia
ElseIf opcpos = 5 Or opcpos = 6 Then
   txtfiltro = TablaPosFwd & " WHERE TIPOPOS = 1 OR TIPOPOS = 3"    'REAL + INTRADIA
End If
FechasPosFwd = FechasPosicion(txtfiltro)
End Function

Function NomPosFwd()
Dim txtfiltro As String
  txtfiltro = TablaPosFwd & " WHERE TIPOPOS ='2'"
  NomPosFwd = NombresPosicion(txtfiltro)
End Function

Function FechasPosicion(ByVal txtcadena As String) As Date()
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

  txtfiltro = "select count(DISTINCT fecha) from " & txtcadena
  txtfiltro1 = "select fecha from " & txtcadena & " GROUP BY fecha order by fecha"
  rmesa.Open txtfiltro, ConAdo
  noreg = rmesa.Fields(0)
  rmesa.Close
  If noreg <> 0 Then
  rmesa.Open txtfiltro1, ConAdo
  ReDim matff(1 To noreg, 1 To 1) As Date
   For i = 1 To noreg
   If IsDate(rmesa.Fields(0)) Then
    matff(i, 1) = CDate(rmesa.Fields(0))
   Else
    matff(i, 1) = rmesa.Fields(0)
   End If
   rmesa.MoveNext
   Next i
 rmesa.Close
 matff = ROrdenF(matff, 1)
Else
 ReDim matff(0 To 0, 0 To 0) As Date
End If
FechasPosicion = matff
End Function

Function NombresPosicion(ByVal txtcadena As String)
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

  txtfiltro = "select count(DISTINCT NOMPOS) from " & txtcadena
  txtfiltro1 = "select NOMPOS from " & txtcadena & " GROUP BY NOMPOS order by NOMPOS"
  rmesa.Open txtfiltro, ConAdo
  noreg = rmesa.Fields(0)
  rmesa.Close
  If noreg <> 0 Then
  rmesa.Open txtfiltro1, ConAdo
  ReDim matff(1 To noreg, 1 To 1) As Variant
   For i = 1 To noreg
    matff(i, 1) = rmesa.Fields("NOMPOS")
    rmesa.MoveNext
   Next i
 rmesa.Close
 matff = RutinaOrden(matff, 1, SRutOrden)
Else
 ReDim matff(0 To 0, 0 To 0) As Variant
End If
NombresPosicion = matff
End Function

Function VerifVPrecios1(ByRef matvecp() As Variant, ByVal tfecha As Integer)
Dim noreg As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim fvencimiento As Date
Dim femision As Date
Dim f_val As Date
Dim pemision As Long
Dim tm As String
Dim fvaltxt As String
Dim tv As String
Dim emision As String
Dim serie As String
Dim psucio As Double
Dim plimpio As Double
Dim intmd As Double
Dim sobret As Double
Dim dxv As Long
Dim vn As Double
Dim fventxt As String
Dim femisiontxt As String
Dim tcupon As Double
Dim moneda As Integer
Dim pc As Long
Dim tasa As Double

noreg = UBound(matvecp, 1)
ReDim mata(1 To noreg, 1 To 32) As Variant
If noreg <> 0 Then
'se procede a verificar todos los datos t a agregare algunos datos
'del vector

For i = 1 To noreg
    fvencimiento = 0
    femision = 0
    f_val = 0
    pemision = 0
    moneda = 0
    tm = matvecp(i, 2)
    fvaltxt = matvecp(i, 3)
    tv = Trim(UCase(matvecp(i, 4)))
    emision = UCase(Trim(matvecp(i, 5)))
    serie = matvecp(i, 6)
    psucio = Val(matvecp(i, 7))
    plimpio = Val(matvecp(i, 8))
    intmd = Val(matvecp(i, 9))
    sobret = Val(matvecp(i, 10))
    dxv = Val(matvecp(i, 11))
    vn = Val(matvecp(i, 12))
    fventxt = matvecp(i, 13)
femisiontxt = matvecp(i, 14)
moneda = Val(matvecp(i, 15))
tcupon = Val(matvecp(i, 16))
pc = Val(matvecp(i, 17))
f_val = determfecha(fvaltxt, tfecha)
fvencimiento = determfecha(fventxt, tfecha)
femision = determfecha(femisiontxt, tfecha)


If IsDate(fvencimiento) And IsDate(femision) Then pemision = fvencimiento - femision
If dxv <> (fvencimiento - f_val) And fvencimiento <> 0 And f_val <> 0 And Val(fvencimiento - f_val) > 0 Then
   dxv = Val(fvencimiento - f_val)
End If
'se procede a corregir algunos nombres debido a que
'al hacer la union de estos para crear la clave de emision
'no se obtiene un nombre de emision valido
   serie = CorreccionSerie(tm, tv, emision, serie)

'estos valores hay que ponerlos en forma de porcentaje
'se procede a realizar los cambios necesarios para este registro
  mata(i, 1) = f_val
  mata(i, 2) = tm
  mata(i, 3) = tv
  mata(i, 4) = emision
  mata(i, 5) = serie
  mata(i, 6) = psucio
  mata(i, 7) = plimpio
  mata(i, 8) = intmd
  mata(i, 9) = sobret
  mata(i, 10) = dxv
  mata(i, 11) = vn
  mata(i, 12) = fvencimiento
  mata(i, 13) = femision
  mata(i, 14) = moneda
  mata(i, 15) = tcupon
  mata(i, 16) = pc
 AvanceProc = i / noreg
 MensajeProc = "Verificando y depurando el vector de precios del " & f_val & ": " & Format(AvanceProc, "#0.00 %")
 DoEvents
Next i
VerifVPrecios1 = mata
End If

End Function

Function CorreccionSerie(ByVal tm As String, ByVal tv As String, ByVal emision As String, ByVal serie As String)
If tm = "ME" And tv = "D1" And Len(serie) = 5 And Left(emision, 3) = "MEX" Then
   CorreccionSerie = "0" & serie
ElseIf emision = "BONOS" And Len(serie) = 5 Then
   CorreccionSerie = "0" & serie
ElseIf emision = "UDIBONO" And Len(serie) = 5 Then
   CorreccionSerie = "0" & serie
ElseIf tv = "F" And emision = "BANOBRA" And Len(serie) = 4 Then
   CorreccionSerie = "0" & serie
ElseIf (Left(tv, 2) = "LT" Or Left(tv, 2) = "XA" Or Left(tv, 2) = "IT" Or Left(tv, 2) = "IP") And Len(serie) = 5 Then
   CorreccionSerie = "0" & serie
ElseIf Left(tv, 1) = "I" And Len(serie) = 4 Then
   CorreccionSerie = "0" & serie
ElseIf emision = "CETES" And tv = "BI" And Len(serie) = 5 Then
   CorreccionSerie = "0" & serie
Else
    CorreccionSerie = serie
End If
End Function

Function determfecha(txtfecha As String, ByVal tfecha As Integer)
   If Format(Val(txtfecha), "00000000") = txtfecha Then
      If tfecha = 0 Then
         determfecha = CDate(Mid(txtfecha, 7, 2) & "/" & Mid(txtfecha, 5, 2) & "/" & Mid(txtfecha, 1, 4))
      Else
         determfecha = CDate(Mid(txtfecha, 5, 2) & "/" & Mid(txtfecha, 7, 2) & "/" & Mid(txtfecha, 1, 4))
      End If
   Else
      determfecha = 0
   End If
End Function



Sub GuardaDatosVAR(ByVal fecha As Date, ByVal txtport As String, ByVal tipotitulos As String, ByVal txttvar As String, ByVal dvol As Long, ByVal horiz As Long, ByVal nconf As Double, ByVal liminf As Double, ByVal limsup As Double)
Dim txtfecha As String
Dim txtcadena As String
'rutina para guardar los datos que se generan por el
'calculo del var
'el indicador nocol indica en que columna de MatResumenVART se
'Call InConexOracle("alm2", conAdo)
 txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtcadena = "DELETE FROM " & TablaResVaR & " WHERE FECHA = " & txtfecha & " AND MPOSICION = '" & txtport & "' AND PORTAFOLIO = '" & tipotitulos & "' and TVAR = '" & txttvar & "' AND DIAS_VOL = " & dvol & " AND HORIZ = " & horiz
 txtcadena = txtcadena & " AND NIV_CONF = " & nconf
 ConAdo.Execute txtcadena
 txtcadena = "INSERT INTO " & TablaResVaR & " VALUES("
 txtcadena = txtcadena & txtfecha & ","
 txtcadena = txtcadena & "'" & txtport & "',"
 txtcadena = txtcadena & "'" & tipotitulos & "',"
 txtcadena = txtcadena & "'" & txttvar & "',"
 txtcadena = txtcadena & dvol & ","
 txtcadena = txtcadena & horiz & ","
 txtcadena = txtcadena & nconf & ","
 txtcadena = txtcadena & liminf & ","
 txtcadena = txtcadena & limsup & ")"
 ConAdo.Execute txtcadena
 DoEvents
End Sub

Sub GenerarPosRiesgo(ByVal txtbase As String, ByVal fecha As Date)
Dim base1 As DAO.Database
Dim base2 As DAO.Database
Dim registros1 As DAO.recordset
Dim registros2 As DAO.recordset
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'mesa de derivados

Set base1 = OpenDatabase(DirBases & "\" & txtbase, dbDriverNoPrompt, False, ";Pwd=" & ContraseñaCatalogos)
Set base2 = OpenDatabase(DirBases & "\" & txtbase, dbDriverNoPrompt, False, ";Pwd=" & ContraseñaCatalogos)
Set registros1 = base1.OpenRecordset("SELECT  * FROM [posicion global] WHERE [fecha registro] = " & CLng(fecha) & " and [MESA O POSICION] = '3' ORDER BY [clave de la emisión]", dbOpenDynaset, dbReadOnly)
Set registros2 = base1.OpenRecordset("SELECT  * FROM [POSICION RIESGO]", dbOpenDynaset)
If registros1.RecordCount <> 0 Then
base2.Execute "DELETE FROM [POSICION RIESGO] where [fecha posicion] = " & CLng(fecha) & " AND [MESA O POSICION] = '3'"
registros1.MoveLast
noreg = registros1.RecordCount
registros1.MoveFirst
nocampos = registros1.Fields.Count
For i = 1 To noreg
registros2.AddNew
Call GrabarTAccess(registros2, 0, fecha, i)
 For j = 1 To nocampos
  Call GrabarTAccess(registros2, j, LeerTAccess(registros1, j - 1, i), i)
 Next j
registros2.Update
registros1.MoveNext
AvanceProc = i / noreg
MensajeProc = "Generando la posicion de derivados del " & fecha & " " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i
End If
registros2.Close
registros1.Close
base2.Close
base1.Close
Call MostrarMensajeSistema("Se genero la posicion de la mesa de derivados del " & fecha, frmProgreso.Label2, 1, Date, Time, NomUsuario)
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function LeerParamRE(ByVal fecha As Date) As Variant()
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim fecha1 As Date
Dim fechax As Date
Dim matb() As Variant
Dim matc() As Double
Dim noreg As Integer
Dim j As Integer
Dim txtfechax As String
Dim rmesa As New ADODB.recordset

'esta rutina actualiza los parametros del resumen ejecutivo
'incluyendo los limites de VaR y capital
'1 valor udi
'2 dolar fix
'3 dolar 24
'4 dolar 48
ReDim mata(1 To 1, 1 To 16) As Variant
mata(1, 1) = fecha
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "select * from " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha & " AND CONCEPTO ='UDI'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   mata(1, 2) = rmesa.Fields("VALOR")   'udi
   rmesa.Close
Else
   mata(1, 2) = 0
End If
txtfiltro = "select * from " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha & " AND CONCEPTO ='DOLAR PIP FIX'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   mata(1, 3) = rmesa.Fields("VALOR")  'dolar fix
   rmesa.Close
Else
   mata(1, 3) = 0
End If
txtfiltro = "select * from " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha & " AND CONCEPTO ='YEN PIP'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   mata(1, 4) = rmesa.Fields("VALOR")  'yen
   rmesa.Close
Else
   mata(1, 4) = 0
End If
txtfiltro = "select * from " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha & " AND CONCEPTO ='EURO PIP'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   mata(1, 5) = rmesa.Fields("VALOR")  'euro
   rmesa.Close
Else
   mata(1, 5) = 0
End If
mata(1, 6) = DevLimitesVaR(fecha, MatCapitalSist, "CAPITAL NETO B")  'capital neto
'se necesita el capital tabla de 3 meses antes
matb = LeerDLimiteBanxico(fecha)
mata(1, 7) = matb(1)
mata(1, 8) = matb(2)
mata(1, 9) = matb(3)
mata(1, 10) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR CON")
mata(1, 11) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR MD")
mata(1, 12) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR MC")
mata(1, 13) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV")
mata(1, 14) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV EST")
mata(1, 15) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV10")
mata(1, 16) = CDate(DevFechaLimite(fecha, MatCapitalSist, "CAPITAL NETO B"))
LeerParamRE = mata

End Function

Function LeerDLimiteBanxico(ByVal fecha As Date)
Dim fechax As Date
Dim fecha1 As Date
Dim fecha2 As Date
Dim fecha3 As Date
Dim txtfechax As String
Dim rmesa As New ADODB.recordset
Dim noreg As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim mata(1 To 3) As Variant
  
   fecha1 = DateSerial(Year(fecha), Month(fecha), 1)
   fecha2 = DateAdd("m", -2, fecha1) - 1
   fecha3 = PBD(fecha2, "MX")
   mata(1) = DevLimitesVaR(fecha3, MatCapitalSist, "CAP BASICO CONV")
   txtfechax = "to_date('" & Format(fecha3, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro2 = "select * from " & TablaFRiesgoO & " WHERE FECHA = " & txtfechax & " AND CONCEPTO ='DOLAR PIP FIX'"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      mata(2) = rmesa.Fields("VALOR")  'dolar fix
      rmesa.Close
   Else
      mata(2) = 0
   End If
   mata(3) = CLng(fecha3)
   LeerDLimiteBanxico = mata
End Function


Function CalcularMaxCapBasico(fecha)

Dim mes As Integer
Dim anio As Integer
Dim noreg As Integer
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
'determinar las fechas a leer del capital

mes = Month(fecha)
anio = Year(fecha)



txtfiltro2 = "SELECT * FROM " & PrefijoBD & TablaLimites & " WHERE CONCEPTO = 'CAPITAL NETO CONV'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"

For i = 1 To noreg
Next i

For i = 1 To 12

Next i
CalcularMaxCapBasico = 1
End Function

Function LeerCapBasico3M(ByVal fecha As Date)
Dim mes As Integer
Dim anio As Integer
Dim fecha1 As Date
Dim fecha2 As Date
Dim fecha3  As Date

   mes = Month(fecha)
   anio = Year(fecha)
   fecha1 = DateSerial(anio, mes, 1)
   fecha2 = DateAdd("m", -2, fecha1) - 1
   fecha3 = PBD(fecha2, "MX")
   LeerCapBasico3M = DevLimitesVaR(fecha3, MatCapitalSist, "CAP BASICO CONV")
End Function


Sub GuardaParamRE(mata, txttabla, conex, rbase)
Dim noreg As Integer
Dim noreg1 As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim txtfiltro As String
Dim txtcadena As String

'se abre el resumen ejecutivo
noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
If noreg > 0 Then
   For i = 1 To UBound(mata, 1)
   txtfiltro = "SELECT COUNT(*) FROM [" & txttabla & "] WHERE FECHA = " & CLng(mata(i, 1))
   rbase.Open txtfiltro, conex
   noreg1 = rbase.Fields(0)
   rbase.Close
   If noreg1 <> 0 Then
   
   Else
      txtcadena = "INSERT INTO[" & txttabla & "] VALUES("
      txtcadena = txtcadena & CLng(mata(i, 1)) & ","
      For j = 2 To nocampos - 1
      txtcadena = txtcadena & mata(i, j) & ","
      Next j
      txtcadena = txtcadena & CLng(mata(i, nocampos)) & ")"
      conex.Execute txtcadena
   End If
   Next i
End If
End Sub

Sub GuardaDesgloseDiv(ByVal fecha As Date, ByRef matd() As Variant)
Dim nodivisas As Long
Dim i As Long
Dim r As Long
Dim txtcadena As String
Dim txtfecha As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta rutina guarda la posicion global por
'tipo de divisa en una tabla de datos
'esta rutina solo se corre cuando se ha leido la posicion global
If Not EsArrayVacio(matd) Then
nodivisas = UBound(matd, 1)
'se verfica que la matriz tenga datos
If IsArray(matd) And UBound(matd, 1) >= 1 And UBound(matd, 2) >= 1 Then
 txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 ConAdo.Execute "DELETE FROM " & TablaPosDivCon & " WHERE FECHA = " & txtfecha
 For i = 1 To NoGruposPort
 For r = 1 To nodivisas
 If matd(r, i, 1) <> 0 Or matd(r, i, 2) <> 0 Then
  txtcadena = "INSERT INTO " & TablaPosDivCon & " VALUES("
  txtcadena = txtcadena & txtfecha & ","
  txtcadena = txtcadena & "'" & txtportBanobras & "',"
  txtcadena = txtcadena & "'" & MatGruposPortPos(i, 3) & "',"
  txtcadena = txtcadena & r & ","
  If Not EsVariableVacia(matd(r, i, 1)) Then
   txtcadena = txtcadena & matd(r, i, 1) & ","
  Else
   txtcadena = txtcadena & "null,"
  End If
  If Not EsVariableVacia(matd(r, i, 2)) Then
   txtcadena = txtcadena & matd(r, i, 2) & ")"
  Else
   txtcadena = txtcadena & "null)"
  End If
  ConAdo.Execute txtcadena
 End If
 Next r
 Call MostrarMensajeSistema("Guardando la posicion por divisa " & Format(i / NoGruposPort, "###0.00"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
 DoEvents
 Next i
End If
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function LeerDesgloseDiv(ByVal fecha As Date, ByRef valpos As Double)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtport As String
Dim i As Integer
Dim j As Integer
Dim noreg As Long
Dim rmesa As New ADODB.recordset

txtport = txtportCalc1
Dim matport(1 To 4) As String
Dim mata(1 To 3, 1 To 4) As Variant
matport(1) = "MC POS DOLARES"
matport(2) = "MC POS EUROS"
matport(3) = "MC POS YENES"
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
valpos = 0
For i = 1 To 3
    mata(i, 1) = "F" & CLng(fecha) & "P" & i + 2
    mata(i, 2) = CLng(fecha)
    mata(i, 3) = i + 2
    txtfiltro2 = "select * from " & TablaValPosPort & " WHERE FECHAP = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND FECHAFR = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND FECHAV = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO ='" & txtport & "'"
    txtfiltro2 = txtfiltro2 & " AND SUBPORT ='" & matport(i) & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       mata(i, 4) = rmesa.Fields(7)
       valpos = valpos + mata(i, 4)
       rmesa.Close
    Else
       mata(i, 4) = 0
    End If
Next i
LeerDesgloseDiv = mata
End Function

Function LeerPosDivTot(ByVal fecha As Date) As Variant()
On Error GoTo hayerror:
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim valdolar As Double
Dim valeuro As Double
Dim valyen As Double
Dim sumadolar As Double
Dim sumaeuro As Double
Dim sumayen As Double
Dim signo As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro1 = "select COUNT(*) from " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtfiltro2 = "select * from " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   For i = 1 To noreg
       For j = 1 To nocampos
           mata(i, j) = rmesa.Fields(j - 1)
       Next j
       rmesa.MoveNext
   Next i
   rmesa.Close
   'se buscan los tipos de cambio correspondientes a esta fecha
   txtfiltro1 = "select * from " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = 'DOLAR PIP FIX'"
   rmesa.Open txtfiltro1, ConAdo
   If rmesa.RecordCount <> 0 Then
     valdolar = rmesa.Fields(3)
   Else
     valdolar = 0
   End If
   rmesa.Close
   txtfiltro1 = "select * from " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = 'EURO PIP'"
   rmesa.Open txtfiltro1, ConAdo
   If rmesa.RecordCount <> 0 Then
     valeuro = rmesa.Fields(3)
   Else
     valeuro = 0
   End If
   rmesa.Close
   txtfiltro1 = "select * from " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = 'YEN PIP'"
   rmesa.Open txtfiltro1, ConAdo
   If rmesa.RecordCount <> 0 Then
     valyen = rmesa.Fields(3)
   Else
     valyen = 0
   End If
   rmesa.Close
   ReDim matb(1 To 3, 1 To 4) As Variant
   matb(1, 1) = "F" & CLng(fecha) & "P3"
   matb(2, 1) = "F" & CLng(fecha) & "P4"
   matb(3, 1) = "F" & CLng(fecha) & "P5"
   matb(1, 2) = CLng(fecha)
   matb(2, 2) = CLng(fecha)
   matb(3, 2) = CLng(fecha)
   matb(1, 3) = 3
   matb(2, 3) = 4
   matb(3, 3) = 5
   sumadolar = 0
   sumaeuro = 0
   sumayen = 0
   For i = 1 To noreg
       If mata(i, 8) = 1 Then
          signo = 1
       Else
          signo = -1
       End If
       If mata(i, 12) = "*CMXPUSDFIX" Then sumadolar = sumadolar + signo * mata(i, 13)
       If mata(i, 12) = "*CMXPEUREUR" Then sumaeuro = sumaeuro + signo * mata(i, 13)
       If mata(i, 12) = "*CMXPJPYJPY" Then sumayen = sumayen + signo * mata(i, 13)
   Next i
   matb(1, 4) = sumadolar * valdolar
   matb(2, 4) = sumaeuro * valeuro
   matb(3, 4) = sumayen * valyen
Else
   ReDim matb(0 To 0, 0 To 0) As Variant
End If
LeerPosDivTot = matb
On Error GoTo 0
Exit Function
hayerror:
MsgBox "leerposdivtot" & error(Err())
End Function

Sub ExpDesgloseDiv(ByVal fecha As Date, ByRef mata() As Variant, ByVal nomtabla As String, conex, rbase)
Dim txtcadena As String
Dim noreg As Integer
Dim nocampos As Integer
Dim i As Integer

'se abre el resumen ejecutivo
noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
If noreg <> 0 Then
For i = 1 To noreg
txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
txtcadena = txtcadena & "'" & mata(i, 1) & "',"
txtcadena = txtcadena & CLng(mata(i, 2)) & ","
txtcadena = txtcadena & "'" & mata(i, 3) & "',"
txtcadena = txtcadena & mata(i, 4) & ")"
conex.Execute txtcadena
Next i
Else
 MsgBox "Faltan registros de la posición de divisas para esta fecha"
End If
End Sub

Sub CalcularBacktesting(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef exito As Boolean)

Dim txtfecha As String
Dim txtfecha2 As String
Dim txtborra As String
Dim txtinserta As String
Dim exito1 As Boolean
Dim sicontinuar As Boolean
Dim fecha As Date
Dim fechaev2 As Date
Dim mtasasv0() As Double
Dim mtasasv1() As Double
Dim matprecios2() As New resValIns
Dim noreg As Integer
Dim i As Integer
Dim indicet0 As Integer
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim parval As ParamValPos
Dim mrvalflujo() As resValFlujo
Dim exito2 As Boolean
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim exito3 As Boolean
Dim exito5 As Boolean
Dim txtmsg3 As String


'rutina para calcular el backtesting de una posicion
'se obtiene la posicion, se valua con los valores de mercado del dia
'de ayer y los nuevo y se obtiene la valuacion de
'se dimensiona una matriz para hacer un desglose del backtesting
'por grupos de titulos
'se estructura una tabla con la estructura del portafolio
fecha = fecha1
Do While fecha <= fecha2
'primero se busca la fecha en la historia de factores de riesgo
   indicet0 = 0
   indicet0 = BuscarValorArray(fecha, MatFechasVaR, 1)
   If indicet0 <> 0 Then
' POSICION:  ayer
' TASAS:     ayer
' VALUACION: ayer
      mattxt = CrearFiltroPosPort(fecha, txtport)
      Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito1)
      noreg = UBound(matpos, 1)
      If noreg <> 0 Then
         Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
         If exito2 Then
            ReDim matres(1 To noreg, 1 To 3) As Double
            mtasasv0 = CargaFR1Dia(fecha, exito1)
            Call RutinaCargaFR(fecha, exito1)
            MatCurvasT = LeerCurvaCompleta(fecha, exito5)
            Call AnexarDatosVPrecios(fecha, matposmd)
            Set parval = DeterminaPerfilVal("BACKTESTING")
            MatPrecios = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, mtasasv0, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
            For i = 1 To noreg
                matres(i, 1) = MatPrecios(i).mtm_sucio
            Next i
   'se busca los datos para el día siguiente
            If indicet0 + 1 <= UBound(MatFechasVaR, 1) Then
               fechaev2 = MatFechasVaR(indicet0 + 1, 1)
               mtasasv1 = CargaFR1Dia(fechaev2, exito1)
               MatCurvasT = LeerCurvaCompleta(fechaev2, exito1)
                  '        posicion:  ayer
      '        tasas:     hoy
      '        valuacion: ayer
               'Call AnexarDatosVPrecios( fecha,matposmd)
               Set parval = DeterminaPerfilVal("BACKTESTING")
               matprecios2 = CalcValuacion(fecha, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, mtasasv1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
               For i = 1 To noreg
                   matres(i, 2) = matprecios2(i).mtm_sucio
                   matres(i, 3) = matres(i, 2) - matres(i, 1)
               Next i
               txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
               txtborra = "DELETE FROM " & TablaResBack & " WHERE FECHA = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "'"
               ConAdo.Execute txtborra
               For i = 1 To noreg
                   txtfecha2 = "to_date('" & Format(matpos(i).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
                   txtinserta = "INSERT INTO " & TablaResBack & " VALUES("
                   txtinserta = txtinserta & txtfecha & ","                              'fecha
                   txtinserta = txtinserta & "'" & txtport & "',"                        'nombre del portafolio
                   txtinserta = txtinserta & matpos(i).C_Posicion & ","           'clave de la posicion
                   txtinserta = txtinserta & txtfecha2 & ","                             'fecha de registro
                   txtinserta = txtinserta & "'" & matpos(i).c_operacion & "',"   'clave de operacion
                   txtinserta = txtinserta & matres(i, 1) & ","                          'valuacion t0
                   txtinserta = txtinserta & matres(i, 2) & ","                          'valuacion t1
                   txtinserta = txtinserta & matres(i, 3) & ")"                          'diferencia
                   ConAdo.Execute txtinserta
                   AvanceProc = i / noreg
                   MensajeProc = "Guardando los resultados del calculo de backtesting " & Format(AvanceProc, "##0.00 %")
                   DoEvents
               Next i
            End If
            exito = True
            txtmsg = "El proceso finalizo correctamente"
         Else
            exito = False
            txtmsg = txtmsg2
         End If
      Else
         exito = False
         txtmsg = "No ha registros de la posicion"
      End If
   End If
fecha = fecha + 1
Loop
On Error GoTo 0
Exit Sub
hayerror:
MsgBox error(Err())
If Err() = 3113 Then
   Call ReiniciarConexOracleP(ConAdo)
   exito = False
End If
On Error GoTo 0
End Sub

Sub CalcResBackPortPos(ByVal fecha As Date, ByVal txtport As String, ByVal txtsubport As String)
Dim noreg As Integer
Dim i As Integer
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim suma1 As Double
Dim suma2 As Double
Dim suma3 As Double
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim txtborra As String
Dim txtinserta As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaResBack & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND (CPOSICION, FECHAREG, COPERACION) IN"
txtfiltro2 = txtfiltro2 & " (SELECT CPOSICION,FECHAREG,COPERACION FROM " & TablaPortPosicion & " "
txtfiltro2 = txtfiltro2 & " WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtsubport & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   suma1 = 0
   suma2 = 0
   suma3 = 0
   For i = 1 To noreg
       valor1 = rmesa.Fields("VALOR1")
       valor2 = rmesa.Fields("VALOR2")
       valor3 = rmesa.Fields("DIFERENCIA")
       suma1 = suma1 + valor1
       suma2 = suma2 + valor2
       suma3 = suma3 + valor3
       rmesa.MoveNext
   Next i
   rmesa.Close
   txtborra = "DELETE FROM " & TablaBackPort & " WHERE FECHA = " & txtfecha
   txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
   txtborra = txtborra & " AND SUBPORT = '" & txtsubport & "'"
   ConAdo.Execute txtborra
   txtinserta = "INSERT INTO " & TablaBackPort & " VALUES("
   txtinserta = txtinserta & txtfecha & ","
   txtinserta = txtinserta & "'" & txtport & "',"
   txtinserta = txtinserta & "'" & txtsubport & "',"
   txtinserta = txtinserta & suma1 & ","
   txtinserta = txtinserta & suma2 & ","
   txtinserta = txtinserta & suma3 & ")"
   ConAdo.Execute txtinserta
End If
End Sub

Sub AnexarDatosVPrecios(ByVal f_val As Date, ByRef matposmd() As propPosMD)
Dim fechax As Date
   If f_val <> FechaVPrecios Then
      MatVPreciosT = LeerPVPrecios(f_val)
   End If
   Call CompletarPosMesaD(matposmd, MatVPreciosT)
End Sub

Function FiltrarPosG1(ByRef Pos() As propPosRiesgo, ByRef matb() As Variant)
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim j As Long
Dim kk As Long
'esta funcion filtra la posicion en funcion de una matriz matb
noreg = UBound(Pos, 1)
contar = 0
ReDim mata(1 To 1) As New propPosRiesgo
For j = 1 To UBound(matb, 1)
    For i = 1 To noreg
        If Pos(i).c_operacion = matb(j, 1) Then
           contar = contar + 1
             ReDim Preserve mata(1 To contar) As New propPosRiesgo
             Set mata(contar) = Pos(i)
           End If
    Next i
Next j
If contar = 0 Then
  ReDim mata(0 To 0) As New propPosRiesgo
End If
FiltrarPosG1 = mata
End Function


Function LeerHistCurvas()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta funcion lee la historia de la estructura de la tabla de curvas del proveedor
'====================================================
txtfiltro1 = "select count(*) from " & PrefijoBD & TablaHistCurvas
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 5) As Variant
txtfiltro2 = "select * from " & PrefijoBD & TablaHistCurvas & " ORDER BY NOMBRE, FECHA"
rmesa.Open txtfiltro2, ConAdo
rmesa.MoveFirst
For i = 1 To noreg
 mata(i, 1) = rmesa.Fields("ID_CURVA")
 mata(i, 2) = rmesa.Fields("CURVA")
 mata(i, 3) = rmesa.Fields("FECHA")
 mata(i, 4) = rmesa.Fields("NOCOLUMNA")
 mata(i, 5) = rmesa.Fields("NOMBRE")
 rmesa.MoveNext
 AvanceProc = i / noreg
 MensajeProc = "Cargando la historia de las curvas del Proveedor de precios " & Format(AvanceProc, "##0.00 %")
Next i
rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerHistCurvas = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function LeerTCCurvasO()
'esta funcion lee la historia de la estructura de la tabla de curvas del proveedor
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

'====================================================
txtfiltro1 = "select * from " & PrefijoBD & TablaCatCurvas & " ORDER BY ID_CURVA, CURVA"
txtfiltro2 = "select COUNT(*) from (" & txtfiltro1 & ")"
rmesa.Open txtfiltro2, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 2) As Variant
rmesa.Open txtfiltro1, ConAdo
rmesa.MoveFirst
For i = 1 To noreg
    mata(i, 1) = rmesa.Fields("ID_CURVA")
    mata(i, 2) = rmesa.Fields("CURVA")
    'mata(i, 3) = RMesa.Fields("DESCRIPCION")
    rmesa.MoveNext
    AvanceProc = i / noreg
    MensajeProc = "Cargando el catalogo de curvas " & Format(AvanceProc, "##0.00 %")
Next i
rmesa.Close
Else
 ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerTCCurvasO = mata
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Function BuscarFechaCupon(ByVal fecha As Date, ByRef mata() As Variant, ByVal pc As Integer) As Date
Dim noreg As Long
Dim i As Long
Dim fechai As Date

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
noreg = UBound(mata, 1)
For i = 2 To noreg
If fecha >= mata(i - 1, 1) And fecha < mata(i, 1) Then
 fechai = mata(i - 1, 1)
 Exit For
End If
Next i
BuscarFechaCupon = fechai
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Function

Sub SubirArchFTP(ByVal nomarch1 As String, ByVal nomarch2 As String, ByVal nhost As String, ByVal usuario As String, ByVal pass As String)
Dim txtclave$
Dim salida As Variant
Dim exitoarch As Boolean
 Call VerificarSalidaArchivo(DirArchBat & "\script1.txt", 1, exitoarch)
 If exitoarch Then
 Print #1, usuario
 Print #1, pass
 Print #1, "put " & nomarch1 & " " & nomarch2
 Print #1, "quit"
 Print #1, "exit"
 Close #1
 txtclave$ = "ftp -s:" & DirArchBat & "\script1.txt " & nhost
 salida = ExecCmd(txtclave$)
 End If
End Sub

Sub BajarArchSFTP(ByVal nomarch1 As String, ByVal nomarch2 As String, ByVal nhost As String, ByVal usuario As String, ByVal pass As String)
Dim txtclave$
Dim salida As Variant
Dim exitoarch As Boolean
Dim txtnomarch As String
txtnomarch = DirArchBat & "\script1.txt"

 Call VerificarSalidaArchivo(txtnomarch, 1, exitoarch)
 If exitoarch Then
    Print #1, "open sftp://" & usuario & ":" & pass & "@" & nhost & "/"
    Print #1, "get " & nomarch1 & " " & nomarch2
    Print #1, "close"
    Print #1, "exit"
    Close #1
    txtclave$ = "c:\Program Files (x86)\WinSCP\WinSCP.com /script=" & Chr(34) & txtnomarch & Chr(34)
    salida = ExecCmd(txtclave$)
 End If
End Sub

Sub SubirArchSFTP(ByVal nomarch1 As String, ByVal nomarch2 As String, ByVal nhost As String, ByVal usuario As String, ByVal pass As String)
Dim txtclave$
Dim salida As Variant
Dim exitoarch As Boolean

 Call VerificarSalidaArchivo(DirArchBat & "\script1.txt", 1, exitoarch)
 If exitoarch Then
 Print #1, "put " & nomarch1 & " " & nomarch2
 Print #1, "quit"
 Print #1, "exit"
 Close #1
 txtclave$ = "psftp " & nhost & " -l " & usuario & " -pw " & pass & " -b " & DirArchBat & "\script1.txt "
 salida = ExecCmd(txtclave$)
 End If
End Sub

Sub TransferirVect1(ByVal fecha As Date, ByVal prefijo As String, ByVal sufijo As String, ByVal ext As String, ByVal direc1 As String, ByVal direc2 As String, ByVal direc3 As String, ByVal siacc1 As Boolean, ByVal siacc2 As Boolean, ByVal txtarchbatch As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim nomarch1 As String
Dim nomarch2 As String
Dim nomarch3 As String
Dim nomarch4 As String
Dim nomarch5 As String
Dim sihayarch1 As Boolean
Dim salida As Variant
Dim exitoarch As Boolean

exito = False
txtfecha = Format(fecha, "yyyymmdd")
nomarch1 = prefijo & txtfecha & sufijo & ".zip"
nomarch2 = direc1 & "\" & prefijo & txtfecha & sufijo & ".zip"
nomarch3 = direc1 & "\" & prefijo & txtfecha & sufijo & "." & ext
nomarch4 = direc2 & "\" & prefijo & txtfecha & sufijo & "." & ext
nomarch5 = direc3 & "\" & prefijo & txtfecha & sufijo & ".zip"
If siacc1 Then
   Call BajarArchSFTP(nomarch1, Chr(34) & nomarch2 & Chr(34), NomSRVPIP, usersftpPIP, passsftpPIP)
End If
sihayarch1 = VerifAccesoArch(nomarch2)
If sihayarch1 = True And siacc2 Then
If TamañoArch(nomarch2) <> 0 Then
'se crea el archivo que realiza los ListaProcesos en bat
 Call VerificarSalidaArchivo(DirArchBat & "\" & txtarchbatch, 1, exitoarch)
 If exitoarch Then
    'Print #1, "echo off"
     Print #1, Chr(34) & DirWinRAR & "\winrar" & Chr(34) & " e -o+ " & Chr(34) & nomarch2 & Chr(34) & " " & Chr(34) & direc1 & "\" & Chr(34)
     Print #1, "copy " & Chr(34) & nomarch3 & Chr(34) & " " & Chr(34) & nomarch4 & Chr(34)
     Print #1, "copy " & Chr(34) & nomarch2 & Chr(34) & " " & Chr(34) & nomarch5 & Chr(34)
     Print #1, "del " & Chr(34) & nomarch2 & Chr(34)
    Close #1
    salida = ExecCmd(Chr(34) & DirArchBat & "\" & txtarchbatch & Chr(34))
    MensajeProc = "Se creo el archivo " & direc2 & "\" & prefijo & txtfecha & "M.xls"
    exito = True
    txtmsg = "El proceso finalizo correctamente"
 End If
Else
 Kill (nomarch2)
 exito = False
 txtmsg = "El archivo esta vacio " & nomarch2
End If
Else
 exito = False
 txtmsg = "No hay acceso al archivo " & nomarch1
End If
End Sub

Sub TransferirVect2(ByVal fecha As Date, prefijo As String, ext As String, direc1 As String, direc2 As String, direc3 As String, siacc1, siacc2, ByVal txtarchbatch As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim fs As Object
Dim txtfecha As String
Dim nomarch1 As String
Dim nomarch2 As String
Dim nomarch3 As String
Dim nomarch4 As String
Dim nomarch5 As String
Dim sihayarch2 As Boolean
Dim sihayarch3 As Boolean
Dim salida As Variant
Dim exitoarch As Boolean

exito = False
Set fs = CreateObject("Scripting.FileSystemObject")
txtfecha = Format(fecha, "yyyymmdd")
nomarch1 = prefijo & txtfecha & ".zip"
nomarch2 = direc1 & "\" & prefijo & txtfecha & ".zip"
nomarch3 = direc1 & "\" & prefijo & txtfecha & "." & ext
nomarch4 = direc2 & "\" & prefijo & txtfecha & "." & ext
nomarch5 = direc3 & "\" & prefijo & txtfecha & ".zip"
If siacc1 Then
  Call BajarArchSFTP(nomarch1, Chr(34) & nomarch2 & Chr(34), NomSRVPIP, usersftpPIP, passsftpPIP)
End If
sihayarch2 = VerifAccesoArch(nomarch2)
If sihayarch2 And siacc2 Then
   If TamañoArch(nomarch2) <> 0 Then
      Call VerificarSalidaArchivo(DirArchBat & "\" & txtarchbatch, 1, exitoarch)
      If exitoarch Then
         'Print #1, "echo off"
         Print #1, Chr(34) & DirWinRAR & "\winrar" & Chr(34) & " e -o+ " & Chr(34) & nomarch2 & Chr(34) & " " & Chr(34) & direc1 & "\" & Chr(34)
         Print #1, "copy " & Chr(34) & nomarch3 & Chr(34) & " " & Chr(34) & nomarch4 & Chr(34)
         Print #1, "copy " & Chr(34) & nomarch2 & Chr(34) & " " & Chr(34) & nomarch5 & Chr(34)
         Print #1, "del " & Chr(34) & nomarch2 & Chr(34)
         Close #1
         salida = ExecCmd(Chr(34) & DirArchBat & "\" & txtarchbatch & Chr(34))
         If sihayarch3 Then Call DepurarVMD(nomarch3, nomarch3)
         exito = True
         txtmsg = "El proceso finalizo correctamente"
      End If
   Else
      Kill (nomarch2)
      exito = False
      txtmsg = "Es un archivo de tamaño nulo " & nomarch2
   End If
Else
 exito = False
 txtmsg = "No existe el archivo " & nomarch1
End If

End Sub

Sub DepurarVMD(ByVal nomarch1 As String, ByVal nomarch2 As String)
Dim mata() As Variant
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim txtcad As String
Dim salida As Variant
Dim noreg As Integer
Dim exitoarch As Boolean
'corrige los errores de el archivo csv

   mata = LeerArchTexto(nomarch1, ",", "Leyendo el " & nomarch1)
   Dim valor As String
   Call VerificarSalidaArchivo(nomarch2, 1, exitoarch)
   If exitoarch Then
      noreg = UBound(mata, 1)
      nocampos = UBound(mata, 2)
      For i = 1 To noreg
          If Not EsVariableVacia(mata(i, 1)) Then
             txtcad = ""
             For j = 1 To nocampos
                 valor = CStr(mata(i, j))
                 txtcad = txtcad & Trim(valor)
                 If j < nocampos Then   'agrega la coma
                   txtcad = txtcad & ","
                 End If
             Next j
             Print #1, txtcad
          End If
      Next i
      Close #1
   End If
   Call VerificarSalidaArchivo(DirArchBat & "\transf3.bat", 1, exitoarch)
   If exitoarch Then
      'Print #1, "echo off"
      Print #1, "copy " & Chr(34) & nomarch2 & Chr(34) & " " & Chr(34) & nomarch1 & Chr(34)
      Close #1
      salida = ExecCmd(Chr(34) & DirArchBat & "\transf3.bat" & Chr(34))
'se borra el archivo anterior
      MensajeProc = "Se creo el archivo " & nomarch2
   End If

End Sub

Sub ObtenerArchSFTP(ByVal fecha As Date, ByVal prefijo As String, ByVal sufijo As String, ByVal ext As String, ByVal direc1 As String, ByVal direc2 As String, ByVal direc3 As String, ByVal txtarchbat As String, ByRef txtmsg As String, ByRef exito As Boolean)
'el objetivo de esta funcion es el descargar el archivo solicitado desde un servidor
'sftp, para ello hace uso de scripts en msdos, con el fin de llamar a funciones externas
'primero llama a la funcion BajarArchSFTP que descarga el archivo
'despues crea un script que realiza la descompresion del archivo y al final
'realiza copias tanto del archivo zip como del archivo sin comprimir a
'una lista de carpetas definidas
'precondiciones de datos de entrada
'fecha en VAR_FECHAS_VAR
'prefijo - el que se indique
'sufijo  - el que se indique
'ext     - la extension del archivo que se espera obtener del archivo zip
'direc1  - la ubicacion de la descargar inicial
'direc2  - la ubicacion de la copia del archivo sin comprimir
'direc3  - la ubicacion de la copia del archivo comprimido
'txtarchbat -  el nombre del archivo batch donde se escribira el codigo a ejecutar

'salida
'los archivos nomarch2,nomarch3,nomarch4 y nomarch5
'la cadena de texto txtmsg indicando el estado del proceso
'exito indica si el proceso finalizo con exito

Dim txtfecha As String
Dim nomarch1 As String
Dim nomarch2 As String
Dim nomarch3 As String
Dim nomarch4 As String
Dim nomarch5 As String
Dim sihayarch1 As Boolean
Dim salida As Variant
Dim exito1 As Boolean
Dim txtmsg1 As String
Dim exitoarch As Boolean

txtfecha = Format(fecha, "yyyymmdd")
nomarch1 = prefijo & txtfecha & sufijo & ".ZIP"
nomarch2 = direc1 & "\" & prefijo & txtfecha & sufijo & ".ZIP"
nomarch3 = direc1 & "\" & prefijo & txtfecha & sufijo & "." & ext
nomarch4 = direc2 & "\" & prefijo & txtfecha & sufijo & "." & ext
nomarch5 = direc3 & "\" & prefijo & txtfecha & sufijo & ".ZIP"
sihayarch1 = VerifAccesoArch(nomarch2)
If sihayarch1 Then Kill (nomarch2)
Call BajarArchSFTP(nomarch1, Chr(34) & nomarch2 & Chr(34), NomSRVPIP, usersftpPIP, passsftpPIP)
exito1 = True
If exito1 Then
   sihayarch1 = VerifAccesoArch(nomarch2)
   If sihayarch1 Then
      If TamañoArch(nomarch2) <> 0 Then
         Call VerificarSalidaArchivo(DirArchBat & "\" & txtarchbat, 1, exitoarch)
         If exitoarch Then
           'Print #1, "echo off"
            Print #1, Chr(34) & DirWinRAR & "\winrar" & Chr(34) & " e -o+ " & Chr(34) & nomarch2 & Chr(34) & " " & Chr(34) & direc1 & Chr(34)
            Print #1, "copy " & Chr(34) & nomarch3 & Chr(34) & " " & Chr(34) & nomarch4 & Chr(34)
            Print #1, "copy " & Chr(34) & nomarch2 & Chr(34) & " " & Chr(34) & nomarch5 & Chr(34)
            Print #1, "del " & Chr(34) & nomarch2 & Chr(34)
            Close #1
            salida = ExecCmd(Chr(34) & DirArchBat & "\" & txtarchbat & Chr(34))
            MensajeProc = "Se creo el archivo de curvas del " & fecha
            txtmsg = "El proceso finalizo correctamente"
            exito = True
         End If
      Else
         Kill (nomarch2)
         exito = False
         txtmsg = "El archivo esta vacio"
      End If
   Else
     MensajeProc = "No se encontro el archivo " & nomarch2 & " para la fecha " & fecha
     txtmsg = MensajeProc
     exito = False
   End If
Else
  txtmsg = txtmsg1
  exito = False
End If
End Sub

Sub ObtVectAnalitico(ByVal fecha As Date, ByVal prefijo As String, ByVal sufijo As String, ByVal ext As String, ByVal direc1 As String, ByVal direc2 As String, ByVal direc3 As String, ByVal siacc1 As Boolean, ByVal siacc2 As Boolean, ByVal txtarchbatch As String, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim nomarch1 As String
Dim nomarch2 As String
Dim nomarch3 As String
Dim nomarch4 As String
Dim nomarch5 As String
Dim sihayarch1 As Boolean
Dim salida As Variant
Dim exitoarch As Boolean

txtfecha = Format(fecha, "yyyymmdd")
nomarch1 = prefijo & txtfecha & sufijo & "." & ext
nomarch2 = direc1 & "\" & prefijo & txtfecha & sufijo & "." & ext
nomarch3 = direc1 & "\" & prefijo & txtfecha & sufijo & ".zip"
nomarch4 = direc2 & "\" & prefijo & txtfecha & sufijo & "." & ext
nomarch5 = direc3 & "\" & prefijo & txtfecha & sufijo & ".zip"

If siacc1 Then
 Call BajarArchSFTP(nomarch1, Chr(34) & nomarch2 & Chr(34), NomSRVPIP, usersftpPIP, passsftpPIP)
End If
sihayarch1 = VerifAccesoArch(nomarch2)
If sihayarch1 = True And siacc2 Then
   If TamañoArch(nomarch2) <> 0 Then
      Call VerificarSalidaArchivo(DirArchBat & "\" & txtarchbatch, 1, exitoarch)
      If exitoarch Then
        'Print #1, "echo off"
         Print #1, Chr(34) & DirWinRAR & "\winrar" & Chr(34) & " a -o+ " & Chr(34) & nomarch3 & Chr(34) & " " & Chr(34) & nomarch2 & Chr(34)
         Print #1, "copy " & Chr(34) & nomarch2 & Chr(34) & " " & Chr(34) & nomarch4 & Chr(34)
         Print #1, "copy " & Chr(34) & nomarch3 & Chr(34) & " " & Chr(34) & nomarch5 & Chr(34)
         Print #1, "del " & Chr(34) & nomarch3 & Chr(34)
         Close #1
         salida = ExecCmd(Chr(34) & DirArchBat & "\" & txtarchbatch & Chr(34))
         MensajeProc = "Se creo el archivo de curvas del " & fecha
         exito = True
         txtmsg = "El proceso finalizo correctamente"
      End If
   Else
      Kill (nomarch2)
      exito = False
      txtmsg = "El archivo esta vacio"
   End If
Else
  MensajeProc = "No se encontro el vector analitico día " & fecha
  txtmsg = MensajeProc
  exito = False
End If
End Sub

Sub TransfArchDWH(ByVal fecha As Date, ByVal direc1 As String, ByVal nomarch As String, ByRef siacc1 As Boolean, ByRef exito As Boolean)
Dim txtfecha As String
Dim nomarch1 As String
Dim sihayarch As Boolean

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
exito = False
txtfecha = Format(fecha, "yyyymmdd")
nomarch1 = direc1 & "\" & nomarch
sihayarch = VerifAccesoArch(nomarch1)
If sihayarch Then
  Call SubirArchSFTP(nomarch1, nomarch, "172.22.103.117", "ftp_dwh_dar", "(r13sg0s)")
  exito = True
Else
  MensajeProc = "No existe el archivo para subir al servidor DWH"
End If
On Error GoTo 0
Exit Sub
ControlErrores:
exito = False
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub RealizarBoots(ByVal fecha As Date)
Dim matyen() As propCurva
Dim nomarch As String
Dim txtfecha As String
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim notablas As Long
Dim inbucle As TableDef
Dim cont1 As Long
Dim noreg As Long
Dim i As Long
Dim nocampos As Long
Dim exitoarch As Boolean

'esta funcion fue habilitada para el VaR de mercado
nomarch = DirBases & "\" & "ccs udi.xls"
txtfecha = Format(fecha, "yyyymmdd")

 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
'se verifica cual es el nombre de la tabla que contiene los datos
    notablas = base1.TableDefs.Count
    ReDim matnom(1 To notablas) As String
    cont1 = 1
    For Each inbucle In base1.TableDefs
     matnom(cont1) = inbucle.Name
     cont1 = cont1 + 1
    Next inbucle
 Set registros1 = base1.OpenRecordset(matnom(5), dbOpenDynaset, dbReadOnly)
 If registros1.RecordCount <> 0 Then
  registros1.MoveLast
  noreg = registros1.RecordCount
  registros1.MoveFirst
  nocampos = registros1.Fields.Count
  ReDim mattas(1 To noreg, 1 To 2) As Variant
  For i = 1 To noreg
   mattas(i, 2) = LeerTAccess(registros1, 0, i)
   mattas(i, 1) = LeerTAccess(registros1, 1, i)
  registros1.MoveNext
  Next i
  registros1.Close
  base1.Close
  Call Bootstrapp1(fecha, mattas, matyen, 180, noreg, 0)
  nomarch = DirResVaR & "\boots.txt"
  
  ReDim matsal(1 To 12000, 1 To 2) As Variant
  For i = 1 To 12000
  matsal(i, 1) = i
  matsal(i, 2) = CalculaTasa(matyen, i, 1)
  Next i
  Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
  If exitoarch Then
  For i = 1 To 12000
   Print #1, matsal(i, 1) & Chr(9) & matsal(i, 2)
  Next i
  Close #1
  End If
End If

End Sub

Sub RealizarBoots1(ByVal nomarch As String)
Dim fecha As Date
Dim matres() As propCurva
Dim sihay As Boolean
Dim base1 As DAO.Database
Dim inbucle As TableDef
Dim registros1 As DAO.recordset
Dim cont1 As Long
Dim noreg As Long
Dim i As Long
Dim notablas As Long
Dim nocampos As Long
Dim exitoarch As Boolean
 
'esta funcion fue habilitada para el VaR de mercado
sihay = VerifAccesoArch(nomarch)
If sihay Then


 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
'se verifica cual es el nombre de la tabla que contiene los datos
    notablas = base1.TableDefs.Count
    ReDim matnom(1 To notablas) As String
    cont1 = 1
    For Each inbucle In base1.TableDefs
     matnom(cont1) = inbucle.Name
     cont1 = cont1 + 1
    Next inbucle
 Set registros1 = base1.OpenRecordset(matnom(6), dbOpenDynaset, dbReadOnly)
 If registros1.RecordCount <> 0 Then
  registros1.MoveLast
  noreg = registros1.RecordCount
  registros1.MoveFirst
  nocampos = registros1.Fields.Count
  ReDim mattas(1 To noreg, 1 To 2) As Variant
  For i = 1 To noreg
   mattas(i, 2) = LeerTAccess(registros1, 0, i)
   mattas(i, 1) = LeerTAccess(registros1, 1, i) / 100
  registros1.MoveNext
  Next i
  registros1.Close
  base1.Close
  Call Bootstrapp1(fecha, mattas, matres, 28, noreg, 0)
  nomarch = DirResVaR & "\resultados bootstraping.txt"
  
  ReDim matsal(1 To noreg, 1 To 2) As Variant
  For i = 1 To noreg
      matsal(i, 1) = mattas(i, 2)
      matsal(i, 2) = CalculaTasa(matres, mattas(i, 2), 1)
  Next i
  Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
  If exitoarch Then
     For i = 1 To noreg
         Print #1, matsal(i, 1) & Chr(9) & matsal(i, 2)
     Next i
     Close #1
  End If
End If
End If
End Sub

Sub RealizarBootsLY(ByVal fecha1 As Date, ByVal fecha2 As Date)
'esta funcion fue habilitada para el VaR de mercado
Dim fecha As Date
Dim matyen() As Variant
Dim cont1 As Long
Dim notablas As Long
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim nomarch As String
Dim i As Long
Dim noreg As Long
Dim noreg1 As Long
Dim valor As Double
Dim txtinsert As String
Dim sihayarch As Boolean
Dim inbucle As TableDef
Dim nocampos As Long
Dim indice As Long
Dim txtfecha As String
Dim plazo As Long


fecha = fecha1
Do While fecha <= fecha2
nomarch = DirCurvas & "\CURVAS" & Format(fecha, "yyyymmdd") & ".XLS"
sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then

 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
'se verifica cual es el nombre de la tabla que contiene los datos
    notablas = base1.TableDefs.Count
    ReDim matnom(1 To notablas) As String
    cont1 = 1
    For Each inbucle In base1.TableDefs
     matnom(cont1) = inbucle.Name
     cont1 = cont1 + 1
    Next inbucle
 Set registros1 = base1.OpenRecordset(matnom(1), dbOpenDynaset, dbReadOnly)
 If registros1.RecordCount <> 0 Then
  registros1.MoveLast
  noreg = registros1.RecordCount
  registros1.MoveFirst
  nocampos = registros1.Fields.Count
  ReDim mata(1 To noreg, 1 To 2) As Variant

  For i = 1 To noreg
   mata(i, 2) = i
   mata(i, 1) = LeerTAccess(registros1, 48, i)
  registros1.MoveNext
  Next i
  
  ReDim mattas(1 To 401, 1 To 2) As Variant
  mattas(1, 1) = mata(1, 1)
  mattas(1, 2) = 1
  For i = 1 To 400
  mattas(i + 1, 1) = mata(30 * i, 1)
  mattas(i + 1, 2) = 30 * i
  Next i
  registros1.Close
  base1.Close
  Call Bootstrapp(mattas, matyen, 30, 360, 12000, 1)
  noreg1 = UBound(matyen, 1)
  txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
  For i = 1 To NoFactores
  If MatCaracFRiesgo(i).nomFactor = "C LIBOR YEN" Then
  plazo = MatCaracFRiesgo(i).plazo
  indice = BuscarValorArray(plazo, matyen, 2)
  If indice <> 0 Then
    valor = matyen(indice, 1)
    txtinsert = "INSERT INTO " & TablaFRiesgoO & " VALUES("
    txtinsert = txtinsert & txtfecha & ","
    txtinsert = txtinsert & "'C LIBOR YEN',"
    txtinsert = txtinsert & MatCaracFRiesgo(i).plazo & ","
    txtinsert = txtinsert & valor & ","
    txtinsert = txtinsert & "'" & CLng(fecha) & Trim("C LIBOR YEN") & Trim(Format(MatCaracFRiesgo(i).plazo, "0000000")) & "')"
    ConAdo.Execute txtinsert
  End If
  End If
  Next i
   
  
  
End If
End If
fecha = fecha + 1

Loop
End Sub

Sub CompletarPosMesaD(ByRef matpos() As propPosMD, ByRef matvp() As Variant)
Dim i As Long
Dim indice As Long
Dim txtemision As String
Dim tv As String
Dim noreg As Long

If IsArray(matpos) And UBound(matvp, 1) <> 0 Then
noreg = UBound(matpos, 1)
For i = 1 To noreg
'se lee la historia fondeo bancario promedio
    txtemision = matpos(i).cEmisionMD
    If Not EsVariableVacia(txtemision) And (matpos(i).Tipo_Mov = 1 Or matpos(i).Tipo_Mov = 4 Or matpos(i).Tipo_Mov = 6 Or matpos(i).Tipo_Mov = 7) Then      'if 1
       indice = BuscarValorArray(txtemision, matvp, 6)
       If indice <> 0 Then
          matpos(i).valSucioPIP = matvp(indice, 2)             'precio sucio pip
          matpos(i).valLimpioPIP = matvp(indice, 3)            'precio limpio pip
          matpos(i).tCuponVigenteMD = matvp(indice, 4)         'tasa cupon vigente
          matpos(i).fVencMD = matvp(indice, 8)                 'FECHA DE VENCIMIENTO
          matpos(i).Calif1MD = matvp(indice, 10)               'calificacion sp
          matpos(i).Calif2MD = matvp(indice, 11)               'calificacion fitch
          matpos(i).Calif3MD = matvp(indice, 12)               'calificacion moodys
          matpos(i).Calif4MD = matvp(indice, 13)               'calificacion hr
          matpos(i).reglaCuponMD = matvp(indice, 14)           'regla cupon
   
          If matpos(i).vNominalMD = 0 And matvp(indice, 7) <> 0 Then matpos(i).vNominalMD = matvp(indice, 7)
          If Val(matpos(i).pCuponVigenteMD) = 0 Then matpos(i).pCuponVigenteMD = matvp(indice, 5) 'periodo cupon
          tv = matpos(i).tValorMD
          If (tv = "XA" Or tv = "LD") Then
             matpos(i).intDevengMD = matvp(indice, 9)
          End If
       Else
          MensajeProc = "Falta la emision " & txtemision & " en el vector de precios"
       End If
    End If   'if 1
    AvanceProc = i / noreg
    MensajeProc = "Comparando con el vector de precios  y agregando datos faltantes " & Format(AvanceProc, "##0.00 %")
Next i
Else
   MensajeProc = "No se leyo el vector de precios para esta fecha"
End If
End Sub

Function CompletarPosMVP(ByVal fecha As Date, ByRef matp() As Variant, ByRef matv() As Variant)
Dim noreg As Long
Dim i As Long
Dim indice As Long
Dim tvnominal As Integer
Dim tfven As String
Dim txtemision As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'esta es la estructura de la informacion actual
'1     fecha de la posicion
'2     intencion
'3     fecha de la posicion
'4     tipo valor
'5     emisora
'6     serie
'7     fecha de vencimiento de la operacion
'8     periodo cupon del papel
'9     dias x vencer operacion
'10    tasa cupon vigente de la operacion
'11    tasa del reporto
'12    no de titulos
'13    mesa o posicion
'14    precio pactado
'15    fecha de inicio de la operacion
'16    clave de la emision
'17    valor nominal del papel
'18    valor nominal del papel

If IsArray(matp) And UBound(matv, 1) > 0 Then
noreg = UBound(matp, 1)
For i = 1 To noreg
'se verifica que compagine con el vector de precios
If matp(i, 3) = 1 Or matp(i, 3) = 4 Or matp(i, 3) = 6 Or matp(i, 3) = 7 Or matp(i, 3) = 10 Then
txtemision = matp(i, 16)
indice = BuscarValorArray(txtemision, matv, 18)
If indice <> 0 Then
   If matp(i, 4) <> matv(indice, 3) Then matp(i, 4) = matv(indice, 3)           'tipo valor
   If matp(i, 5) <> matv(indice, 4) Then matp(i, 5) = matv(indice, 4)           'emisora
   If matp(i, 6) <> matv(indice, 5) Then matp(i, 6) = matv(indice, 5)           'serie
   If matp(i, 7) <> matv(indice, 12) Then matp(i, 7) = matv(indice, 12)         'fecha de vencimiento papel
   If matp(i, 8) <> matv(indice, 16) Then matp(i, 8) = matv(indice, 16)         'periodo cupon
   If matp(i, 10) <> matv(indice, 15) Then matp(i, 10) = matv(indice, 15)       'tasa cupon
   If matp(i, 17) <> matv(indice, 11) Then matp(i, 17) = matv(indice, 11)       'VALOR NOMINAL PAPEL
Else
 MensajeProc = "Falta la emision " & txtemision & " en el vector de precios del " & fecha
 MsgBox MensajeProc
 If EsVariableVacia(matp(i, 17)) Then
  If matp(i, 4) = "F" Then
   tvnominal = 100
  ElseIf matp(i, 4) = "I" Then
   tvnominal = 1
  Else
   tvnominal = InputBox("Dame el valor nominal de " & txtemision, , 0)
  End If
  matp(i, 17) = Val(tvnominal)
 End If
 If EsVariableVacia(matp(i, 7)) Then
  tfven = InputBox("dame la fecha de vencimiento", , Date)
  matp(i, 7) = CDate(tfven)
 End If
End If
'MsgBox "Faltan datos del vector de precios para este día"
End If
AvanceProc = i / noreg
MensajeProc = "Comparando con el vector de precios  y agregando datos faltantes " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i
CompletarPosMVP = matp
End If
On Error GoTo 0
Exit Function
ControlErrores:
MsgBox "CompletarPosMesaD " & error(Err())
On Error GoTo 0
End Function

Sub BuscarEmVP(ByVal fecha As Date, ByRef matp() As Variant, ByRef matv() As Variant)
Dim noreg As Long
Dim i As Long
Dim indice As Long
Dim txtemision As String


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

If IsArray(matp) And UBound(matv, 1) > 0 Then
noreg = UBound(matp, 1)
For i = 1 To noreg
'se verifica que compagine con el vector de precios
    If matp(i, 3) = 1 Or matp(i, 3) = 4 Or matp(i, 3) = 6 Or matp(i, 3) = 7 Or matp(i, 3) = 10 Then
       txtemision = matp(i, 7)
       indice = BuscarValorArray(txtemision, matv, 18)
       If indice = 0 Then
          MensajeProc = "Falta la emision " & txtemision & " en el vector de precios del " & fecha
          MsgBox MensajeProc
          Call MostrarMensajeSistema(MensajeProc, frmProgreso.Label2, 2, Date, Time, NomUsuario)
       End If
    End If
    Call MostrarMensajeSistema("Comparando con el vector de precios " & Format(AvanceProc, "#,##0.00 %"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
    DoEvents
Next i
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox "BuscarEmVP " & error(Err())
On Error GoTo 0
End Sub

Function DevValorB(ByVal txtcad As String, ByRef mata() As Variant, ByVal ind1 As Integer, ByVal ind2 As Integer)
Dim indice As Long

indice = BuscarValorArray(txtcad, mata, ind1)
If indice <> 0 Then
 DevValorB = mata(indice, ind2)
Else
 DevValorB = 0
End If
End Function


Sub ActBaseOracleAccess(ByVal txtnbase As String, ByVal texto1 As String, ByVal texto2 As String, ByVal nr As Integer)
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim nocampos As Long
Dim sql_num_mesa As String
Dim sql_mesa As String
Dim rmesa As New ADODB.recordset

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'se actualiza una tabla de oracle a access

sql_num_mesa = "select count(*) from " & texto2
rmesa.Open sql_num_mesa, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
sql_mesa = "select * from " & texto2
rmesa.Open sql_mesa, ConAdo
rmesa.MoveFirst
nocampos = rmesa.Fields.Count
ReDim mata(1 To noreg, 1 To nocampos) As Variant
For i = 1 To noreg
For j = 1 To nocampos
 mata(i, j) = rmesa.Fields(j - 1)
Next j
rmesa.MoveNext
AvanceProc = i / noreg
MensajeProc = "Actualizando la tabla " & txtnbase & " " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i
rmesa.Close
End If

Set base1 = OpenDatabase(txtnbase, dbDriverNoPrompt, False, ";Pwd=" & ContraseñaCatalogos)
Set registros1 = base1.OpenRecordset("SELECT * FROM " & texto1, dbOpenDynaset)
base1.Execute "DELETE FROM " & texto1
nocampos = registros1.Fields.Count
For i = 1 To noreg
registros1.AddNew
For j = 1 To nocampos
If Len(mata(i, j)) <> 0 Then
 If registros1.Fields(j - 1).Type = 8 Then
  Call GrabarTAccess(registros1, j - 1, CDate(mata(i, j)), i)
 ElseIf registros1.Fields(j - 1).Type = 10 Then
  Call GrabarTAccess(registros1, j - 1, CStr(mata(i, j)), i)
 ElseIf registros1.Fields(j - 1).Type = 4 Then
  Call GrabarTAccess(registros1, j - 1, CLng(mata(i, j)), i)
 ElseIf registros1.Fields(j - 1).Type = 7 Then
  Call GrabarTAccess(registros1, j - 1, CDbl(mata(i, j)), i)
 End If
End If
Next j
registros1.Update
AvanceProc = i / noreg
MensajeProc = "Actualizando la tabla " & txtnbase & " " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i
registros1.Close
base1.Close
nr = noreg
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub CatValuacion(ByVal nomarch As String)
Dim txtfiltro As String
Dim nombase As String
 'valuacion bonos cupon cero
 Call ActCatAccessOracle(nomarch, TablaValBonosC0, PrefijoBD & TablaValBonosC0)
 'reportos
 Call ActCatAccessOracle(nomarch, TablaValReportos, PrefijoBD & TablaValReportos)
 'bonos a tasa fija
 Call ActCatAccessOracle(nomarch, TablaValBonos, PrefijoBD & TablaValBonos)
 Call ActCatAccessOracle(nomarch, TablaValBSC, PrefijoBD & TablaValBSC)
 'bonos con sobretasa descuento
 Call ActCatAccessOracle(nomarch, TablaValBSD, PrefijoBD & TablaValBSD)
 'swaps
 Call ActCatAccessOracle(nomarch, TablaValSwaps1, PrefijoBD & TablaValSwaps1)
 Call ActCatAccessOracle(nomarch, TablaValSwaps2, PrefijoBD & TablaValSwaps2)
'deuda
 Call ActCatAccessOracle(nomarch, TablaValDeuda, PrefijoBD & TablaValDeuda)
 'relacion swaps deuda
 Call ActCatAccessOracle(nomarch, TablaRSwapsDeuda, PrefijoBD & TablaRSwapsDeuda)
 Call ActCatAccessOracle(nomarch, TablaRelSwapIKOSS, PrefijoBD & TablaRelSwapIKOSS)
 Call ActCatAccessOracle(nomarch, TablaBlEsc, PrefijoBD & TablaBlEsc)
 'FORWARDS TIPO CAMBIO
 Call ActCatAccessOracle(nomarch, TablaValFwds1, PrefijoBD & TablaValFwds1)
 Call ActCatAccessOracle(nomarch, TablaValFwds2, PrefijoBD & TablaValFwds2)
 Call ActCatAccessOracle(nomarch, TablaValInds, PrefijoBD & TablaValInds)
 Call ActTablaOracle(TablaDerivEst2, TablaDerEstandar, TablaDerivEst2, ConAdo, ConAdo)


End Sub

Sub CatDerivados(ByVal nomarch As String)
Dim txtfiltro As String
Dim nombase As String
   Call ActCatAccessOracle(nomarch, TablaPosPrimarias, PrefijoBD & TablaPosPrimarias)
   Call ActCatAccessOracle(nomarch, TablaBlackList, PrefijoBD & TablaBlackList)
   Call ActCatAccessOracle(nomarch, TablaGruposDeriv, PrefijoBD & TablaGruposDeriv)
   Call ActCatAccessOracle(nomarch, TablaCalificaciones, PrefijoBD & TablaCalificaciones)
   Call ActCatAccessOracle(nomarch, TablaRelSwapEm, PrefijoBD & TablaRelSwapEm)
   'Call ActCatAccessOracle(nomarch, TablaCalendSwapsO, PrefijoBD &TablaCalendSwapsO)
   'Call ActCatAccessOracle(nomarch, TablaCalendSwaps, PrefijoBD &TablaCalendSwaps)

End Sub

Sub CatEstReportes(ByVal nomarch As String)

 Call ActCatAccessOracle(nomarch, TablaGruposPortPos, PrefijoBD & TablaGruposPortPos)
 Call ActCatAccessOracle(nomarch, TablaReporteCVaR, PrefijoBD & TablaReporteCVaR)
 'Call ActCatAccessOracle(nomarch, TablaSQLPort, PrefijoBD & TablaSQLPort)
 Call ActCatAccessOracle(nomarch, TablaGruposPapelFP, PrefijoBD & TablaGruposPapelFP)
End Sub

Sub CatContrapartes(ByVal nomarch As String)

 Call ActCatAccessOracle(nomarch, TablaContrapartes, PrefijoBD & TablaContrapartes)
 Call ActCatAccessOracle(nomarch, TablaEquivContrap, PrefijoBD & TablaEquivContrap)
 Call ActCatAccessOracle(nomarch, TablaTreshCont, PrefijoBD & TablaTreshCont)
 Call ActCatAccessOracle(nomarch, TablaCalifContrapF, PrefijoBD & TablaCalifContrapF)
 Call ActCatAccessOracle(nomarch, TablaCalifContrapNF, PrefijoBD & TablaCalifContrapNF)
 Call ActCatAccessOracle(nomarch, TablaCalifContrapEmision, PrefijoBD & TablaCalifContrapEmision)
 Call ActCatAccessOracle(nomarch, TablaBLTresh, PrefijoBD & TablaBLTresh)
 Call ActCatAccessOracle(nomarch, TablaRecInt, PrefijoBD & TablaRecInt)
 Call ActCatAccessOracle(nomarch, TablaRecNacional, PrefijoBD & TablaRecNacional)
 Call ActCatAccessOracle(nomarch, TablaMTrans, PrefijoBD & TablaMTrans)
 Call ActCatAccessOracle(nomarch, TablaEmxContrap, PrefijoBD & TablaEmxContrap)
 Call ActCatAccessOracle(nomarch, TablaSectorEscEm, PrefijoBD & TablaSectorEscEm)
 Call ActCatAccessOracle(nomarch, TablaEscCortoLargo, PrefijoBD & TablaEscCortoLargo)

End Sub

Sub CatParametros(ByVal nomarch As String)
Dim txtfiltro As String
Dim nombase As String

 Call ActCatAccessOracle(nomarch, TablaLimites, PrefijoBD & TablaLimites)
 Call ActCatAccessOracle(nomarch, TablaParamSistema, PrefijoBD & TablaParamSistema)
 Call ActCatAccessOracle(nomarch, TablaFechasEscEstres, PrefijoBD & TablaFechasEscEstres)
 Call ActCatAccessOracle(nomarch, TablaCatPortPos, PrefijoBD & TablaCatPortPos)
 Call ActCatAccessOracle(nomarch, TablaPortPosEstructural, PrefijoBD & TablaPortPosEstructural)
 Call ActCatAccessOracle(nomarch, TablaCEmision, PrefijoBD & TablaCEmision)
 Call ActCatAccessOracle(nomarch, TablaPEmision, PrefijoBD & TablaPEmision)
 'Call ActCatAccessOracle(nomarch, TablaBaseCalendO,PrefijoBD & TablaBaseCalendO)
 Call ActCatAccessOracle(nomarch, TablaPortPrincipales, PrefijoBD & TablaPortPrincipales)

End Sub

Sub CatFactRiesgo(ByVal nomarch As String)
Dim txtfiltro As String
Dim nombase As String
'portafolio factores riesgo
 Call ActCatAccessOracle(nomarch, TablaPortFR, PrefijoBD & TablaPortFR)
 Call ActCatAccessOracle(nomarch, TablaHistCurvas, PrefijoBD & TablaHistCurvas)
 Call ActCatAccessOracle(nomarch, TablaNodosCurvas, PrefijoBD & TablaNodosCurvas)
 Call ActCatAccessOracle(nomarch, TablaIndVecPreciosO, PrefijoBD & TablaIndVecPreciosO)
 Call ActCatAccessOracle(nomarch, TablaCatCurvas, PrefijoBD & TablaCatCurvas)
 Call ActCatAccessOracle(nomarch, TablaSplits, PrefijoBD & TablaSplits)

End Sub

Sub CatMO(ByVal nomarch As String)
 Call ActCatAccessOracle(nomarch, TablaMOEmSectorMD, PrefijoBD & TablaMOEmSectorMD)
 Call ActCatAccessOracle(nomarch, TablaMOEmSectorPI, PrefijoBD & TablaMOEmSectorPI)
 Call ActCatAccessOracle(nomarch, TablaMOperCalif, PrefijoBD & TablaMOperCalif)
 Call ActCatAccessOracle(nomarch, TablaMOEmPriv, PrefijoBD & TablaMOEmPriv)
 Call ActCatAccessOracle(nomarch, TablaMOMon, PrefijoBD & TablaMOMon)
 Call ActCatAccessOracle(nomarch, TablaMOContrapCalifPI, PrefijoBD & TablaMOContrapCalifPI)
 Call ActCatAccessOracle(nomarch, TablaMOContrap, PrefijoBD & TablaMOContrap)
 Call ActCatAccessOracle(nomarch, TablaMOGub, PrefijoBD & TablaMOGub)
End Sub

Sub CatProcesos(ByVal nomarch As String)

'la lista de tareas díarias a realizar
 Call ActCatAccessOracle(nomarch, TablaCatProcesos, PrefijoBD & TablaCatProcesos)
 Call ActCatAccessOracle(nomarch, TablaSecProcesos, PrefijoBD & TablaSecProcesos)
 Call ActCatAccessOracle(nomarch, TablaSecSubProc, PrefijoBD & TablaSecSubProc)

End Sub

Sub SincCatAccessOracle(ByVal nomarch As String)
  Call CatFactRiesgo(nomarch)
  Call CatEstReportes(nomarch)
  Call CatProcesos(nomarch)
  Call CatDerivados(nomarch)
  Call CatValuacion(nomarch)
  Call CatContrapartes(nomarch)
  Call CatParametros(nomarch)
  Call CatMO(nomarch)
  MensajeProc = "Se realizo una actualizacion de los catalogos del sistema"
  Call GuardaDatosBitacora(5, "Administracion", 0, MensajeProc, NomUsuario, Date, MensajeProc, 1)
End Sub

Sub ActCatAccessOracle(ByVal txtnbase As String, ByVal texto1 As String, ByVal texto2 As String)
'funcion : copiar el contenido de una tabla en access a una tabla en oracle
'datos de entrada
'txtnbase    - ruta y nombre del archivo de access con las tablas a copiar
'texto1      - nombre de la tabla de access a copiar
'texto2      - nombre de la tabla de oracle cuyos datos seran borrados
'condicion inicial: las tablas texto1 y texto2 deben tener el mismo numero de columnas
'y el mismo tipo de datos en cada campo
'resultado del proceso: los datos de la tabla texto2 son reemplazados en su totalidad
'por los datos de la tabla texto1
'esta rutina se utiliza con el fin de actualizar los catalogos del sistema de riesgos


Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim nocampos As Integer
Dim txtcadena As String
Dim txtfecha As String

ConAdo.Execute "DELETE FROM " & texto2

Set base1 = OpenDatabase(txtnbase, dbDriverNoPrompt, False, ";Pwd=" & ContraseñaCatalogos)
Set registros1 = base1.OpenRecordset(texto1, dbOpenDynaset, dbReadOnly)
registros1.MoveLast
noreg = registros1.RecordCount
registros1.MoveFirst
nocampos = registros1.Fields.Count
For i = 1 To noreg
    ReDim mata(1 To nocampos) As Variant
    ReDim matt(1 To nocampos) As Variant
    For j = 1 To nocampos
    If Not EsVariableVacia(LeerTAccess(registros1, j - 1, i)) Then
       mata(j) = LeerTAccess(registros1, j - 1, i)
    Else
       mata(j) = ""
    End If
    matt(j) = registros1.Fields(j - 1).Type
    Next j
    txtcadena = "INSERT INTO " & texto2 & " VALUES("
    For j = 1 To nocampos - 1
        If matt(j) = 1 Then
           If Not EsVariableVacia(mata(j)) Then
              txtfecha = "to_date('" & Format(mata(j), "dd/mm/yyyy") & "','dd/mm/yyyy')"
              txtcadena = txtcadena & txtfecha & ","            'fecha
           Else
              txtcadena = txtcadena & "null,"
           End If
        ElseIf matt(j) = 8 Then
       If Not EsVariableVacia(mata(j)) Then
          txtfecha = "to_date('" & Format(mata(j), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtcadena = txtcadena & txtfecha & ","            'fecha
       Else
          txtcadena = txtcadena & "null,"
       End If
    ElseIf matt(j) = 10 Then
       If Not EsVariableVacia(mata(j)) Then
          txtcadena = txtcadena & "'" & ReemplazaCadenaTexto(mata(j), "'", "''") & "',"     'texto
       Else
          txtcadena = txtcadena & "null,"
       End If
    ElseIf matt(j) = 4 Then
       If Not EsVariableVacia(mata(j)) Then
          txtcadena = txtcadena & mata(j) & ","             'numerico
       Else
          txtcadena = txtcadena & "null,"
       End If
    ElseIf matt(j) = 7 Then
       If Not EsVariableVacia(mata(j)) Then
          txtcadena = txtcadena & mata(j) & ","             'numerico
       Else
          txtcadena = txtcadena & "null,"
       End If
    ElseIf matt(j) = 3 Then
       If Not EsVariableVacia(mata(j)) Then
          txtcadena = txtcadena & mata(j) & ","             'numerico
       Else
          txtcadena = txtcadena & "null,"
       End If
    ElseIf matt(j) = 12 Then
       If Not EsVariableVacia(mata(j)) Then
          txtcadena = txtcadena & "'" & ReemplazaCadenaTexto(mata(j), "'", "''") & "',"        'memo
       Else
          txtcadena = txtcadena & "null,"
       End If
    Else
       MsgBox "no se clasifico el campo" & registros1.Fields(j - 1).Name
    End If
    Next j
If matt(nocampos) = 1 Then
 If Not EsVariableVacia(mata(nocampos)) Then
  txtfecha = "to_date('" & Format(mata(nocampos), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtcadena = txtcadena & txtfecha & ")"                   'fecha
 Else
  txtcadena = txtcadena & "null)"
 End If
ElseIf matt(nocampos) = 8 Then
 If Not EsVariableVacia(mata(nocampos)) Then
  txtfecha = "to_date('" & Format(mata(nocampos), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtcadena = txtcadena & txtfecha & ")"                   'fecha
 Else
  txtcadena = txtcadena & "null)"
 End If
ElseIf matt(nocampos) = 10 Then
If Not EsVariableVacia(mata(nocampos)) Then
 txtcadena = txtcadena & "'" & ReemplazaCadenaTexto(mata(nocampos), "'", "''") & "')"       'texto
Else
 txtcadena = txtcadena & "null)"
End If
ElseIf matt(nocampos) = 12 Then
If Not EsVariableVacia(mata(nocampos)) Then
 txtcadena = txtcadena & "'" & ReemplazaCadenaTexto(mata(nocampos), "'", "''") & "')"       'memo
Else
 txtcadena = txtcadena & "null)"
End If
ElseIf matt(nocampos) = 4 Then
 If Not EsVariableVacia(mata(nocampos)) Then
  txtcadena = txtcadena & mata(nocampos) & ")"            'numerico
 Else
  txtcadena = txtcadena & "null)"
 End If
ElseIf matt(nocampos) = 7 Then
 If Not EsVariableVacia(mata(nocampos)) Then
  txtcadena = txtcadena & mata(nocampos) & ")"            'numerico
 Else
  txtcadena = txtcadena & "null)"
 End If
End If
'MsgBox txtcadena
ConAdo.Execute txtcadena
registros1.MoveNext
AvanceProc = i / noreg
MensajeProc = "Actualizando la tabla " & texto2 & " " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i
registros1.Close
base1.Close

End Sub

Sub ActBaseAccessOracle(ByVal txtbase1 As String, ByVal txtfiltro1 As String, ByVal txtbase2 As String, ByVal txtfiltro2 As String)
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim txtfecha1 As String
Dim txtconcepto As String
Dim txtplazo As String
Dim txtfecha As String
Dim txtcadena As String

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

Set base1 = OpenDatabase(txtbase1, dbDriverNoPrompt, False, ";Pwd=" & ContraseñaCatalogos)
Set registros1 = base1.OpenRecordset(txtfiltro1, dbOpenDynaset, dbReadOnly)
If registros1.RecordCount <> 0 Then
registros1.MoveLast
noreg = registros1.RecordCount
registros1.MoveFirst
nocampos = registros1.Fields.Count


For i = 1 To noreg
ReDim mata(1 To nocampos) As Variant
ReDim matt(1 To nocampos) As Variant
For j = 1 To nocampos
If Not EsVariableVacia(LeerTAccess(registros1, j - 1, i)) Then
 mata(j) = LeerTAccess(registros1, j - 1, i)
Else
 mata(j) = ""
End If
matt(j) = registros1.Fields(j - 1).Type
Next j
txtfecha1 = "to_date('" & Format(mata(1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtconcepto = mata(2)
txtplazo = mata(3)

txtcadena = "INSERT INTO " & txtbase2 & " VALUES("
For j = 1 To nocampos - 1
If matt(j) = 1 Then
    txtcadena = txtcadena & ReemplazaVacioValorO(mata(j), Null, 1) & ","        'fecha
ElseIf matt(j) = 8 Then
    txtcadena = txtcadena & ReemplazaVacioValorO(mata(j), Null, 1) & ","       'fecha
ElseIf matt(j) = 10 Then
    txtcadena = txtcadena & ReemplazaVacioValorO(mata(j), Null, 3) & ","       'texto
ElseIf matt(j) = 4 Then
    txtcadena = txtcadena & ReemplazaVacioValorO(mata(j), Null, 2) & ","       'valor
ElseIf matt(j) = 7 Then
    txtcadena = txtcadena & ReemplazaVacioValorO(mata(j), Null, 2) & ","       'valor
ElseIf matt(j) = 3 Then
    txtcadena = txtcadena & ReemplazaVacioValorO(mata(j), Null, 2) & ","       'valor
ElseIf matt(j) = 12 Then
    txtcadena = txtcadena & ReemplazaVacioValorO(mata(j), Null, 3) & ","       'memo
Else
     MsgBox "no se clasifico el campo" & registros1.Fields(j - 1).Name
End If
Next j
If matt(nocampos) = 1 Then
   txtcadena = txtcadena & ReemplazaVacioValorO(mata(nocampos), Null, 1) & ")"      'fecha
ElseIf matt(nocampos) = 8 Then
   txtcadena = txtcadena & ReemplazaVacioValorO(mata(nocampos), Null, 1) & ")"      'fecha
ElseIf matt(nocampos) = 10 Then
   txtcadena = txtcadena & ReemplazaVacioValorO(mata(nocampos), Null, 3) & ")"      'texto
ElseIf matt(nocampos) = 4 Then
   txtcadena = txtcadena & ReemplazaVacioValorO(mata(nocampos), Null, 2) & ")"      'valor
ElseIf matt(nocampos) = 7 Then
   txtcadena = txtcadena & ReemplazaVacioValorO(mata(nocampos), Null, 2) & ")"      'valor
ElseIf matt(j) = 3 Then
   txtcadena = txtcadena & ReemplazaVacioValorO(mata(nocampos), Null, 2) & ")"      'valor
ElseIf matt(nocampos) = 12 Then
   txtcadena = txtcadena & ReemplazaVacioValorO(mata(nocampos), Null, 3) & ")"      'memo
End If
ConAdo.Execute txtcadena
registros1.MoveNext
AvanceProc = i / noreg
MensajeProc = "Actualizando la tabla " & txtbase2 & " " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i
End If
registros1.Close
base1.Close

On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function ReemplazaVacioValor(x, Y)
'esta funcion verifica si el valor x es un valor nulo
'en ese caso reemplaza a x por y

On Error GoTo hayerror
If Not EsVariableVacia(x) Then
   ReemplazaVacioValor = x
Else
   ReemplazaVacioValor = Y
End If
Exit Function
hayerror:
ReemplazaVacioValor = Y
End Function

Function ReemplazaVacioValorO(x, Y, ByVal tvalor As Integer)
If Not EsVariableVacia(x) Then
   If tvalor = 1 Then    'cadena
      ReemplazaVacioValorO = "to_date(" & Format(x, "dd/MM/yyyy") & "),'dd/mm/yyyy')"
   ElseIf tvalor = 2 Then    'fecha
      ReemplazaVacioValorO = x
   ElseIf tvalor = 3 Then    'valor
      ReemplazaVacioValorO = "'" & x & "'"
   End If
Else
   ReemplazaVacioValorO = Y
End If
End Function


Sub ActTablaOracle(ByVal txttabla As String, ByVal txtfiltro1 As String, ByVal txtfiltro2 As String, ByRef origen1 As ADODB.Connection, ByRef origen2 As ADODB.Connection)
Dim noreg As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim txtfecha As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

'primero define la tabla a leer
rmesa.Open "SELECT COUNT(*) from " & txtfiltro1, origen1
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
origen2.Execute "DELETE FROM " & txtfiltro2
rmesa.Open "SELECT * from " & txtfiltro1, origen1
rmesa.MoveFirst
nocampos = rmesa.Fields.Count
For i = 1 To noreg
ReDim mata(1 To nocampos) As Variant
ReDim matt(1 To nocampos) As Variant
For j = 1 To nocampos
If Not EsVariableVacia(rmesa.Fields(j - 1)) Then
 mata(j) = rmesa.Fields(j - 1)
Else
 mata(j) = ""
End If
 matt(j) = rmesa.Fields(j - 1).Type
Next j
txtcadena = "INSERT INTO " & txttabla & " VALUES("
For j = 1 To nocampos - 1
Select Case matt(j)
Case 1, 8, 135
 If Not EsVariableVacia(mata(j)) Then
  txtfecha = "to_date('" & Format(mata(j), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtcadena = txtcadena & txtfecha & ","                'fecha
 Else
  txtcadena = txtcadena & "null,"
 End If
Case 10, 200, 129
 If Not EsVariableVacia(mata(j)) Then
  txtcadena = txtcadena & "'" & mata(j) & "',"          'texto
 Else
  txtcadena = txtcadena & "null,"
 End If
Case 3, 4, 7, 131
    If Not EsVariableVacia(mata(j)) Then
     txtcadena = txtcadena & mata(j) & ","               'numerico
    Else
     txtcadena = txtcadena & "null,"
    End If
Case 12
    If Not EsVariableVacia(mata(j)) Then
       txtcadena = txtcadena & "'" & mata(j) & "',"        'memo
    Else
       txtcadena = txtcadena & "null,"
    End If
Case Else
    MsgBox "no se clasifico el campo" & rmesa.Fields(j - 1).Name
End Select
Next j
Select Case matt(nocampos)
Case 1, 8, 135
 If Not EsVariableVacia(mata(nocampos)) Then
  txtfecha = "to_date('" & Format(mata(nocampos), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtcadena = txtcadena & txtfecha & ")"                   'fecha
 Else
  txtcadena = txtcadena & "null)"
 End If
Case 10, 200, 129
 If Not EsVariableVacia(mata(nocampos)) Then
  txtcadena = txtcadena & "'" & mata(nocampos) & "')"       'texto
 Else
  txtcadena = txtcadena & "null)"
 End If
Case 12
 If Not EsVariableVacia(mata(nocampos)) Then
  txtcadena = txtcadena & "'" & mata(nocampos) & "')"       'memo
 Else
  txtcadena = txtcadena & "null)"
 End If
Case 3, 4, 7, 131
 If Not EsVariableVacia(mata(nocampos)) Then
  txtcadena = txtcadena & mata(nocampos) & ")"            'numerico
 Else
  txtcadena = txtcadena & "null)"
 End If
Case Else
   MsgBox "no se clasifico el campo" & rmesa.Fields(j - 1).Name
End Select

origen2.Execute txtcadena
rmesa.MoveNext
AvanceProc = i / noreg
MensajeProc = "Actualizando la tabla " & txtfiltro2 & " " & Format(AvanceProc, "##0.00 %")
DoEvents
Next i
rmesa.Close
End If
End Sub

Function LeerParamUsuario(ByVal usuario As String)
'lee los parametros personalizados del usuario y los carga en un array en memoria
'estos parametros son:
'ubicacion de archivos
'datos de entrada:
'usuario   - nombre del usuario que acceso al sistema
salida:
'array con los parametros del usuario

Dim txtfiltro As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

txtfiltro2 = "select * from " & TablaParamUsuario & " WHERE USUARIO = '" & usuario & "' ORDER BY PARAMETRO"
txtfiltro = "select count(*) from " & TablaParamUsuario & " WHERE USUARIO = '" & usuario & "'"
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mata(1 To noreg, 1 To 2) As Variant
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("PARAMETRO")
       mata(i, 2) = ReemplazaCadenaTexto(ReemplazaVacioValor(rmesa.Fields("VALOR"), ""), "Subdireccisn", "Subdirección")
       rmesa.MoveNext
   Next i
   rmesa.Close
   mata = RutinaOrden(mata, 1, SRutOrden)
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerParamUsuario = mata
End Function

Function LeerParametrosSist()
'lee la lista de parametros genericos del sistema y los carga en un array de memoria
'esta rutina se llama para crear los parametros de un nuevo usuario en el sistema

Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim rmesa As New ADODB.recordset

txtfiltro2 = "select * from " & PrefijoBD & TablaParamSistema & " ORDER BY ID_PARAM"
txtfiltro1 = "select count(*) from (" & txtfiltro2 & ")"

rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
ReDim mata(1 To noreg, 1 To 3) As Variant
 rmesa.Open txtfiltro2, ConAdo
 rmesa.MoveFirst
 For i = 1 To noreg
 mata(i, 1) = rmesa.Fields("ID_PARAM")
 mata(i, 2) = rmesa.Fields("PARAMETRO")
 mata(i, 3) = ReemplazaCadenaTexto(rmesa.Fields("VALOR"), "Subdireccisn", "Subdirección")
 rmesa.MoveNext
 Next i
 rmesa.Close
 mata = RutinaOrden(mata, 1, SRutOrden)
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerParametrosSist = mata
End Function

Sub ValidarParamUsuario(usuario)
Dim i As Integer
Dim indice As Integer

MatParamUsuario = LeerParamUsuario(usuario)
'ubicacions de factores de riesgo
For i = 1 To UBound(MatParamSistema, 1)
indice = BuscarValorArray(MatParamSistema(i, 2), MatParamUsuario, 1)
  If indice = 0 Then
     MsgBox "Falta el parametro " & MatParamSistema(i, 2) & " en la tabla de parametros"
  End If
Next i
For i = 1 To UBound(MatParamUsuario, 1)
If MatParamUsuario(i, 1) = "DirCurvasCSV1" Then
   DirCurvasCSV1 = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirCurvas" Then
   DirCurvas = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirCurvasZ" Then
   DirCurvasZ = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirVPrecios" Then
   DirVPrecios = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirVPreciosZ" Then
   DirVPreciosZ = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirFlujosEm" Then
   DirFlujosEm = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirFlujosEmZ" Then
   DirFlujosEmZ = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirVAnalitico" Then
   DirVAnalitico = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirVAnaliticoZ" Then
   DirVAnaliticoZ = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirTemp" Then
   DirTemp = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirIndFecha" Then
   DirIndFecha = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirPosMesaD" Then
   DirPosMesaD = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirPosDiv" Then
   DirPosDiv = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirPosSwaps" Then
   DirPosSwaps = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirPosFwdTC" Then
   DirPosFwdTC = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirPosPrimSwaps" Then
   DirPosPrimSwaps = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirPosPrimFwd" Then
   DirPosPrimFwd = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirPosPensiones" Then
   DirPosPensiones = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirPosPensiones2" Then
   DirPosPensiones2 = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirArchBat" Then
   DirArchBat = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirBases" Then
   DirBases = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirReportes" Then
   DirReportes = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirResVaR" Then
   DirResVaR = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "DirWinRAR" Then
   DirWinRAR = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "NomArchRVaR" Then
   NomArchRVaR = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "NomSRVPIP" Then
   NomSRVPIP = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "usersftpPIP" Then
   usersftpPIP = MatParamUsuario(i, 2)
ElseIf MatParamUsuario(i, 1) = "passsftpPIP" Then
   passsftpPIP = MatParamUsuario(i, 2)
End If
Next i
End Sub

Sub CargaEscEstresOra(fecha1, fecha2, nomarch, txtcad)
'primero los cetes con imp
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "cetesimp$", "CETES IMP", txtcad)
'descuento irs
 Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "descirs$", "DESC IRS", txtcad)
'TASA REAL
 Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "rswap$", "CCS UDI-TIIE", txtcad)
'BONDES D
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "bondesd$", "BONDES D", txtcad)
'rep g1
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "repg1$", "REP G1 IMP", txtcad)
'rep b1
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "repb1$", "REP B1", txtcad)
'ipab sem
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "ipabsem$", "IPAB SEM", txtcad)
'bpat
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "bpat$", "BPAT", txtcad)
'ipab
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "ipab$", "IPAB", txtcad)
'ban b1
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "banb1$", "BAN B1", txtcad)
'bonde ls
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "bondels$", "BONDE LS", txtcad)
'ccmid
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "ccmid$", "CCMID", txtcad)
'libor
Call CargaEscenariosExlOra(fecha1, fecha2, nomarch, "libor$", "LIBOR", txtcad)
'todos los indices
Call CargaEscenariosExlOra2(fecha1, fecha2, nomarch, "indices$", txtcad)
End Sub

Sub CargaEscenariosExlOra(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal nomarch As String, ByVal nomtabla As String, ByVal txtconcepto As String, ByVal txtport As String)
Dim noregt As Long
Dim i As Long
Dim j As Long
Dim noreg2 As Long
Dim inicio As Long
Dim fin As Long
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim nodiasf As Long
Dim nocampos As Long
Dim contar As Long
Dim txtfecha As String
Dim txtborra As String
Dim txtcadena As String
Dim tablaescenarios As String

'se debe cargar previamente el portafolio de curvas global
noregt = UBound(MatResFRiesgo, 1)
For i = 1 To noregt
If MatResFRiesgo(i, 1) = txtconcepto Then
 noreg2 = MatResFRiesgo(i, 2)
 inicio = MatResFRiesgo(i, 3)
 fin = inicio + noreg2 - 1
 Exit For
End If
Next i

'primero se lee la tabla de escenarios

Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
Set registros1 = base1.OpenRecordset(nomtabla, dbOpenDynaset, dbReadOnly)
registros1.MoveLast
nodiasf = registros1.RecordCount
nocampos = registros1.Fields.Count
registros1.MoveFirst
ReDim matplz(1 To nocampos - 3) As Integer
For i = 1 To nocampos - 3
 matplz(i) = Val(registros1.Fields(i + 1).Name)
Next i
ReDim mata(1 To nodiasf, 1 To nocampos - 1) As Variant

For i = 1 To nodiasf
 For j = 1 To nocampos - 1
  mata(i, j) = LeerTAccess(registros1, j - 1, i)
 Next j
 registros1.MoveNext
Next i
registros1.Close
base1.Close
'una vez cargada la tabla de escenarios se deben de calcular los escenarios correctos para cada plazo
ReDim matb(1 To nodiasf, 1 To noreg2) As Variant
ReDim matc(1 To nocampos - 3) As propCurva
For i = 1 To nodiasf
For j = 1 To nocampos - 3
 matc(j).valor = mata(i, j + 2)      'factor
 matc(j).plazo = matplz(j)           'plazo
Next j

For j = 1 To noreg2
matb(i, j) = CalculaTasa(matc, MatCaracFRiesgo(inicio + j - 1, 4), 1)
Next j
Next i
'ahora con la matriz de tasas se debe de obtener las tasas para ese dia correspondiente



  contar = 0
  For i = 1 To nodiasf
  If mata(i, 1) >= fecha1 And mata(i, 1) <= fecha2 Then
  For j = 1 To noreg2
  If EsVariableVacia(txtconcepto) Then MsgBox "El concepto esta vacio"
  If mata(i, 1) <= #1/1/2004# Then
  txtfecha = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtborra = "DELETE FROM " & tablaescenarios & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = " & MatCaracFRiesgo(inicio + j - 1, 4) & " AND TIPOESC = '" & txtport & "'"
  ConAdo.Execute txtborra
  txtcadena = "INSERT INTO " & tablaescenarios & " VALUES("
  txtcadena = txtcadena & "'" & txtport & "',"
  txtcadena = txtcadena & txtfecha & ","                    'FECHA
  txtcadena = txtcadena & "'" & txtconcepto & "',"          'CONCEPTO

  txtcadena = txtcadena & MatCaracFRiesgo(inicio + j - 1, 4) & ","          'plazo
  txtcadena = txtcadena & Val(matb(i, j)) & ","             'valor
  txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & Trim(txtconcepto) & Trim(Format(MatCaracFRiesgo(inicio + j - 1, 4), "0000000")) & "')"
  ConAdo.Execute txtcadena
  contar = contar + 1
  DoEvents
  End If
  Next j
  End If
 Next i
End Sub


Sub CargaEscenariosExlOra2(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal nomarch As String, ByVal nomtabla As String, ByVal txtport As String)
'carga escenarios de indices a una tabla de oracle
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim contar As Long
Dim txtfecha As String
Dim txtconcepto As String
Dim txtcadena As String
Dim tablaescenarios As String


Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
Set registros1 = base1.OpenRecordset(nomtabla, dbOpenDynaset, dbReadOnly)
registros1.MoveLast
noreg = registros1.RecordCount
nocampos = registros1.Fields.Count
registros1.MoveFirst
ReDim mata(1 To noreg, 1 To nocampos) As Variant
For i = 1 To noreg
 For j = 1 To nocampos
  mata(i, j) = LeerTAccess(registros1, j - 1, i)
 Next j
 registros1.MoveNext
Next i
registros1.Close
base1.Close
  contar = 0
  For i = 1 To noreg
  If mata(i, 1) >= fecha1 And mata(i, 1) <= fecha2 Then
  txtfecha = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
' cetes 28
  txtconcepto = "SP CETES 28 DIAS"
  ConAdo.Execute "DELETE FROM " & tablaescenarios & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = " & 0 & " AND TIPOESC = '" & txtport & "'"
  txtcadena = "INSERT INTO " & tablaescenarios & " VALUES("
  txtcadena = txtcadena & "'" & txtport & "',"
  txtcadena = txtcadena & txtfecha & ","                    'Fecha
  txtcadena = txtcadena & "'" & txtconcepto & "',"          'Concepto
  txtcadena = txtcadena & 0 & ","                           'plazo
  txtcadena = txtcadena & Val(mata(i, 3)) & ","             'valor
  txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & Trim(txtconcepto) & Trim(Format(0, "0000000")) & "')"
  ConAdo.Execute txtcadena
  contar = contar + 1
' cetes 91
  txtconcepto = "SP CETES 91 DIAS"
  ConAdo.Execute "DELETE FROM " & tablaescenarios & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = " & 0 & " AND TIPOESC = '" & txtport & "'"
  txtcadena = "INSERT INTO " & tablaescenarios & " VALUES("
  txtcadena = txtcadena & "'" & txtport & "',"
  txtcadena = txtcadena & txtfecha & ","                    'FECHA
  txtcadena = txtcadena & "'" & txtconcepto & "',"          'CONCEPTO
  txtcadena = txtcadena & 0 & ","                           'plazo
  txtcadena = txtcadena & Val(mata(i, 4)) & ","             'valor
  txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & Trim(txtconcepto) & Trim(Format(0, "0000000")) & "')"
  ConAdo.Execute txtcadena
  contar = contar + 1
' cetes 182
  txtconcepto = "SP CETES 182 DIAS"
  ConAdo.Execute "DELETE FROM " & tablaescenarios & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = " & 0 & " AND TIPOESC = '" & txtport & "'"
  txtcadena = "INSERT INTO " & tablaescenarios & " VALUES("
  txtcadena = txtcadena & "'" & txtport & "',"
  txtcadena = txtcadena & txtfecha & ","                    'FECHA
  txtcadena = txtcadena & "'" & txtconcepto & "',"          'CONCEPTO
  txtcadena = txtcadena & 0 & ","                           'plazo
  txtcadena = txtcadena & Val(mata(i, 5)) & ","             'valor
  txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & Trim(txtconcepto) & Trim(Format(0, "0000000")) & "')"
  ConAdo.Execute txtcadena
  contar = contar + 1
' tiie 28
  txtconcepto = "TIIE 28"
  ConAdo.Execute "DELETE FROM " & tablaescenarios & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = " & 0 & " AND TIPOESC = '" & txtport & "'"
  txtcadena = "INSERT INTO " & tablaescenarios & " VALUES("
  txtcadena = txtcadena & "'" & txtport & "',"
  txtcadena = txtcadena & txtfecha & ","                    'FECHA
  txtcadena = txtcadena & "'" & txtconcepto & "',"          'CONCEPTO
  txtcadena = txtcadena & 0 & ","                           'plazo
  txtcadena = txtcadena & Val(mata(i, 7)) & ","             'valor
  txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & Trim(txtconcepto) & Trim(Format(0, "0000000")) & "')"
  ConAdo.Execute txtcadena
  contar = contar + 1
' tpfb
  txtconcepto = "TPFB"
  ConAdo.Execute "DELETE FROM " & tablaescenarios & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = " & 0 & " AND TIPOESC = '" & txtport & "'"
  txtcadena = "INSERT INTO " & tablaescenarios & " VALUES("
  txtcadena = txtcadena & "'" & txtport & "',"
  txtcadena = txtcadena & txtfecha & ","                    'FECHA
  txtcadena = txtcadena & "'" & txtconcepto & "',"          'CONCEPTO
  txtcadena = txtcadena & 0 & ","                           'plazo
  txtcadena = txtcadena & Val(mata(i, 8)) & ","             'valor
  txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & Trim(txtconcepto) & Trim(Format(0, "0000000")) & "')"
  ConAdo.Execute txtcadena
  contar = contar + 1
' dolar
  txtconcepto = "DOLAR PIP MD"
  ConAdo.Execute "DELETE FROM " & tablaescenarios & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = " & 0 & " AND TIPOESC = '" & txtport & "'"
  txtcadena = "INSERT INTO " & tablaescenarios & " VALUES("
  txtcadena = txtcadena & "'" & txtport & "',"
  txtcadena = txtcadena & txtfecha & ","                    'FECHA
  txtcadena = txtcadena & "'" & txtconcepto & "',"          'CONCEPTO
  txtcadena = txtcadena & 0 & ","                           'plazo
  txtcadena = txtcadena & Val(mata(i, 10)) & ","            'valor
  txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & Trim(txtconcepto) & Trim(Format(0, "0000000")) & "')"
  ConAdo.Execute txtcadena
  contar = contar + 1
' yen
  txtconcepto = "YEN PIP"
  ConAdo.Execute "DELETE FROM " & tablaescenarios & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = " & 0 & " AND TIPOESC = '" & txtport & "'"
  txtcadena = "INSERT INTO " & tablaescenarios & " VALUES("
  txtcadena = txtcadena & "'" & txtport & "',"
  txtcadena = txtcadena & txtfecha & ","                    'FECHA
  txtcadena = txtcadena & "'" & txtconcepto & "',"          'CONCEPTO
  txtcadena = txtcadena & 0 & ","                           'plazo
  txtcadena = txtcadena & Val(mata(i, 11)) & ","            'valor
  txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & Trim(txtconcepto) & Trim(Format(0, "0000000")) & "')"
  ConAdo.Execute txtcadena
  contar = contar + 1
' euro
  txtconcepto = "EURO PIP"
  ConAdo.Execute "DELETE FROM " & tablaescenarios & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = " & 0 & " AND TIPOESC = '" & txtport & "'"
  txtcadena = "INSERT INTO " & tablaescenarios & " VALUES("
  txtcadena = txtcadena & "'" & txtport & "',"
  txtcadena = txtcadena & txtfecha & ","                    'FECHA
  txtcadena = txtcadena & "'" & txtconcepto & "',"          'CONCEPTO
  txtcadena = txtcadena & 0 & ","                           'plazo
  txtcadena = txtcadena & Val(mata(i, 12)) & ","            'valor
  txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & Trim(txtconcepto) & Trim(Format(0, "0000000")) & "')"
  ConAdo.Execute txtcadena
  contar = contar + 1
  
' LIBRA
'  txtconcepto = "LIBRA PIP"
'  conAdo.Execute "DELETE FROM " & TablaEscenarios & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = " & 0 & " AND TIPOESC = '" & txtport & "'"
  'txtcadena = "INSERT INTO " & TablaEscenarios & " VALUES("
  'txtcadena = txtcadena & "'" & txtport & "',"
  'txtcadena = txtcadena & txtfecha & ","                     'FECHA
  'txtcadena = txtcadena & "'" & txtconcepto & "',"           'CONCEPTO
  'txtcadena = txtcadena & 0 & ","                            'plazo
  'txtcadena = txtcadena & Val(mata(i, 13)) & ","             'valor
  'txtcadena = txtcadena & "'" & CLng(mata(i, 1)) & Trim(txtconcepto) & Trim(Format(0, "0000000")) & "')"
  'conAdo.Execute txtcadena
  contar = contar + 1
End If
 Next i
 'MsgBox contar & " registros"

End Sub

Sub ImpPosFwdArch(ByVal fecha As Date, ByVal nomarch As String, ByVal nompos As String, ByVal coper As String, ByVal toper As Integer, ByVal intencion As String, ByRef nr1 As Integer)
Dim sihayarch As Boolean
Dim Base As DAO.Database
Dim registros As DAO.recordset
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim txtfecha As String
Dim txtcadena As String


sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then

 Set Base = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
 Set registros = Base.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)
If registros.RecordCount <> 0 Then
registros.MoveLast
noreg = registros.RecordCount
nocampos = registros.Fields.Count
ReDim mata(1 To noreg, 1 To nocampos) As Variant
registros.MoveFirst
For i = 1 To noreg
    For j = 1 To nocampos
        mata(i, j) = LeerTAccess(registros, j - 1, i)
    Next j
If mata(i, 2) = "largo" Then
 mata(i, 2) = 1
Else
 mata(i, 2) = 4
End If
registros.MoveNext
Next i
registros.Close
Base.Close
coper = mata(1, 1)
toper = mata(1, 7)
intencion = mata(1, 10)

txtfecha = Format(fecha, "dd/mm/yyyy")
ConAdo.Execute "DELETE FROM " & TablaPosFwd & " WHERE FECHAREG = '" & nompos & "'"
nr1 = 0
For i = 1 To noreg
 nr1 = nr1 + 1
 txtcadena = "INSERT INTO " & TablaPosFwd & " VALUES("
 txtfecha = Format(fecha, "dd/mm/yyyy")
 txtcadena = txtcadena & "'" & nompos & "',"        'nombre de la posicion
 txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtcadena = txtcadena & "" & txtfecha & ","         'FECHA DE LA POSICION
 txtcadena = txtcadena & "'" & mata(i, 1) & "',"    'clave de la operacion
 txtcadena = txtcadena & "" & mata(i, 2) & ","      'largo o corto
 txtcadena = txtcadena & "" & mata(i, 3) & ","      'no titulos
 txtcadena = txtcadena & "" & mata(i, 4) & ","      'monto nocional
 txtfecha = "to_date('" & Format(mata(i, 5), "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtcadena = txtcadena & txtfecha & ","             'fecha inicio
 txtfecha = "to_date('" & Format(mata(i, 6), "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtcadena = txtcadena & txtfecha & ","             'fecha vencimiento
 txtfecha = "to_date('" & Format(mata(i, 6), "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtcadena = txtcadena & txtfecha & ","                          'fecha liquidacion
 txtcadena = txtcadena & "'" & mata(i, 7) & "',"                 'tipo de fwd
 txtcadena = txtcadena & CLng(mata(i, 8)) & ","                   'plazo fwd
 txtcadena = txtcadena & "'" & mata(i, 10) & "',"                 'intencion
 txtcadena = txtcadena & "" & mata(i, 9) & ")"                    'tipo pactado
 ConAdo.Execute txtcadena
Next i
MsgBox "Se exportaron " & nr1 & " registros de forwards"
End If
End If

End Sub


Function ObFechasPos(ByVal txtport As String)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset

txtfiltro2 = "SELECT FECHA_PORT FROM " & TablaPortPosicion & "  WHERE PORTAFOLIO = '" & txtport & "' GROUP BY FECHA_PORT ORDER BY FECHA_PORT"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 1) As Date
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields(0)
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Date
End If
ObFechasPos = mata
End Function


Function ObPortPosSim()
Dim mata() As Variant
Dim mata1() As Variant
Dim mata2() As Variant
Dim mata3() As Variant
Dim mata4() As Variant


    ReDim mata(0 To 0, 0 To 0) As Variant
    mata1 = NomPosMD(4)       'mercado de dinero
    mata2 = NomPosDiv()       'mesa de cambios
    mata3 = NomPosSwaps()     'swaps
    mata4 = NomPosFwd()       'forwards
    mata = UnirTablas(mata, mata1, 1)
    mata = UnirTablas(mata, mata2, 1)
    mata = UnirTablas(mata, mata3, 1)
    mata = UnirTablas(mata, mata4, 1)
    mata = ObtFactUnicos(mata, 1)
    ObPortPosSim = mata

End Function

Sub CrearCurvasCSVIKOS(ByVal fecha As Date, ByVal direc1 As String, ByVal direc2 As String, ByRef txtmsg As String, ByRef exito As Boolean)
'esta rutina crea los archivos que se van a subir al sistema ikos para la carga de curvas
Dim txtfecha As String
Dim nomarch As String
Dim sihayarch1 As Boolean
Dim notablas As Integer
Dim cont1 As Integer
Dim inbucle As TableDef
Dim registros1 As DAO.recordset
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim txtsalida As String
Dim exitoarch As Boolean

exito = False
txtfecha = Format(fecha, "yyyymmdd")
nomarch = direc1 & "\CURVAS" & Format(fecha, "yyyymmdd") & ".XLS"
sihayarch1 = VerifAccesoArch(nomarch)
If sihayarch1 Then
Dim base1 As DAO.Database
    Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
'se verifica cual es el nombre de la tabla que contiene los datos
    notablas = base1.TableDefs.Count
    ReDim matnom(1 To notablas) As String
    cont1 = 1
    For Each inbucle In base1.TableDefs
     matnom(cont1) = inbucle.Name
     cont1 = cont1 + 1
    Next inbucle
    Set registros1 = base1.OpenRecordset(matnom(1), dbOpenDynaset, dbReadOnly)
    If registros1.RecordCount <> 0 Then
       registros1.MoveLast
       noreg = registros1.RecordCount
       registros1.MoveFirst
       nocampos = registros1.Fields.Count
  ReDim matcrv(1 To noreg, 1 To nocampos) As Variant
  ReDim matcamp(1 To nocampos) As Variant
       For i = 1 To nocampos
           matcamp(i) = registros1.Fields(i - 1).Name
       Next i
       For i = 1 To noreg
           For j = 1 To nocampos
               matcrv(i, j) = LeerTAccess(registros1, j - 1, i)
           Next j
           registros1.MoveNext
       Next i
End If
registros1.Close
base1.Close
Call VerificarSalidaArchivo(direc2 & "\CURVAS" & Format(fecha, "YYYYMMDD") & ".CSV", 2, exitoarch)
If exitoarch Then
txtsalida = ""
For i = 1 To nocampos
If i <> nocampos Then
 txtsalida = txtsalida & i & ","
Else
 txtsalida = txtsalida & i
End If
Next i
Print #2, txtsalida
txtsalida = ""
For i = 1 To nocampos
If i <> nocampos Then
txtsalida = txtsalida & matcamp(i) & ","
Else
txtsalida = txtsalida & matcamp(i)
End If
Next i
Print #2, txtsalida
For i = 1 To noreg
    txtsalida = ""
    For j = 1 To nocampos
    If j <> nocampos Then
       txtsalida = txtsalida & matcrv(i, j) & ","
    Else
       txtsalida = txtsalida & matcrv(i, j)
    End If
    Next j
Print #2, txtsalida
AvanceProc = i / noreg
MensajeProc = "Creando el archivo CSV " & Format(AvanceProc, "##0.00 %")
Next i
Close #2
exito = True
MensajeProc = "Se creo el archivo de curvas csv del dia " & fecha
txtmsg = "El proceso finalizo correctamente"
End If
Else
  MensajeProc = "No hay acceso a de curvas del día " & fecha
  txtmsg = MensajeProc
  exito = False
End If
End Sub

Sub CrearArchivoCSVC(ByVal fecha As Date, ByVal direc1 As String, ByVal direc2 As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim nomarch As String
Dim sihayarch1 As Boolean
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim inbucle As TableDef
Dim cont1 As Integer
Dim noreg As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim txtsalida As String
Dim notablas As Long
Dim exitoarch As Boolean

'esta rutina crea los archivos que se van a subir al dwh
exito = False
txtfecha = Format(fecha, "yyyymmdd")
nomarch = direc1 & "\CURVAS" & Format(fecha, "yyyymmdd") & ".XLS"
sihayarch1 = VerifAccesoArch(nomarch)
If sihayarch1 Then

Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
'se verifica cual es el nombre de la tabla que contiene los datos
    notablas = base1.TableDefs.Count
    ReDim matnom(1 To notablas) As String
    cont1 = 1
    For Each inbucle In base1.TableDefs
     matnom(cont1) = inbucle.Name
     cont1 = cont1 + 1
    Next inbucle
 Set registros1 = base1.OpenRecordset(matnom(1), dbOpenDynaset, dbReadOnly)
 If registros1.RecordCount <> 0 Then
  registros1.MoveLast
  noreg = registros1.RecordCount
  registros1.MoveFirst
  nocampos = registros1.Fields.Count
  ReDim matcrv(1 To noreg, 1 To nocampos) As Variant
  ReDim matcamp(1 To nocampos) As Variant
  For i = 1 To nocampos
  matcamp(i) = registros1.Fields(i - 1).Name
  Next i
For i = 1 To noreg
 For j = 1 To nocampos
  matcrv(i, j) = LeerTAccess(registros1, j - 1, i)
 Next j
 registros1.MoveNext
Next i
End If
registros1.Close
base1.Close
Call VerificarSalidaArchivo(direc2 & "\CURVAS" & Format(fecha, "YYYYMMDD") & ".CSV", 2, exitoarch)
If exitoarch Then
txtsalida = ""
For i = 1 To nocampos
If i <> nocampos Then
   txtsalida = txtsalida & matcamp(i) & ","
Else
   txtsalida = txtsalida & matcamp(i)
End If
Next i
Print #2, txtsalida
For i = 1 To noreg
txtsalida = ""
For j = 1 To nocampos
If j <> nocampos Then
txtsalida = txtsalida & matcrv(i, j) & ","
Else
txtsalida = txtsalida & matcrv(i, j)
End If
Next j
Print #2, txtsalida
MensajeProc = "Creando el archivo de curvas CSV"
Next i
Close #2
exito = True
MensajeProc = "Se creo el archivo csv de curvas del dia " & fecha
End If
Else
   MensajeProc = "No hay acceso a de curvas del día"
   Call MostrarMensajeSistema(MensajeProc, frmProgreso.Label2, 1, Date, Time, NomUsuario)
End If
End Sub

Sub GenProcesosDia(ByVal dtfecha As Date, ByRef matt() As Variant, ByVal opcion As Integer)
'esta rutina genera las tareas a ejecutar para la fecha indicada
'estas se guardaran en un archivo que le indicada al sistema que esta pendiente de realizar
'
    Dim txtfecha As String, txtfiltro As String, txtcadena As String
    Dim txtborra As String
    Dim i As Integer, j As Integer
    Dim noreg As Integer
    Dim noreg1 As Integer
    Dim siejec As Boolean
    Dim contar As Long
    Dim txttabla As String
    If opcion = 1 Then
       txttabla = TablaProcesos1
    ElseIf opcion = 2 Then
       txttabla = TablaProcesos2
    End If
    If dtfecha = PBD1(Date, 1, "MX") Or dtfecha = Date Then
       siejec = False
    Else
       siejec = True
    End If
    noreg = UBound(matt, 1)
    txtfecha = "to_date('" & Format$(dtfecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    For i = 1 To noreg
        If frmProcesos.MSFlexGrid1.TextMatrix(i, 2) = "S" Then
           contar = contar + 1
           txtborra = "DELETE FROM " & txttabla & " WHERE FECHAP = " & txtfecha & " AND ID_TAREA = " & matt(i, 1)
           ConAdo.Execute txtborra
           txtcadena = "INSERT INTO " & txttabla & " VALUES("
           txtcadena = txtcadena & matt(i, 1) & ","                                            'id de tarea (unico)
           txtcadena = txtcadena & "'" & matt(i, 2) & "',"                                     'id del proceso
           txtcadena = txtcadena & "'" & matt(i, 3) & "',"                                     'descripcion del proceso
           For j = 1 To 10
               txtcadena = txtcadena & "'" & ReemplazaVacioValor(matt(i, 2 * j + 2), "") & "',"                 'descrip param
               txtcadena = txtcadena & "'" & ReempCadParam(ReemplazaVacioValor(matt(i, 2 * j + 3), "")) & "',"  'parametro
           Next j
           txtcadena = txtcadena & txtfecha & ","                                              'fecha a la que corresponde el proceso
           txtcadena = txtcadena & "null,"                                                     'fecha de inicio tarea
           txtcadena = txtcadena & "null,"                                                     'hora inicio tarea
           txtcadena = txtcadena & "null,"                                                     'fecha de fin tarea
           txtcadena = txtcadena & "null,"                                                     'hora fin tarea
           txtcadena = txtcadena & "'N',"                                                      'bloqueada
           txtcadena = txtcadena & "'N',"                                                      'finalizada
           txtcadena = txtcadena & "null,"                                                     'comentario
           txtcadena = txtcadena & "null,"                                                     'usuario que realizo el proceso
           txtcadena = txtcadena & "null)"                                                     'direccion ip
           ConAdo.Execute txtcadena
        End If
    Next i
    MensajeProc = "Se generaron " & noreg & " procesos para el dia " & dtfecha

End Sub

Function ReempCadParam(texto)
If texto = "sql_DirCurvasCSV1" Then
   ReempCadParam = DirCurvasCSV1
ElseIf texto = "sql_DirCurvas" Then
   ReempCadParam = DirCurvas
ElseIf texto = "sql_DirCurvasZ" Then
   ReempCadParam = DirCurvasZ
ElseIf texto = "sql_DirVPrecios" Then
   ReempCadParam = DirVPrecios
ElseIf texto = "sql_DirVPreciosZ" Then
   ReempCadParam = DirVPreciosZ
ElseIf texto = "sql_DirFlujosEm" Then
   ReempCadParam = DirFlujosEm
ElseIf texto = "sql_DirFlujosEmZ" Then
   ReempCadParam = DirFlujosEmZ
ElseIf texto = "sql_DirTemp" Then
   ReempCadParam = DirTemp
ElseIf texto = "sql_DirPosDiv" Then
   ReempCadParam = DirPosDiv
  
Else
   ReempCadParam = texto
End If
End Function



Function GenListaProc(ByVal fecha As Date) As Variant()
Dim i As Integer
For i = 1 To UBound(MatCatProcesos, 1)

Next i
End Function

Sub GuardaResVaRIKOS2(ByVal fecha As Date, ByRef obj1 As ADODB.Connection, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtport As String
Dim fecha1 As Date
Dim txtfecha As String
Dim i As Integer
Dim noreg1 As Integer
Dim txtborra As String
Dim txtcadena As String

txtport = txtportCalc1
fecha1 = FBD(fecha, "MX")
'rutina que guarda los resultados del calculo del VaR
'fecha    fecha a la que se actualizara la informacion de VaR
'dtfechalim   fecha de los limites de VaR
Dim matport(1 To 4) As String
Dim matv(1 To 4) As Double
Dim matlim(1 To 4) As Double
Dim matpor(1 To 4) As Double
Dim txtfiltro1 As String
Dim txtfiltro2 As String

matport(1) = txtportBanobras
matport(2) = "DERIVADOS DE NEGOCIACION"
matport(3) = "DERIVADOS ESTRUCTURALES"
matport(4) = "DERIVADOS NEGOCIACION RECLASIFICACION"
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
For i = 1 To 4
   matv(i) = -LeerCVaRHist(fecha, txtport, matport(i), 0.03, 500, 1)
Next i
CapitalNeto = DevLimitesVaR(fecha, MatCapitalSist, "CAPITAL NETO B") * 1000000
If CapitalNeto <> 0 Then
      matpor(1) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR CON")
      matpor(2) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV")
      matpor(3) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV EST")
      matpor(4) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV10")
      For i = 1 To 4
          matlim(i) = CapitalNeto * matpor(i)
      Next i
      txtborra = "DELETE FROM " & TablaVaRIKOS & " WHERE FECHAOPER = " & Format(fecha1, "YYYYMMDD")
      obj1.Execute txtborra
      txtcadena = "INSERT INTO " & TablaVaRIKOS & " VALUES("
      txtcadena = txtcadena & Format(fecha1, "YYYYMMDD") & ","   'la fecha de la posicion
      txtcadena = txtcadena & "0,"                               'el no de swap
      txtcadena = txtcadena & matv(1) & ","                      'el VaR consolidado
      txtcadena = txtcadena & matv(2) & ","                      'el VaR de la posicion de Derivados
      txtcadena = txtcadena & matv(3) & ","                      'VaR Derivados estructurales
      txtcadena = txtcadena & matv(4) & ","                      'VaR Derivados recalsificacion
      txtcadena = txtcadena & matlim(1) & ","                    'lim var global
      txtcadena = txtcadena & matlim(2) & ","                    'lim var derivados
      txtcadena = txtcadena & matlim(3) & ","                    'lim var estructurales
      txtcadena = txtcadena & matlim(4) & ")"                    'lim var reclasifiacion
      obj1.Execute txtcadena
      MensajeProc = "Exportando el VaR a la tabla " & TablaVaRIKOS & " para la fecha " & fecha
      MensajeProc = "Se exporto el VaR a la tabla " & TablaVaRIKOS & " para la fecha " & fecha
      exito = True
      txtmsg = "El proceso finalizo correctamente"
Else
   MensajeProc = "No hay capital neto en el sistema para esta fecha " & fecha
   txtmsg = MensajeProc
   exito = False
End If
End Sub

Sub GuardaResCVaRIKOS3(ByVal fecha As Date, ByRef mata() As Variant, ByRef matport() As String, ByRef obj1 As ADODB.Connection, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtport As String
Dim fecha1 As Date
Dim txtfecha As String
Dim i As Integer
Dim noreg1 As Integer
Dim txtborra As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

txtport = "CONSOLIDADO ID"
fecha1 = FBD(fecha, "MX")
'rutina que guarda los resultados del calculo del VaR
'fecha    fecha a la que se actualizara la informacion de VaR
'dtfechalim   fecha de los limites de VaR
Dim matv(1 To 4) As Double
Dim matlim(1 To 4) As Double
Dim matpor(1 To 4) As Double
Dim txtfiltro1 As String
Dim txtfiltro2 As String

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
For i = 1 To UBound(matport)
    txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & matport(i) & "'"
    txtfiltro2 = txtfiltro2 & " AND TVAR = 'CVARH'"
    txtfiltro2 = txtfiltro2 & " AND NOESC = 500"
    txtfiltro2 = txtfiltro2 & " AND HTIEMPO = 1"
    txtfiltro2 = txtfiltro2 & " AND NCONF = 0.03"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg1 = rmesa.Fields(0)
    rmesa.Close
    If noreg1 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       matv(i) = -rmesa.Fields("VALOR")
       rmesa.Close
    Else
       matv(i) = 0
    End If
Next i
CapitalNeto = DevLimitesVaR(fecha, MatCapitalSist, "CAPITAL NETO B") * 1000000
If CapitalNeto <> 0 Then
      matpor(1) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR CON")
      matpor(2) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV")
      matpor(3) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV EST")
      matpor(4) = DevLimitesVaR(fecha, MatCapitalSist, "LIM VAR DERIV10")
      If matpor(1) <> 0 And matpor(2) <> 0 And matpor(3) <> 0 And matpor(4) <> 0 Then
         For i = 1 To 4
             matlim(i) = CapitalNeto * matpor(i)
         Next i
         txtborra = "DELETE FROM " & TablaVaRIKOS & " WHERE FECHAOPER = " & Format(fecha, "YYYYMMDD")
         obj1.Execute txtborra
         For i = 1 To UBound(mata, 1)
            txtcadena = "INSERT INTO " & TablaVaRIKOS & " VALUES("
            txtcadena = txtcadena & Format(fecha, "YYYYMMDD") & ","    'la fecha del calculo
            txtcadena = txtcadena & mata(i, 1) & ","                   'el no de swap
            txtcadena = txtcadena & matv(1) & ","                      'el VaR consolidado
            txtcadena = txtcadena & matv(2) & ","                      'el VaR de la posicion de Derivados
            txtcadena = txtcadena & matv(3) & ","                      'VaR Derivados estructurales
            txtcadena = txtcadena & matv(4) & ","                      'VaR Derivados recalsificacion
            txtcadena = txtcadena & matlim(1) & ","                    'lim var global
            txtcadena = txtcadena & matlim(2) & ","                    'lim var derivados
            txtcadena = txtcadena & matlim(3) & ","                    'lim var estructurales
            txtcadena = txtcadena & matlim(4) & ")"                    'lim var reclasifiacion
            obj1.Execute txtcadena
            MensajeProc = "Exportando el VaR a la tabla " & TablaVaRIKOS & " para la fecha " & fecha
         Next i
         exito = True
         txtmsg = "El proceso finalizo correctamente"
      Else
        exito = False
        txtmsg = "Alguno de los límites de CVaR es nulo"
      End If
Else
   MensajeProc = "No hay capital neto en el sistema para esta fecha " & fecha
   txtmsg = MensajeProc
   exito = False
   MsgBox MensajeProc
End If
End Sub

Sub CalculoEfRetroSwap(ByVal fecha1 As Date, _
                       ByVal fecha2 As Date, _
                       ByRef matpos() As propPosRiesgo, _
                       ByRef matposmd() As propPosMD, _
                       ByRef matposdiv() As propPosDiv, _
                       ByRef matposswaps() As propPosSwaps, _
                       ByRef matposfwd() As propPosFwd, _
                       ByRef matflswap() As estFlujosDeuda, _
                       ByRef matposdeuda() As propPosDeuda, _
                       ByRef matfldeuda() As estFlujosDeuda, _
                       ByRef matresef1() As Variant, _
                       ByRef matrelac() As propRelSwapPrim)

Dim mrvalflujo() As resValFlujo
Dim exito As Boolean
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim matpos1() As New propPosRiesgo
Dim matflswap1() As New estFlujosDeuda
Dim matflujos2() As New estFlujosDeuda
Dim matflujos3() As New estFlujosDeuda
Dim matflujos4() As New estFlujosDeuda
Dim MDblTasas1() As Double
Dim MatTasas2() As Double
Dim mcurvat1() As Variant
Dim mcurvat2() As Variant
Dim parval As ParamValPos
Dim contar As Long
Dim noreg As Long
Dim norelac As Long
Dim i As Long
Dim j As Long
Dim indice As Long
Dim matcoperacion() As Variant
Dim nofval As Long
Dim contreg As Long
Dim nregcalefec As Long
Dim indice0 As Long
Dim matv() As New resValIns
Dim contarf As Long
Dim indswap As Long
Dim indppact As Long
Dim indpppas As Long
Dim exito3 As Boolean
Dim txtmsg3 As String

'matpos es la matriz con la posicion de swaps
'matrelac es la relacion de los swaps con sus posiciones primarias
'columnas de la matriz matresef1
' 1 Clave de operación



'se debe de saber que tipo de operaciones y como empatan los flujos
contar = 0
'se toma la matriz de riesgo
noreg = UBound(matpos, 1)
norelac = UBound(matrelac, 1)    'aqui se indica cuantos calculo de efectividad se haran
ReDim matarchtext(1 To 27, 1 To 1) As Variant
'se ordena la posición por el tipo de titulo
'para facilitar el calculo de las eficiencias
'se obtienen las emisiones de los swaps
'matcoperacion = ObtFactUnicos(matpos, 3)
'se cuenta el no de emisiones en la posicion
ValExacta = True
nofval = 2
'lo primero es establecer las fechas en las que se va a valuar el swap
'se pone primero la fecha en cada caso
'matvalores se dimensiona del no de patas al no de fechas de valuacion
ReDim matfecha(1 To nofval) As Date
ReDim matresef1(1 To 27, 1 To 1) As Variant
matfecha(1) = fecha1
matfecha(2) = fecha2
'se obtienen los factores adicionales para el calculo de la efectividad
contreg = 0
'se cambia la estrategia, se filtra la posicion por cada emision y se calcula la
'efectividad

mcurvat1 = LeerCurvaCompleta(matfecha(1), exito1)
mcurvat2 = LeerCurvaCompleta(matfecha(2), exito2)
MDblTasas1 = CargaFR1Dia(matfecha(1), exito)
MatTasas2 = CargaFR1Dia(matfecha(2), exito)
contar = 0
For i = 1 To noreg               'inicio de la relacion
    If matpos(i).C_Posicion = ClavePosDeriv Then
       indice = 0
       For j = 1 To UBound(matrelac, 1)
           If matpos(i).c_operacion = matrelac(j).coperacion Then
              indice = j
              Exit For
           End If
       Next j
       If indice <> 0 Then
          ReDim TXTEM(1 To 3, 1 To 1) As Variant
          TXTEM(1, 1) = matrelac(indice).coperacion                         'Clave de operación
          TXTEM(2, 1) = matrelac(indice).c_ppactiva                         'clave de la posicion primaria activa
          TXTEM(3, 1) = matrelac(indice).c_pppasiva                         'clave de la posicion primaria pasiva
          matpos1 = FiltrarPosG1(matpos, TXTEM)                             'la posicion conjunta
'se procede a crear la posicion con la que se calculara la efectividad
          nregcalefec = UBound(matpos1, 1)
          If nregcalefec <> 0 Then
             If nregcalefec <> 2 And (matrelac(indice).t_efect = 1 Or matrelac(indice).t_efect = 2) Then
                MensajeProc = "no se puede calcular la efectividad de la operacion " & matrelac(indice, 1) & ". Faltan datos"
                MsgBox MensajeProc
             ElseIf nregcalefec <> 3 And matrelac(indice).t_efect = 3 Then
                MensajeProc = "no se puede calcular la efectividad de la operacion " & matrelac(indice).coperacion & ". Faltan datos"
                MsgBox MensajeProc
             Else
                contar = contar + 1
                ReDim Preserve matresef1(1 To 27, 1 To contar) As Variant
                ReDim res_efi(1 To nofval) As New resEficSwap
                matresef1(1, contar) = TXTEM(1, 1)                          'Clave de operación
                matresef1(2, contar) = TXTEM(2, 1)                          'posicion primaria activa
                matresef1(3, contar) = TXTEM(3, 1)                          'posicion primaria pasiva
                indice0 = matpos(i).IndPosicion
                matresef1(4, contar) = matposswaps(indice0).ClaveProdSwap   'clave de tipo de producto
                matresef1(5, contar) = matposswaps(indice0).FvencSwap
                matresef1(6, contar) = matrelac(indice).t_efect             'tipo de eficiencia de cob
                If matresef1(6, contar) = 1 Then
                    indswap = 1: indpppas = 2
                ElseIf matresef1(6, contar) = 2 Then
                    indswap = 1: indppact = 2
                ElseIf matresef1(6, contar) = 3 Then
                    indswap = 1: indppact = 2: indpppas = 3
                End If
                If EsVariableVacia(matrelac(indice).t_efect) Then
                   MensajeProc = "No se como calcular la eficiencia del instrumento " & matrelac(indice, 1)
                   MsgBox MensajeProc
                End If
                Set parval = DeterminaPerfilVal("VALUACION")
                parval.sicalcdur = False
                parval.sicalcdv01 = False
                For j = 1 To nofval
'se calculan las valuaciones en la fecha j de la posicion i
                    If j = 1 Then
                       matv = CalcValuacion(matfecha(j), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap1, matposdeuda, matflujos3, MDblTasas1, mcurvat1, parval, mrvalflujo, txtmsg3, exito3)
                    Else
                       matv = CalcValuacion(matfecha(j), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap1, matposdeuda, matflujos3, MatTasas2, mcurvat2, parval, mrvalflujo, txtmsg3, exito3)
                    End If
                    res_efi(j).val_efect = 0
                    Call CalculosEfecSwap(matrelac(indice).t_efect, res_efi, matv, j, indswap, indppact, indpppas)
                Next j
                If res_efi(2).val_efect < 0.8 Or res_efi(2).val_efect > 1.25 Then
                   MsgBox "no es eficiente la operacion " & matresef1(1, contar) & " " & res_efi(2).val_efect
                End If
                matresef1(27, contar) = Format(res_efi(2).val_efect * 100, "##0.00")
                'matresef1(27, i) = 100
                MensajeProc = "Procesando la operacion " & matresef1(1, contar)
             End If
             DoEvents
          End If
       Else
          MsgBox "no se encontro la relacion de posicion primaria en la tabla de datos de la operacion " & matpos(i).c_operacion
       End If
    End If
Next i

'primero se tiene que obtener los swaps que estan presentes en el analisis de la eficiencia
'aguas esto solo aplica para una eficiencia de 1 periodo
matresef1 = MTranV(matresef1)
'se agregar los resultados al archivo de texto
contarf = UBound(matarchtext, 2)
ReDim Preserve matarchtext(1 To 27, 1 To contarf + norelac) As Variant
For i = 1 To UBound(matresef1, 1) - 2
For j = 1 To 27
 matarchtext(j, i + contarf) = matresef1(i, j)
Next j
Next i
Call GuardaMatArchTexto(MTranV(matarchtext), DirResVaR & "\Eficiencia retro Swaps " & Format(fecha2, "yyyy-mm-dd") & ".txt")
'se ponen los resultados en pantalla
End Sub

Sub CalculoEfRetroSwapAct(ByVal fecha1 As Date, _
                          ByVal fecha2 As Date, _
                          ByVal txtport As String, _
                          ByRef txtmsg As String, _
                          ByRef final As Boolean, _
                          ByRef exito As Boolean)

Dim mrvalflujo() As resValFlujo
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim txtmsg0 As String
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim parval As ParamValPos
Dim contar As Long
Dim noreg As Long
Dim norelac As Long
Dim i As Long
Dim j As Long
Dim indice As Long
Dim matcoperacion() As Variant
Dim nofval As Long
Dim contreg As Long
Dim nregcalefec As Long
Dim indice0 As Long
Dim matv() As New resValIns
Dim contarf As Long
Dim indswap As Integer
Dim indpact As Integer
Dim fecha3 As Date
Dim exito3 As Boolean
Dim txtmsg2 As String
Dim txtmsg3 As String

final = False
exito = False
mattxt = CrearFiltroPosPort(fecha2, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
'se debe de saber que tipo de operaciones y como empatan los flujos
noreg = UBound(matpos, 1)
If noreg <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   ValExacta = True
   nofval = 2
   ReDim matfecha(1 To nofval) As Date
   matfecha(1) = fecha1
   matfecha(2) = fecha2
   fecha3 = PBD1(fecha2, 1, "MX")
   If matfecha(1) <> fechaFactR1 Then
      MatCurvasT1 = LeerCurvaCompleta(matfecha(1), exito1)
      MatFactRiesgo1 = CargaFR1Dia(matfecha(1), exito)
      fechaFactR1 = matfecha(1)
   End If
   If fecha3 <> fechaFactR2 Then
      MatCurvasT2 = LeerCurvaCompleta(fecha3, exito2)
      MatFactRiesgo2 = CargaFR1Dia(fecha3, exito)
      fechaFactR2 = fecha3
   End If
   Call DetermSwapPrimAct(matpos, indswap, indpact)
   If indswap <> 0 And indpact <> 0 Then
      ReDim res_efi(1 To nofval) As New resEficSwap
      Set parval = DeterminaPerfilVal("VALUACION")
      parval.si_int_flujos = True
      parval.sicalcdur = False
      parval.sicalcdv01 = False
      For j = 1 To nofval
          If j = 1 Then
             matv = CalcValuacion(matfecha(j), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactRiesgo1, MatCurvasT1, parval, mrvalflujo, txtmsg3, exito3)
          Else
             matv = CalcValuacion(matfecha(j), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactRiesgo2, MatCurvasT2, parval, mrvalflujo, txtmsg3, exito3)
          End If
          res_efi(j).val_efect = 0
          Call CalculosEfecSwapAct(res_efi, matv, j, indswap, indpact)
      Next j
      Call GuardarResEfRetroSwap2(fecha2, matpos, matposswaps, indswap, 2, res_efi(2).val_efect)
      Call GuardaResEficRetro(fecha2, matpos(1).c_operacion, res_efi(2).val_efect, ConAdo)
      If res_efi(2).val_efect < 0.8 Or res_efi(2).val_efect > 1.25 Then
         txtmsg = "no es eficiente la operacion " & matpos(1).c_operacion & " " & res_efi(2).val_efect
         final = True
         exito = False
      Else
         txtmsg = "El proceso finalizo correctamente"
         final = True
         exito = True
      End If
   Else
      txtmsg = "No hay los registros que se esperaban " & matpos(1).c_operacion
      final = True
     exito = False
   End If
End If
End Sub

Sub CalculoEfRetroSwapPas(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef final As Boolean, ByRef exito As Boolean)
Dim mattxt() As String
Dim mrvalflujo() As resValFlujo
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim parval As ParamValPos
Dim noreg As Long
Dim norelac As Long
Dim i As Long
Dim j As Long
Dim indice As Long
Dim nofval As Long
Dim contreg As Long
Dim nregcalefec As Long
Dim indice0 As Long
Dim matv() As New resValIns
Dim contarf As Long
Dim indswap As Integer
Dim indppas As Integer
Dim fecha3 As Date
Dim exito3 As Boolean
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim txtmsg3 As String

final = False
mattxt = CrearFiltroPosPort(fecha2, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
'se debe de saber que tipo de operaciones y como empatan los flujos
noreg = UBound(matpos, 1)
If noreg <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
'se debe de saber que tipo de operaciones y como empatan los flujos
   ValExacta = True
   nofval = 2
   ReDim matfecha(1 To nofval) As Date
   matfecha(1) = fecha1
   matfecha(2) = fecha2
   fecha3 = PBD1(fecha2, 1, "MX")
   If matfecha(1) <> fechaFactR1 Then
      MatCurvasT1 = LeerCurvaCompleta(matfecha(1), exito1)
      MatFactRiesgo1 = CargaFR1Dia(matfecha(1), exito)
      fechaFactR1 = matfecha(1)
   End If
   If fecha3 <> fechaFactR2 Then
      MatCurvasT2 = LeerCurvaCompleta(fecha3, exito2)
      MatFactRiesgo2 = CargaFR1Dia(fecha3, exito)
      fechaFactR2 = fecha3
   End If
   Call DetermSwapPrimPas(matpos, indswap, indppas)
   If indswap <> 0 And indppas <> 0 Then
      Set parval = DeterminaPerfilVal("VALUACION")
      parval.si_int_flujos = True
      ReDim res_efi(1 To nofval) As New resEficSwap
      parval.sicalcdur = False
      parval.sicalcdv01 = False
      For j = 1 To nofval
          If j = 1 Then
             matv = CalcValuacion(matfecha(j), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactRiesgo1, MatCurvasT1, parval, mrvalflujo, txtmsg3, exito3)
          Else
             matv = CalcValuacion(matfecha(j), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactRiesgo2, MatCurvasT2, parval, mrvalflujo, txtmsg3, exito3)
          End If
          res_efi(j).val_efect = 0
          Call CalculosEfecSwapPas(res_efi, matv, j, indswap, indppas)
      Next j
      Call GuardarResEfRetroSwap2(fecha2, matpos, matposswaps, indswap, 1, res_efi(2).val_efect)
      Call GuardaResEficRetro(fecha2, matpos(1).c_operacion, res_efi(2).val_efect, ConAdo)
      If res_efi(2).val_efect < 0.8 Or res_efi(2).val_efect > 1.25 Then
         MensajeProc = "no es eficiente la operacion " & matpos(1).c_operacion & " " & res_efi(2).val_efect
         final = True
         exito = False
      Else
         txtmsg = "El proceso finalizo correctamente"
         final = True
         exito = True
      End If
   Else
      txtmsg = "Falta alguna posicion de " & txtport
      final = True
      exito = False
   End If
End If
          
End Sub

Sub DetermSwapPrimPas(ByRef matpos() As propPosRiesgo, ByRef indswap As Integer, ByRef indppas As Integer)
If UBound(matpos, 1) = 2 Then
   If matpos(1).fValuacion = "SWAP" And matpos(2).fValuacion = "DEUDA" Then
      indswap = 1
      indppas = 2
   ElseIf matpos(2).fValuacion = "SWAP" And matpos(1).fValuacion = "DEUDA" Then
      indswap = 2
      indppas = 1
   End If
Else
   indswap = 0
   indppas = 0
End If
End Sub

Sub DetermSwapPrimAct(ByRef matpos() As propPosRiesgo, ByRef indswap As Integer, ByRef indpact As Integer)
If UBound(matpos, 1) = 2 Then
   If matpos(1).fValuacion = "SWAP" And matpos(2).fValuacion = "DEUDA" Then
      indswap = 1
      indpact = 2
   ElseIf matpos(2).fValuacion = "SWAP" And matpos(1).fValuacion = "DEUDA" Then
      indswap = 2
      indpact = 1
   End If
Else
  indswap = 0
  indpact = 0

End If
End Sub

Sub DetermSwapPrimActPas(ByRef matpos() As propPosRiesgo, ByRef indswap As Integer, ByRef indpact As Integer, ByRef indppas As Integer)
If UBound(matpos, 1) = 3 Then
   If matpos(1).fValuacion = "SWAP" And matpos(2).fValuacion = "DEUDA" And matpos(2).Tipo_Mov = 1 And matpos(3).fValuacion = "DEUDA" And matpos(3).Tipo_Mov = 4 Then
      indswap = 1
      indpact = 2
      indppas = 3
   ElseIf matpos(1).fValuacion = "SWAP" And matpos(2).fValuacion = "DEUDA" And matpos(2).Tipo_Mov = 4 And matpos(3).fValuacion = "DEUDA" And matpos(3).Tipo_Mov = 1 Then
      indswap = 1
      indpact = 3
      indppas = 2
   ElseIf matpos(2).fValuacion = "SWAP" And matpos(1).fValuacion = "DEUDA" And matpos(1).Tipo_Mov = 1 And matpos(3).fValuacion = "DEUDA" And matpos(3).Tipo_Mov = 4 Then
      indswap = 2
      indpact = 1
      indppas = 3
   ElseIf matpos(2).fValuacion = "SWAP" And matpos(1).fValuacion = "DEUDA" And matpos(1).Tipo_Mov = 4 And matpos(3).fValuacion = "DEUDA" And matpos(3).Tipo_Mov = 1 Then
      indswap = 2
      indpact = 3
      indppas = 1
   ElseIf matpos(3).fValuacion = "SWAP" And matpos(1).fValuacion = "DEUDA" And matpos(1).Tipo_Mov = 1 And matpos(2).fValuacion = "DEUDA" And matpos(2).Tipo_Mov = 4 Then
      indswap = 3
      indpact = 1
      indppas = 2
   ElseIf matpos(3).fValuacion = "SWAP" And matpos(1).fValuacion = "DEUDA" And matpos(1).Tipo_Mov = 4 And matpos(3).fValuacion = "DEUDA" And matpos(3).Tipo_Mov = 1 Then
      indswap = 3
      indpact = 2
      indppas = 1
   End If
Else
   indswap = 0
   indpact = 0
   indppas = 0
End If
End Sub

Sub DetermSwapProxySwap(ByRef matpos() As propPosRiesgo, ByRef indswap As Integer, ByRef indpswap As Integer)
If UBound(matpos, 1) > 1 Then
If matpos(1).fValuacion = "SWAP" And matpos(2).C_Posicion = 7 Then
   indswap = 1
   indpswap = 2
ElseIf matpos(2).fValuacion = "SWAP" And matpos(1).C_Posicion = 7 Then
   indswap = 2
   indpswap = 1
End If
Else
   indswap = 0
   indpswap = 0
End If
End Sub


Sub CalculoEfRetroSwapActPas(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef final As Boolean, ByRef exito As Boolean)
Dim mrvalflujo() As resValFlujo
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim parval As ParamValPos
Dim noreg As Long
Dim norelac As Long
Dim i As Long
Dim j As Long
Dim indice As Long
Dim matcoperacion() As Variant
Dim nofval As Long
Dim contreg As Long
Dim nregcalefec As Long
Dim indice0 As Long
Dim matv() As New resValIns
Dim contarf As Long
Dim fecha3 As Date
Dim indswap As Integer
Dim indppact As Integer
Dim indpppas As Integer
Dim exito3 As Boolean
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim txtmsg3 As String

mattxt = CrearFiltroPosPort(fecha2, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
'se debe de saber que tipo de operaciones y como empatan los flujos
noreg = UBound(matpos, 1)
If noreg <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   ValExacta = True
   nofval = 2
   ReDim matfecha(1 To nofval) As Date
   matfecha(1) = fecha1
   matfecha(2) = fecha2
   fecha3 = PBD1(fecha2, 1, "MX")
'se obtienen los factores adicionales para el calculo de la efectividad
   If matfecha(1) <> fechaFactR1 Then
      MatCurvasT1 = LeerCurvaCompleta(matfecha(1), exito1)
      MatFactRiesgo1 = CargaFR1Dia(matfecha(1), exito)
      fechaFactR1 = matfecha(1)
   End If
   If fecha3 <> fechaFactR2 Then
      MatCurvasT2 = LeerCurvaCompleta(fecha3, exito2)
      MatFactRiesgo2 = CargaFR1Dia(fecha3, exito)
      fechaFactR2 = fecha3
   End If
   Call DetermSwapPrimActPas(matpos, indswap, indppact, indpppas)
   If indswap <> 0 And indppact <> 0 And indpppas <> 0 Then
      ReDim res_efi(1 To nofval) As New resEficSwap
      Set parval = DeterminaPerfilVal("VALUACION")
      parval.si_int_flujos = True
      parval.sicalcdur = False
      parval.sicalcdv01 = False
      For j = 1 To nofval
          If j = 1 Then
             matv = CalcValuacion(matfecha(j), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactRiesgo1, MatCurvasT1, parval, mrvalflujo, txtmsg3, exito3)
          Else
             matv = CalcValuacion(matfecha(j), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactRiesgo2, MatCurvasT2, parval, mrvalflujo, txtmsg3, exito3)
          End If
          res_efi(j).val_efect = 0
          Call CalculosEfecSwapActPas(res_efi, matv, j, indswap, indppact, indpppas)
      Next j
      Call GuardarResEfRetroSwap2(fecha2, matpos, matposswaps, indswap, 3, res_efi(2).val_efect)
      Call GuardaResEficRetro(fecha2, matpos(1).c_operacion, res_efi(2).val_efect, ConAdo)
      If res_efi(2).val_efect < 0.8 Or res_efi(2).val_efect > 1.25 Then
         txtmsg = "no es eficiente la operacion " & matpos(1).c_operacion & " " & res_efi(2).val_efect
         final = True
         exito = False
      Else
         txtmsg = "El proceso finalizo correctamente"
         final = True
         exito = True
      End If
   Else
      txtmsg = "No se tienen los registros suficientes para hacer el calculo " & matpos(1).c_operacion
      final = True
      exito = False
   End If
Else
   txtmsg = "No hay datos para el portafolio " & txtport
   final = True
   exito = False
End If

End Sub

Sub CalculoEfRetroProxySwap(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef final As Boolean, ByRef exito As Boolean)
Dim mrvalflujo() As resValFlujo
Dim mattxt() As String
Dim exito1 As Boolean
Dim exito2 As Boolean
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim parval As ParamValPos
Dim contar As Long
Dim noreg As Long
Dim norelac As Long
Dim i As Long
Dim j As Long
Dim indice As Long
Dim matcoperacion() As Variant
Dim nofval As Long
Dim nregcalefec As Long
Dim indice0 As Long
Dim matv() As New resValIns
Dim contarf As Long
Dim indswap As Integer
Dim indpswap As Integer
Dim exito3 As Boolean
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim txtmsg3 As String
Dim fecha3 As Date

mattxt = CrearFiltroPosPort(fecha2, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
noreg = UBound(matpos, 1)
If noreg = 2 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   ValExacta = True
   nofval = 2
   ReDim matfecha(1 To nofval) As Date
   matfecha(1) = fecha1
   matfecha(2) = fecha2
   fecha3 = PBD1(fecha2, 1, "MX")
   If matfecha(1) <> fechaFactR1 Then
      MatCurvasT1 = LeerCurvaCompleta(matfecha(1), exito1)
      MatFactRiesgo1 = CargaFR1Dia(matfecha(1), exito)
      fechaFactR1 = fecha1
   End If
   If fecha3 <> fechaFactR2 Then
      MatCurvasT2 = LeerCurvaCompleta(fecha3, exito2)
      MatFactRiesgo2 = CargaFR1Dia(fecha3, exito)
      fechaFactR2 = fecha3
   End If
   Call DetermSwapProxySwap(matpos, indswap, indpswap)
   If indswap <> 0 And indpswap <> 0 Then
      ReDim res_efi(1 To nofval) As New resEfectProxySwap
      Set parval = DeterminaPerfilVal("VALUACION")
      parval.sicalcdur = False
      parval.sicalcdv01 = False
      For j = 1 To nofval
          If j = 1 Then
             matv = CalcValuacion(matfecha(j), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactRiesgo1, MatCurvasT1, parval, mrvalflujo, txtmsg3, exito3)
          Else
             matv = CalcValuacion(matfecha(j), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatFactRiesgo2, MatCurvasT2, parval, mrvalflujo, txtmsg3, exito3)
          End If
          res_efi(j).val_efect = 0
          Call CalcEfecProxySwap(res_efi, matv, j, indswap, indpswap)
      Next j
      Call GuardarResEfRetroSwap2(fecha2, matpos, matposswaps, indswap, 5, res_efi(2).val_efect)
      Call GuardaResEficRetro(fecha2, matpos(indswap).c_operacion, res_efi(2).val_efect, ConAdo)
      If res_efi(2).val_efect < 0.8 Or res_efi(2).val_efect > 1.25 Then
          txtmsg = "no es eficiente la operacion " & matpos(indswap).c_operacion & " " & res_efi(2).val_efect
          final = True
          exito = False
      Else
          txtmsg = "El proceso finalizo correctamente"
          final = True
          exito = True
      End If
    Else
    
    End If
Else
   txtmsg = "No es el no correcto de operaciones"
   final = True
   exito = False

End If
End Sub


Sub CalculosEfecSwap(ByVal tefec As Integer, res_efi() As resEficSwap, ByRef matv() As resValIns, ByVal ind As Integer, indswap As Long, indppact As Long, indpppas As Long)
Dim valora As Double
Dim valorb As Double
Dim valorc As Double
Dim valord As Double
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim valor4 As Double
Dim valefi As Double
Dim valefi1 As Double
Dim valefi2 As Double

If tefec = 1 Then                                                                      'tipo de calculo de efectividad
   res_efi(ind).v_swapact = matv(indswap).ps_activa                                     'swap activa
   res_efi(ind).v_swappas = matv(indswap).ps_pasiva                                     'swap pasiva
   res_efi(ind).v_primact = matv(indswap).ps_pasiva                                     'primaria activa ficticia
   res_efi(ind).v_primpas = matv(indpppas).ps_pasiva                                    'primaria pasiva
   res_efi(ind).id_swapact = matv(indswap).ps_activa - matv(indswap).pl_activa
   res_efi(ind).id_swappas = matv(indswap).ps_pasiva - matv(indswap).pl_pasiva
   res_efi(ind).id_primact = matv(indswap).ps_pasiva - matv(indswap).pl_pasiva
   res_efi(ind).id_primpas = matv(indpppas).ps_pasiva - matv(indpppas).pl_pasiva
ElseIf tefec = 2 Then                   'credito
   res_efi(ind).v_swappas = matv(indswap).ps_pasiva                             'swap pasiva
   res_efi(ind).v_primact = matv(indppact).ps_activa                            'primaria activa
   res_efi(ind).id_swappas = matv(indswap).ps_pasiva - matv(indswap).pl_pasiva
   res_efi(ind).id_primact = matv(indppact).ps_activa - matv(indppact).pl_activa
ElseIf tefec = 3 Then
   res_efi(ind).v_swapact = matv(indswap).ps_activa                            'swap activa
   res_efi(ind).v_swappas = matv(indswap).ps_pasiva                            'swap pasiva
   res_efi(ind).v_primact = matv(indppact).ps_activa                              'primaria activa
   res_efi(ind).v_primpas = matv(indpppas).ps_pasiva                              'primaria pasiva
   res_efi(ind).id_swapact = matv(indswap).ps_activa - matv(indswap).pl_activa         'int dev pactiva
   res_efi(ind).id_swappas = matv(indswap).ps_pasiva - matv(indswap).pl_pasiva         'int dev pasiva
   res_efi(ind).id_primact = matv(indppact).ps_activa - matv(indppact).pl_activa
   res_efi(ind).id_primpas = matv(indpppas).ps_pasiva - matv(indpppas).pl_pasiva
End If

If ind <> 1 Then
   If tefec = 1 Then                   'si es una eficiencia de tipo 1 entonces
 'cambios en la primaria activa
      res_efi(ind).var_pp_act = res_efi(ind).v_primact - res_efi(ind).id_primact - (res_efi(ind - 1).v_primact - res_efi(ind - 1).id_primact)
 'cambios en la primaria pasiva
      res_efi(ind).var_pp_pas = res_efi(ind).v_primpas - res_efi(ind).id_primpas - (res_efi(ind - 1).v_primpas - res_efi(ind - 1).id_primpas)
 'cambios en la pos activa
      res_efi(ind).var_swap_act = res_efi(ind).v_swapact - res_efi(ind).id_swapact - (res_efi(ind - 1).v_swapact - res_efi(ind - 1).id_swapact)
 'cambios en la pos pasiva
      res_efi(ind).var_swap_pas = res_efi(ind).v_swappas - res_efi(ind).id_swappas - (res_efi(ind - 1).v_swappas - res_efi(ind - 1).id_swappas)
      valor1 = res_efi(ind).var_pp_act + res_efi(ind).var_pp_pas 'cambios prim activa(pasiva swap)+cambios primaria pasiva
      valor2 = res_efi(ind).var_swap_act + res_efi(ind).var_swap_pas 'cambios pata activa+cambios pasiva
      If valor1 <> 0 Then  'swap/primaria
         valefi = valor2 / valor1                         'tipo de eficiencia 1
      Else
         valefi = 0
      End If
      res_efi(ind).val_efect = valefi
    '  If res_efi(ind).val_efect > 1.25 Or res_efi(ind).val_efect < 0.8 Then MsgBox "alerta"
   End If
   
   If tefec = 2 Then
      res_efi(ind).var_pp_act = res_efi(ind).v_primact - res_efi(ind).id_primact - (res_efi(ind - 1).v_primact - res_efi(ind - 1).id_primact)
      res_efi(ind).var_swap_pas = res_efi(ind).v_swappas - res_efi(ind).id_swappas - (res_efi(ind - 1).v_swappas - res_efi(ind - 1).id_swappas)
      If res_efi(ind).var_pp_act <> 0 Then
         valefi = res_efi(ind).var_swap_pas / res_efi(ind).var_pp_act                           'tipo de eficiencia 2
      Else
         valefi = 0
      End If
     res_efi(ind).val_efect = valefi
    ' If res_efi(ind).val_efect > 1.25 Or res_efi(ind).val_efect < 0.8 Then MsgBox "alerta "
End If
If tefec = 3 Then
   res_efi(ind).var_pp_act = res_efi(ind).v_primact - res_efi(ind).id_primact - (res_efi(ind - 1).v_primact - res_efi(ind - 1).id_primact)        'PRIM ACT
'CAMBIOS EN LA PATA ACTIVA
   res_efi(ind).var_swap_act = res_efi(ind).v_swapact - res_efi(ind).id_swapact - (res_efi(ind - 1).v_swapact - res_efi(ind - 1).id_swapact)      'swap ACTIVA
'CAMBIOS EN LA PATA PASIVA
   res_efi(ind).var_swap_pas = res_efi(ind).v_swappas - res_efi(ind).id_swappas - (res_efi(ind - 1).v_swappas - res_efi(ind - 1).id_swappas)      'swap PASIVA
'CAMBIOS EN LA PRIMARIA PASIVA
   res_efi(ind).var_pp_pas = res_efi(ind).v_primpas - res_efi(ind).id_primpas - (res_efi(ind - 1).v_primpas - res_efi(ind - 1).id_primpas)        'PRIM PAS
'tipo calculo 3
   If res_efi(ind).var_pp_act <> 0 Then
      valefi1 = res_efi(ind).var_swap_pas / res_efi(ind).var_pp_act   '(dif pasiva/dif prim act) eficiencia primaria pasiva
   Else
     valefi1 = 0
   End If
   If res_efi(ind).var_pp_pas <> 0 Then
      valefi2 = res_efi(ind).var_swap_act / res_efi(ind).var_pp_pas   'eficiencia primaria activa
   Else
      valefi2 = 0
   End If
   If valefi1 < 0.8 Or valefi1 > 1.25 Then
      valefi1 = 0
   End If
   If valefi2 < 0.8 Or valefi2 > 1.25 Then
     valefi2 = 0
   End If
   res_efi(ind).val_efect = (valefi1 + valefi2) / 2
End If
End If
End Sub

Sub CalculosEfecSwapAct(res_efi() As resEficSwap, ByRef matv() As resValIns, ByVal ind As Integer, indswap As Integer, indppact As Integer)
Dim valora As Double
Dim valorb As Double
Dim valorc As Double
Dim valord As Double
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim valor4 As Double
Dim valefi As Double
Dim valefi1 As Double
Dim valefi2 As Double

res_efi(ind).v_swappas = matv(indswap).ps_pasiva                             'swap pasiva
res_efi(ind).v_primact = matv(indppact).ps_activa                            'primaria activa
res_efi(ind).id_swappas = matv(indswap).ps_pasiva - matv(indswap).pl_pasiva
res_efi(ind).id_primact = matv(indppact).ps_activa - matv(indppact).pl_activa

If ind <> 1 Then
   res_efi(ind).var_pp_act = res_efi(ind).v_primact - res_efi(ind).id_primact - (res_efi(ind - 1).v_primact - res_efi(ind - 1).id_primact)
   res_efi(ind).var_swap_pas = res_efi(ind).v_swappas - res_efi(ind).id_swappas - (res_efi(ind - 1).v_swappas - res_efi(ind - 1).id_swappas)
   If res_efi(ind).var_pp_act <> 0 Then
      valefi = res_efi(ind).var_swap_pas / res_efi(ind).var_pp_act                           'tipo de eficiencia 2
   Else
      valefi = 0
   End If
   res_efi(ind).val_efect = valefi
End If
End Sub

Sub CalculosEfecSwapPas(res_efi() As resEficSwap, ByRef matv() As resValIns, ByVal ind As Integer, indswap As Integer, indpppas As Integer)
Dim valora As Double
Dim valorb As Double
Dim valorc As Double
Dim valord As Double
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim valor4 As Double
Dim valefi As Double
Dim valefi1 As Double
Dim valefi2 As Double

res_efi(ind).v_swapact = matv(indswap).ps_activa                                     'swap activa
res_efi(ind).v_swappas = matv(indswap).ps_pasiva                                     'swap pasiva
res_efi(ind).v_primact = matv(indswap).ps_pasiva                                     'primaria activa ficticia
res_efi(ind).v_primpas = matv(indpppas).ps_pasiva                                    'primaria pasiva
res_efi(ind).id_swapact = matv(indswap).ps_activa - matv(indswap).pl_activa
res_efi(ind).id_swappas = matv(indswap).ps_pasiva - matv(indswap).pl_pasiva
res_efi(ind).id_primact = matv(indswap).ps_pasiva - matv(indswap).pl_pasiva
res_efi(ind).id_primpas = matv(indpppas).ps_pasiva - matv(indpppas).pl_pasiva

If ind <> 1 Then
 'cambios en la primaria activa
   res_efi(ind).var_pp_act = res_efi(ind).v_primact - res_efi(ind).id_primact - (res_efi(ind - 1).v_primact - res_efi(ind - 1).id_primact)
'cambios en la primaria pasiva
   res_efi(ind).var_pp_pas = res_efi(ind).v_primpas - res_efi(ind).id_primpas - (res_efi(ind - 1).v_primpas - res_efi(ind - 1).id_primpas)
 'cambios en la pos activa
   res_efi(ind).var_swap_act = res_efi(ind).v_swapact - res_efi(ind).id_swapact - (res_efi(ind - 1).v_swapact - res_efi(ind - 1).id_swapact)
 'cambios en la pos pasiva
   res_efi(ind).var_swap_pas = res_efi(ind).v_swappas - res_efi(ind).id_swappas - (res_efi(ind - 1).v_swappas - res_efi(ind - 1).id_swappas)
   valor1 = res_efi(ind).var_pp_act - res_efi(ind).var_pp_pas         'cambios prim activa(pasiva swap)+cambios primaria pasiva
   valor2 = res_efi(ind).var_swap_act - res_efi(ind).var_swap_pas    'cambios pata activa+cambios pasiva
   If valor1 <> 0 Then  'swap/primaria
      valefi = -valor2 / valor1                         'tipo de eficiencia 1
   Else
      valefi = 0
   End If
   res_efi(ind).val_efect = valefi
End If
End Sub

Sub CalculosEfecSwapActPas(res_efi() As resEficSwap, ByRef matv() As resValIns, ByVal ind As Integer, indswap As Integer, indppact As Integer, indpppas As Integer)
Dim valora As Double
Dim valorb As Double
Dim valorc As Double
Dim valord As Double
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim valor4 As Double
Dim valefi As Double
Dim valefi1 As Double
Dim valefi2 As Double

res_efi(ind).v_swapact = matv(indswap).ps_activa                               'swap activa
res_efi(ind).v_swappas = matv(indswap).ps_pasiva                               'swap pasiva
res_efi(ind).v_primact = matv(indppact).ps_activa                              'primaria activa
res_efi(ind).v_primpas = matv(indpppas).ps_pasiva                              'primaria pasiva
res_efi(ind).id_swapact = matv(indswap).ps_activa - matv(indswap).pl_activa         'int dev pactiva
res_efi(ind).id_swappas = matv(indswap).ps_pasiva - matv(indswap).pl_pasiva         'int dev pasiva
res_efi(ind).id_primact = matv(indppact).ps_activa - matv(indppact).pl_activa
res_efi(ind).id_primpas = matv(indpppas).ps_pasiva - matv(indpppas).pl_pasiva

If ind <> 1 Then
'cambios en la primaria activa
   res_efi(ind).var_pp_act = res_efi(ind).v_primact - res_efi(ind).id_primact - (res_efi(ind - 1).v_primact - res_efi(ind - 1).id_primact)        'PRIM ACT
'CAMBIOS EN LA PATA ACTIVA
   res_efi(ind).var_swap_act = res_efi(ind).v_swapact - res_efi(ind).id_swapact - (res_efi(ind - 1).v_swapact - res_efi(ind - 1).id_swapact)      'swap ACTIVA
'CAMBIOS EN LA PATA PASIVA
   res_efi(ind).var_swap_pas = res_efi(ind).v_swappas - res_efi(ind).id_swappas - (res_efi(ind - 1).v_swappas - res_efi(ind - 1).id_swappas)      'swap PASIVA
'CAMBIOS EN LA PRIMARIA PASIVA
   res_efi(ind).var_pp_pas = res_efi(ind).v_primpas - res_efi(ind).id_primpas - (res_efi(ind - 1).v_primpas - res_efi(ind - 1).id_primpas)        'PRIM PAS
'tipo calculo 3
   If res_efi(ind).var_pp_act <> 0 Then
      valefi1 = res_efi(ind).var_swap_pas / res_efi(ind).var_pp_act   '(dif pasiva/dif prim act)eficiencia primaria pasiva
   Else
     valefi1 = 0
   End If
   If res_efi(ind).var_pp_pas <> 0 Then
      valefi2 = res_efi(ind).var_swap_act / res_efi(ind).var_pp_pas   'eficiencia primaria activa
   Else
      valefi2 = 0
   End If
   If valefi1 < 0.8 Or valefi1 > 1.25 Then
      valefi1 = 0
   End If
   If valefi2 < 0.8 Or valefi2 > 1.25 Then
     valefi2 = 0
   End If
   res_efi(ind).val_efect = (valefi1 + valefi2) / 2
End If
End Sub

Sub CalcEfecProxySwap(res_efi() As resEfectProxySwap, ByRef matv() As resValIns, ByVal ind As Integer, indswap As Integer, indpswap As Integer)

res_efi(ind).v_swap_act = matv(indswap).ps_activa                                      'swap activa
res_efi(ind).v_swap_pas = matv(indswap).ps_pasiva                                      'swap pasiva
res_efi(ind).v_pswap_act = matv(indpswap).ps_activa                                    'proxy swap activa
res_efi(ind).v_pswap_pas = matv(indpswap).ps_pasiva                                    'proxy swap activa
res_efi(ind).id_swap_act = matv(indswap).ps_activa - matv(indswap).pl_activa
res_efi(ind).id_swap_pas = matv(indswap).ps_pasiva - matv(indswap).pl_pasiva
res_efi(ind).id_pswap_act = matv(indpswap).ps_activa - matv(indpswap).pl_activa
res_efi(ind).id_pswap_pas = matv(indpswap).ps_pasiva - matv(indpswap).pl_pasiva

If ind <> 1 Then
   res_efi(ind).var_swap_limpio = res_efi(ind).v_swap_act - res_efi(ind).v_swap_pas - (res_efi(ind - 1).v_swap_act - res_efi(ind - 1).v_swap_pas) - (res_efi(ind).id_swap_act - res_efi(ind).id_swap_pas) + (res_efi(ind - 1).id_swap_act - res_efi(ind - 1).id_swap_pas)
   res_efi(ind).var_pswap_limpio = res_efi(ind).v_pswap_act - res_efi(ind).v_pswap_pas - (res_efi(ind - 1).v_pswap_act - res_efi(ind - 1).v_pswap_pas) - (res_efi(ind).id_pswap_act - res_efi(ind).id_pswap_pas) + (res_efi(ind - 1).id_pswap_act - res_efi(ind - 1).id_pswap_pas)
   If res_efi(ind).var_pswap_limpio <> 0 Then  'swap/primaria
      res_efi(ind).val_efect = -res_efi(ind).var_swap_limpio / res_efi(ind).var_pswap_limpio                        'tipo de eficiencia 1
   Else
      res_efi(ind).val_efect = 0
   End If
End If

End Sub


Sub VerificarSwapIneficiente(ByVal fecha As Date, ByRef matpos() As propPosRiesgo, ByRef matefic() As Variant)
Dim exito As Boolean
Dim matposriesgo3() As Variant
Dim mateficpros() As Variant
Dim efecpros As Double
Dim contar As Long
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim siefpros As Integer

ReDim MatEmision2(1 To 3, 1 To 1) As Variant
contar = 0
noreg = UBound(matefic, 1)
For i = 1 To noreg
 If Val(matefic(i, 27)) < 80 Or Val(matefic(i, 27)) > 125 Then
   MensajeProc = "El swap " & matefic(i, 1) & " es ineficiente"
   contar = contar + 1
   ReDim Preserve MatEmision2(1 To 3, 1 To contar) As Variant
   MatEmision2(1, contar) = matefic(i, 1)
   MatEmision2(2, contar) = matefic(i, 2)
   MatEmision2(3, contar) = matefic(i, 3)
 End If
Next i
If contar <> 0 Then
   siefpros = MsgBox("Hay swaps ineficientes. Desea correr la eficiencia prospectiva", vbYesNo)
   If siefpros = 7 Then
    Exit Sub
   End If
   noreg = UBound(matpos, 1)

 For i = 1 To contar
  ReDim matc(1 To 3, 1 To 1) As Variant
  matc(1, 1) = MatEmision2(1, i)
  matc(2, 1) = MatEmision2(2, i)
  matc(3, 1) = MatEmision2(3, i)
  matposriesgo3 = FiltrarPosG1(matpos, matc)
  Call CEficProsSwapsPort(fecha, "", "", exito)
 'una vez calculada la eficiencia prospectiva se anexa a los resultados de la eficiencia retro
 For j = 1 To noreg
 If matefic(j, 1) = MREFProsSwap(i, 1) Then
 matefic(j, 20) = Format(MREFProsSwap(i, 2) * 100, "####.00")
 End If
 Next j
 Next i
End If
End Sub

Function DefFechasEfCob(ByVal fecha As Date, ByVal f_val As Date) As Date()
On Error GoTo hayerror
'se obtendran las fechas para el calculo de la eficiencia de la coberura en funcion de los flujos del swap a analizar
Dim contar As Integer
Dim fechax As Date

contar = 1
'se define el calendario de fechas
fechax = CDate("01/" & Month(fecha) & "/" & Year(fecha))
ReDim matfecha2(1 To 1) As Date
matfecha2(contar) = fecha
Do While fechax - 1 <= f_val
   If fechax - 1 > fecha And fechax - 1 < f_val Then
      contar = contar + 1
      ReDim Preserve matfecha2(1 To contar) As Date
      matfecha2(contar) = fechax - 1
   End If
   fechax = DateAdd("m", 1, fechax)
Loop
If contar = 1 Then
ReDim Preserve matfecha2(1 To 2) As Date
   matfecha2(1) = fecha
   matfecha2(2) = f_val
End If
DefFechasEfCob = matfecha2
On Error GoTo 0
Exit Function
hayerror:
  MsgBox "DetFechasEfCob" & error(Err())
End Function

Function DetTEfec(ByRef matpos() As propPosRiesgo, ByVal indppact As Long, ByVal indpppas As Long)
On Error GoTo hayerror
If UBound(matpos, 1) = 2 Then
   If indppact <> 0 Then
      DetTEfec = 2
   ElseIf indpppas <> 0 Then
      DetTEfec = 1
   End If
ElseIf UBound(matpos, 1) = 3 Then
   DetTEfec = 3
Else
   DetTEfec = 0
End If
Exit Function
hayerror:
 MsgBox "DetTEfec " & error(Err())
End Function

Sub DetermIndEfecCob(ByRef matpos() As propPosRiesgo, ByRef indice1 As Long, ByRef indice2 As Long, ByRef indice3 As Long)
On Error GoTo hayerror
Dim noreg As Long
Dim i As Long
noreg = UBound(matpos, 1)
For i = 1 To noreg
    If matpos(i).fValuacion = "SWAP" Then
       indice1 = i
    End If
Next i
For i = 1 To noreg
    If matpos(i).fValuacion = "DEUDA" And matpos(i).Signo_Op = 1 Then
       indice2 = i
    End If
Next i
For i = 1 To noreg
    If matpos(i).fValuacion = "DEUDA" And matpos(i).Signo_Op = -1 Then
       indice3 = i
    End If
Next i
Exit Sub
hayerror:
MsgBox "DetermIndEfecCob " & error(Err())
End Sub

Sub CEficProsSwapActivaPasiva(ByVal fecha As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef exito As Boolean)
If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim exito1 As Boolean
Dim matresefic() As Variant
Dim mateficpros() As Variant
Dim parval1(1 To 2) As Variant
Dim parval2(1 To 4) As Variant
Dim matpact() As Variant
Dim matres() As Variant
Dim noreg As Integer
Dim r As Integer
Dim i As Integer
Dim coperacion As String
Dim fcurva As Date
Dim fvenc As Date
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim matfactestres() As Double
Dim nofval As Integer
Dim noescen As Integer
Dim jj As Integer
Dim j As Integer
Dim valefi As Double
Dim valefi1 As Double
Dim valefi2 As Double
Dim contar3 As Integer
Dim efecpros As Double
Dim txtnomarch As String
Dim mata() As Variant
Dim hinicio As Date
Dim finicio As Date
Dim fechax As Date
Dim matfr0() As Double
Dim matfr() As Double
Dim matfechas() As Date
Dim indswap As Integer
Dim indppact As Integer
Dim indpppas As Integer
Dim parval As New ParamValPos
Dim res_efi() As New resEficSwap
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim txtmsg2 As String
Dim txtmsg3 As String
Dim txtmsg0 As String
Dim mrvalflujo() As New resValFlujo
Dim matv() As New resValIns
Dim suma As Double

finicio = Date
hinicio = Time
matfechas = DetFechasEscEfic
noescen = UBound(matfechas, 1)
mattxt = CrearFiltroPosPort(fecha, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
'se debe de saber que tipo de operaciones y como empatan los flujos
noreg = UBound(matpos, 1)
If noreg <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   If exito2 Then
      Call DetermSwapPrimActPas(matpos, indswap, indppact, indpppas)
      Dim MatFechasEf() As Date
      fvenc = matposswaps(1).FvencSwap
      fcurva = PBD1(fecha, 1, "MX")
      coperacion = matpos(1).c_operacion
      MatFechasEf = DefFechasEfCob(fecha, fvenc) 'se toma la fecha del swap
      nofval = UBound(MatFechasEf, 1)
      matfr0 = CargaFR1Dia(fcurva, exito)
      ReDim matfactestres(1 To 10, 1 To NoFactores) As Double
      ReDim mateficpros(1 To nofval, 1 To noescen + 1) As Variant
      ReDim matrest(1 To 1) As Variant
      Set parval = DeterminaPerfilVal("EFECTIVIDAD")
      ReDim MDblTasas1(1 To NoFactores, 1 To 1) As Double
      ReDim matresef(1 To nofval - 1, 1 To 4) As Variant
      ReDim res_efi(1 To nofval)
      ValExacta = False
      For i = 1 To nofval - 1
          matresef(i, 1) = MatFechasEf(i)
          matresef(i, 2) = MatFechasEf(i + 1)
          matresef(i, 3) = noescen
      Next i
      For r = 1 To noescen    'el no de escenarios de estres
          Call GenEscenariosProspectivos(matfechas(r), MatFechasEf, matfr0, matfr)
          For i = 1 To nofval
              For j = 1 To NoFactores
                  MDblTasas1(j, 1) = matfr(i, j)
              Next j
              res_efi(i).val_efect = 0
              parval.perfwd = MatFechasEf(i) - fecha
              matv = CalcValuacion(MatFechasEf(i), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MDblTasas1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
              Call CalculosEfecSwapActPas(res_efi, matv, i, indswap, indppact, indpppas)
              If i <> 1 Then
                 If res_efi(i).val_efect >= 0.8 And res_efi(i).val_efect <= 1.25 Then
                    matresef(i - 1, 4) = matresef(i - 1, 4) + 1
                 Else
                    matresef(i - 1, 4) = matresef(i - 1, 4) + 0
                 End If
              End If
              DoEvents
           Next i
      Next r
      suma = 0
      For i = 1 To nofval - 1
          suma = suma + matresef(i, 4)
      Next i
      efecpros = suma / ((nofval - 1) * noescen)
      Call IniciarConexOracle(conAdo2, BDIKOS)
      Call GuardaResEfiPros(fecha, coperacion, efecpros, conAdo2)
      conAdo2.Close
      Call GuardarResEfectPros(fecha, coperacion, matresef)
      Call ValidarOperacion3(coperacion, matpos(indswap).HoraRegOp, finicio, hinicio, Date, Time)
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      exito = False
      txtmsg = txtmsg2
   End If
Else
  MsgBox "No hay registros en la posicion simulada"
  exito = False
End If
Exit Sub
hayerror:
    MsgBox "CEficProsSwapsPort " & error(Err())
  exito = False
End Sub


Sub CEficProsSwapActiva(ByVal fecha As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef exito As Boolean)
If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim exito1 As Boolean
Dim matresefic() As Variant
Dim mateficpros() As Variant
Dim parval1(1 To 2) As Variant
Dim parval2(1 To 4) As Variant
Dim matres() As Variant
Dim noreg As Integer
Dim r As Integer
Dim i As Integer
Dim coperacion As String
Dim fcurva As Date
Dim fvenc As Date
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim matfactestres() As Double
Dim nofval As Integer
Dim noescen As Integer
Dim suma As Double
Dim jj As Integer
Dim j As Integer
Dim parval As New ParamValPos
Dim valefi As Double
Dim efecpros As Double
Dim mata() As Variant
Dim fechax As Date
Dim htiempo As Integer
Dim matfr0() As Double
Dim matfr() As Double
Dim matx() As Variant
Dim matx1() As Double
Dim matfechas() As Date
Dim matinctasa() As Double
Dim indswap As Integer
Dim indppact As Integer
Dim res_efi() As New resEficSwap
Dim exito2 As Boolean
Dim txtmsg2 As String
Dim txtmsg0 As String
Dim txtmsg3 As String
Dim exito3 As Boolean
Dim mrvalflujo() As resValFlujo
Dim matv() As New resValIns
Dim hinicio As Date

finicio = Date
hinicio = Time
matfechas = DetFechasEscEfic
noescen = UBound(matfechas, 1)
mattxt = CrearFiltroPosPort(fecha, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
'se debe de saber que tipo de operaciones y como empatan los flujos
noreg = UBound(matpos, 1)
If noreg <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   If exito2 Then
      Call DetermSwapPrimAct(matpos, indswap, indppact)
      Dim MatFechasEf() As Date
      fvenc = matposswaps(indswap).FvencSwap
      fcurva = PBD1(fecha, 1, "MX")
      coperacion = matpos(indswap).c_operacion
      MatFechasEf = DefFechasEfCob(fecha, fvenc) 'se toma la fecha del swap
      nofval = UBound(MatFechasEf, 1)
      matfr0 = CargaFR1Dia(fcurva, exito)
      ValExacta = False
      Set parval = DeterminaPerfilVal("EFECTIVIDAD")
      ReDim MDblTasas1(1 To NoFactores, 1 To 1) As Double
      ReDim matresef(1 To nofval - 1, 1 To 4) As Variant
      ReDim res_efi(1 To nofval)
      For i = 1 To nofval - 1
          matresef(i, 1) = MatFechasEf(i)
          matresef(i, 2) = MatFechasEf(i + 1)
          matresef(i, 3) = noescen
      Next i
      For r = 1 To noescen    'el no de escenarios de estres
          Call GenEscenariosProspectivos(matfechas(r), MatFechasEf, matfr0, matfr)
          For i = 1 To nofval
              For j = 1 To NoFactores
                  MDblTasas1(j, 1) = matfr(r, j)
              Next j
'realiza la valuacion de los flujos
              res_efi(i).val_efect = 0
              parval.perfwd = MatFechasEf(i) - fecha
              matv = CalcValuacion(MatFechasEf(i), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MDblTasas1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
'en matv estan todas las valuaciones
              Call CalculosEfecSwapAct(res_efi, matv, i, indswap, indppact)
              If i <> 1 Then
                 If res_efi(i).val_efect >= 0.8 And res_efi(i).val_efect <= 1.25 Then
                    matresef(i - 1, 4) = matresef(i - 1, 4) + 1
                 Else
                    matresef(i - 1, 4) = matresef(i - 1, 4) + 0
                 End If
              End If
              DoEvents
          Next i
      Next r
      suma = 0
      For i = 1 To nofval - 1
          suma = suma + matresef(i, 4)
      Next i
      efecpros = suma / ((nofval - 1) * noescen)
      Call IniciarConexOracle(conAdo2, BDIKOS)
      Call GuardaResEfiPros(fecha, coperacion, efecpros, conAdo2)
      conAdo2.Close
      Call GuardarResEfectPros(fecha, coperacion, matresef)
      Call ValidarOperacion3(coperacion, matpos(indswap).HoraRegOp, finicio, hinicio, Date, Time)
      txtmsg = "El proceso finalizo correctamente"
      exito = True
Else
  MsgBox "No hay registros en la posicion simulada"
  exito = False
End If
End If
Exit Sub
hayerror:
    MsgBox "CEficProsSwapsPort " & error(Err())
  exito = False
End Sub

Sub GuardarResEfectPros(ByVal fecha As Date, ByVal coperacion As String, ByRef matres() As Variant)
Dim txtfecha As String
Dim txtborra As String
Dim txtcadena As String
Dim i As Integer
Dim txtfecha1 As String
Dim txtfecha2 As String
      
      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtborra = "DELETE FROM " & TablaResEfectPros & " WHERE F_CALCULO = " & txtfecha
      txtborra = txtborra & " AND COPERACION = '" & coperacion & "'"
      ConAdo.Execute txtborra
      For i = 1 To UBound(matres, 1)
        txtfecha1 = "to_date('" & Format(matres(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
        txtfecha2 = "to_date('" & Format(matres(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
        txtcadena = "INSERT INTO " & TablaResEfectPros & " VALUES("
        txtcadena = txtcadena & txtfecha & ","
        txtcadena = txtcadena & "'" & coperacion & "',"
        txtcadena = txtcadena & txtfecha1 & ","
        txtcadena = txtcadena & txtfecha2 & ","
        txtcadena = txtcadena & matres(i, 3) & ","
        txtcadena = txtcadena & matres(i, 4) & ")"
        ConAdo.Execute txtcadena
      Next i

End Sub

Sub GenEscenariosProspectivos(ByVal fechaesc As Date, ByRef matfechasf() As Date, ByRef matfr0() As Double, ByRef matfriesgo() As Double)
Dim i As Integer
Dim j As Integer
Dim matfres() As Double
Dim noreg As Integer
noreg = UBound(matfechasf, 1)
    matfres = EstresFREsc(fechaesc, matfr0)
    ReDim matfriesgo(1 To noreg, 1 To NoFactores) As Double
    For i = 1 To UBound(matfechasf, 1)
        For j = 1 To NoFactores
            matfriesgo(i, j) = matfres(j, 1)
        Next j
      Next i
End Sub

Sub CEficProsSwapPasiva(ByVal fecha As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef exito As Boolean)
If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim exito1 As Boolean
Dim matresefic() As Variant
Dim mateficpros() As Variant
Dim parval1(1 To 2) As Variant
Dim parval2(1 To 4) As Variant
Dim matpact() As Variant
Dim matppas() As Variant
Dim matprimact() As Variant
Dim matprimpas() As Variant
Dim matidact() As Variant
Dim matidpas() As Variant
Dim matidprimact() As Variant
Dim matidprimpas() As Variant
Dim matres() As Variant
Dim noreg As Integer
Dim r As Integer
Dim i As Integer
Dim coperacion As String
Dim fcurva As Date
Dim fvenc As Date
Dim teficiencia As Integer
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim matfactestres() As Double
Dim nofval As Integer
Dim noescen As Integer
Dim suma As Double
Dim jj As Integer
Dim j As Integer
Dim valefi As Double
Dim valefi1 As Double
Dim valefi2 As Double
Dim contar3 As Integer
Dim efecpros As Double
Dim txtnomarch As String
Dim parval As New ParamValPos
Dim mata() As Variant
Dim hinicio As Date
Dim finicio As Date
Dim fechax As Date
Dim htiempo As Integer
Dim matfr0() As Double
Dim matfr() As Double
Dim matfechas() As Date
Dim indswap As Integer
Dim indpppas As Integer
Dim mrvalflujo() As New resValFlujo
Dim res_efi() As New resEficSwap
Dim matv() As New resValIns
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim txtmsg2 As String
Dim txtmsg3 As String
Dim txtmsg0 As String

finicio = Date
hinicio = Time
matfechas = DetFechasEscEfic
noescen = UBound(matfechas, 1)
mattxt = CrearFiltroPosPort(fecha, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
'se debe de saber que tipo de operaciones y como empatan los flujos
noreg = UBound(matpos, 1)
If noreg <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   If exito2 Then
      Call DetermSwapPrimPas(matpos, indswap, indpppas)
      fcurva = PBD1(fecha, 1, "MX")
      Dim MatFechasEf() As Date
      fvenc = matposswaps(indswap).FvencSwap
      coperacion = matpos(indswap).c_operacion
      MatFechasEf = DefFechasEfCob(fecha, fvenc) 'se toma la fecha del swap
      Set parval = DeterminaPerfilVal("EFECTIVIDAD")

      matfr0 = CargaFR1Dia(fcurva, exito)
      nofval = UBound(MatFechasEf, 1)
      ReDim MDblTasas1(1 To NoFactores, 1 To 1) As Double
      ReDim matresef(1 To nofval - 1, 1 To 4) As Variant
      ReDim res_efi(1 To nofval)
      For i = 1 To nofval - 1
          matresef(i, 1) = MatFechasEf(i)
          matresef(i, 2) = MatFechasEf(i + 1)
          matresef(i, 3) = noescen
      Next i
      For r = 1 To noescen    'el no de escenarios de estres
          Call GenEscenariosProspectivos(matfechas(r), MatFechasEf, matfr0, matfr)
          For i = 1 To nofval
              For j = 1 To NoFactores
                  MDblTasas1(j, 1) = matfr(i, j)
              Next j
              res_efi(i).val_efect = 0
              parval.perfwd = MatFechasEf(i) - fecha
              matv = CalcValuacion(MatFechasEf(i), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MDblTasas1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
              Call CalculosEfecSwapPas(res_efi, matv, i, indswap, indpppas)
              If i <> 1 Then
                 If res_efi(i).val_efect >= 0.8 And res_efi(i).val_efect <= 1.25 Then
                    matresef(i - 1, 4) = matresef(i - 1, 4) + 1
                 Else
                    matresef(i - 1, 4) = matresef(i - 1, 4) + 0
                 End If
              End If
              DoEvents
          Next i
      Next r
      suma = 0
      For i = 1 To nofval - 1
          suma = suma + matresef(i, 4)
      Next i
      efecpros = suma / ((nofval - 1) * noescen)
      MensajeProc = "La efectividad de la operacion " & coperacion & " fue del " & Format(efecpros, "##0.00 %")
      Call IniciarConexOracle(conAdo2, BDIKOS)
      Call GuardaResEfiPros(fecha, coperacion, efecpros, conAdo2)
      conAdo2.Close
      Call GuardarResEfectPros(fecha, matpos(indswap).c_operacion, matresef)
      Call ValidarOperacion3(matpos(indswap).c_operacion, matpos(indswap).HoraRegOp, finicio, hinicio, Date, Time)
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      exito = False
      txtmsg = txtmsg2
   End If
Else
  MsgBox "No hay registros en la posicion simulada"
  exito = False
End If
Exit Sub
hayerror:
    MsgBox "CEficProsSwapsPort " & error(Err())
  exito = False
End Sub

Sub CEficProsProxySwap(ByVal fecha As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef exito As Boolean)
If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim exito1 As Boolean
Dim matresefic() As Variant
Dim mateficpros() As Variant
Dim matres() As Variant
Dim noreg As Integer
Dim r As Integer
Dim i As Integer
Dim coperacion As String
Dim fcurva As Date
Dim fvenc As Date
Dim MDblTasas1() As Double
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim parval As New ParamValPos
Dim nofval As Integer
Dim noescen As Integer
Dim contar0 As Integer
Dim j As Integer
Dim valefi As Double
Dim efecpros As Double
Dim fechax As Date
Dim htiempo As Integer
Dim matfr0() As Double
Dim matfr() As Double
Dim indswap As Integer
Dim indpswap As Integer
Dim matv() As New resValIns
Dim mrvalflujo() As New resValFlujo
Dim res_efi() As New resEfectProxySwap
Dim MatFechasEf() As Date
Dim matfechas() As Date
Dim exito2 As Boolean
Dim exito3 As Boolean
Dim txtmsg0 As String
Dim txtmsg2 As String
Dim txtmsg3 As String
Dim suma As Double
Dim finicio As Date
Dim hinicio As Date

finicio = Date
hinicio = Time
matfechas = DetFechasEscEfic
noescen = UBound(matfechas, 1)
mattxt = CrearFiltroPosPort(fecha, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
'se debe de saber que tipo de operaciones y como empatan los flujos
noreg = UBound(matpos, 1)
If noreg <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   If exito2 Then
      SiIncTasaCVig = False
      Call DetermSwapProxySwap(matpos, indswap, indpswap)
      If indswap <> 0 And indpswap <> 0 Then
      fvenc = matposswaps(1).FvencSwap
      fcurva = PBD1(fecha, 1, "MX")
      coperacion = matpos(indswap).c_operacion
      MatFechasEf = DefFechasEfCob(fecha, fvenc) 'se toma la fecha del swap
      Set parval = DeterminaPerfilVal("EFECTIVIDAD")
      ValExacta = False
      matfr0 = CargaFR1Dia(fcurva, exito)
      nofval = UBound(MatFechasEf, 1)
      ReDim MDblTasas1(1 To NoFactores, 1 To 1) As Double
      ReDim matresef(1 To nofval - 1, 1 To 4) As Variant
      ReDim res_efi(1 To nofval)
      For i = 1 To nofval - 1
          matresef(i, 1) = MatFechasEf(i)
          matresef(i, 2) = MatFechasEf(i + 1)
          matresef(i, 3) = noescen
      Next i
      For r = 1 To noescen    'el no de escenarios de estres
          Call GenEscenariosProspectivos(matfechas(r), MatFechasEf, matfr0, matfr)
          For i = 1 To nofval
              For j = 1 To NoFactores
                  MDblTasas1(j, 1) = matfr(i, j)
              Next j
              res_efi(i).val_efect = 0
              parval.perfwd = MatFechasEf(i) - fecha
              matv = CalcValuacion(MatFechasEf(i), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MDblTasas1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
              Call CalcEfecProxySwap(res_efi, matv, i, indswap, indpswap)
              If i <> 1 Then
                 If res_efi(i).val_efect >= 0.8 And res_efi(i).val_efect <= 1.25 Then
                    matresef(i - 1, 4) = matresef(i - 1, 4) + 1
                 Else
                    matresef(i - 1, 4) = matresef(i - 1, 4) + 0
                 End If
              End If
              DoEvents
          Next i
      Next r
      suma = 0
      For i = 1 To nofval - 1
          suma = suma + matresef(i, 4)
      Next i
      efecpros = suma / ((nofval - 1) * noescen)
      MensajeProc = "La efectividad de la operacion " & coperacion & " fue del " & Format(efecpros, "##0.00 %")
      Call IniciarConexOracle(conAdo2, BDIKOS)
      Call GuardaResEfiPros(fecha, matpos(indswap).c_operacion, efecpros, conAdo2)
      SiIncTasaCVig = True
      conAdo2.Close
      Call GuardarResEfectPros(fecha, matpos(indswap).c_operacion, matresef)
      Call ValidarOperacion3(matpos(indswap).c_operacion, matpos(indswap).HoraRegOp, finicio, hinicio, Date, Time)
      txtmsg = "El proceso finalizo correctamente"
      exito = True
      Else
      txtmsg = "No se cargo correctamente la posicion primaria o el derivado"
      exito = False
      End If
   Else
      exito = False
      txtmsg = txtmsg2
   End If
Else
  MsgBox "No hay registros en la posicion simulada"
  exito = False
End If
Exit Sub
hayerror:
    MsgBox "CEficProsSwapsPort " & error(Err())
  exito = False
End Sub


Sub CEficProsSwapsPort(ByVal fecha As Date, ByVal txtport As String, ByRef txtmsg As String, ByRef exito As Boolean)
If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim exito1 As Boolean
Dim matresefic() As Variant
Dim mateficpros() As Variant
Dim parval1(1 To 2) As Variant
Dim parval2(1 To 4) As Variant
Dim matpact() As Variant
Dim matppas() As Variant
Dim matprimact() As Variant
Dim matprimpas() As Variant
Dim matidact() As Variant
Dim matidpas() As Variant
Dim matidprimact() As Variant
Dim matidprimpas() As Variant
Dim matres() As Variant
Dim noreg As Integer
Dim r As Integer
Dim i As Integer
Dim coperacion As String
Dim fcurva As Date
Dim fvenc As Date
Dim valora As Double
Dim valorb As Double
Dim valorc As Double
Dim valord As Double
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim valor4 As Double
Dim teficiencia As Integer
Dim mattxt() As String
Dim matpos() As New propPosRiesgo
Dim matposmd() As New propPosMD
Dim matposdiv() As New propPosDiv
Dim matposswaps() As New propPosSwaps
Dim matposfwd() As New propPosFwd
Dim matposdeuda() As New propPosDeuda
Dim matflswap() As New estFlujosDeuda
Dim matfldeuda() As New estFlujosDeuda
Dim matfactestres() As Double
Dim nofval As Integer
Dim noescen As Integer
Dim contart As Integer
Dim contar0 As Integer
Dim jj As Integer
Dim j As Integer
Dim valefi As Double
Dim valefi1 As Double
Dim valefi2 As Double
Dim contar3 As Integer
Dim efecpros As Double
Dim txtnomarch As String
Dim mata() As Variant
Dim hinicio As Date
Dim finicio As Date
Dim fechax As Date
Dim htiempo As Integer
Dim matfr() As Variant
Dim matx() As Variant
Dim matx1() As Double
Dim matfechas() As Date
Dim matrends() As Double
Dim matinctasa() As Double
Dim indswap As Long
Dim indppact As Long
Dim indpppas As Long
Dim res_efi() As New resEficSwap
Dim exito2 As Boolean
Dim txtmsg2 As String
Dim txtmsg0 As String
Dim exitoarch As Boolean

finicio = Date
hinicio = Time
mattxt = CrearFiltroPosPort(fecha, txtport)
Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, exito)
'se debe de saber que tipo de operaciones y como empatan los flujos
noreg = UBound(matpos, 1)
If noreg <> 0 Then
   Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg2, exito2)
   If exito2 Then
      Call DetermIndEfecCob(matpos, indswap, indppact, indpppas)
      teficiencia = DetTEfec(matpos, indppact, indpppas)
      If teficiencia = 0 Then
         txtmsg = "No es una relacion valida de efectividad"
         MsgBox txtmsg
         exito = False
         Exit Sub
      End If
      Dim MatFechasEf() As Date
      fvenc = matposswaps(1).FvencSwap
      fcurva = PBD1(fecha, 1, "MX")
      coperacion = matpos(1).c_operacion
      MatFechasEf = DefFechasEfCob(fecha, fvenc) 'se toma la fecha del swap
      nofval = UBound(MatFechasEf, 1)

'se obtienen las emisiones de los swaps
'en matresefefcob
      noescen = 10
      ReDim matfechasext(1 To noescen) As Date
      'escenarios negativos
      matfechasext(1) = #3/11/2011#
      matfechasext(2) = #9/2/2011#
      matfechasext(3) = #9/15/2011#
      matfechasext(4) = #9/12/2011#
      matfechasext(5) = #4/16/2012#
      'escenarios positivos
      matfechasext(6) = #10/30/2012#
      matfechasext(7) = #8/27/2012#
      matfechasext(8) = #9/10/2012#
      matfechasext(9) = #9/11/2012#
      matfechasext(10) = #4/27/2012#
      
      MatFactR1 = CargaFR1Dia(fcurva, exito)
      htiempo = 1
      ReDim matfactestres(1 To 10, 1 To NoFactores) As Double
      For i = 1 To noescen
          fechax = DetFechaFNoEsc(matfechasext(i), htiempo + 1)
          Call CrearMatFRiesgo2(fechax, matfechasext(i), matfr, "", exito)
          matx = ExtraerSMatFR(2, 2, matfr, True, SiFactorRiesgo)
          matfechas = ConvArVtDT(ExtraeSubMatrizV(matx, 1, 1, 1 + htiempo, UBound(matx, 1)))
          matx1 = ConvArVtDbl(ExtraeSubMatrizV(matx, 2, UBound(matx, 2), 1, UBound(matx, 1)))
          matrends = GenRends(matx1, htiempo, matfechas)
          ReDim matinctasa(1 To NoFactores, 1 To 1) As Double
          For j = 1 To NoFactores
              matinctasa(j, 1) = Abs(MatFactR1(j, 1)) * matrends(1, j)
              matfactestres(i, j) = MatFactR1(j, 1) + matinctasa(j, 1)
              DoEvents
          Next j
      Next i
      ReDim mateficpros(1 To nofval, 1 To noescen + 1) As Variant
      ReDim matrest(1 To 1) As Variant
      contart = 0
      For r = 1 To noescen    'el no de escenarios de estres
          Call CalcEfProspSwap2(r, matfactestres, MatFechasEf, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, teficiencia, res_efi, matres)
          ReDim Preserve matrest(1 To contart + 1) As Variant
          matrest(contart + 1) = "Escenario " & matfechasext(r)
          contart = contart + 1
          contar0 = UBound(matres, 1)
          ReDim Preserve matrest(1 To contart + contar0) As Variant
          For jj = 1 To contar0
              matrest(contart + jj) = matres(jj)
          Next jj
          For i = 2 To nofval
              mateficpros(i, 1) = MatFechasEf(i)
              mateficpros(i, r + 1) = 0
              If res_efi(i).val_efect >= 0.8 And res_efi(i).val_efect <= 1.25 Then
                 mateficpros(i, r + 1) = mateficpros(i, r + 1) + 1 'recuerda que se acumulan los exitos o fracasos
              Else
                 mateficpros(i, r + 1) = mateficpros(i, r + 1) + 0
              End If
              matrest(contart + i + 1) = matrest(contart + i + 1) & Chr(9) & res_efi(i).val_efect
          Next i
         contart = contart + contar0
      Next r
      Close #5
      ReDim matresefic(1 To nofval - 1, 1 To 3) As Variant
      contar3 = 0
      efecpros = 0
      For i = 1 To nofval - 1
          matresefic(i, 1) = MatFechasEf(i + 1)
          matresefic(i, 2) = noescen
          matresefic(i, 3) = 0
          For j = 1 To noescen
              If mateficpros(i + 1, j + 1) <> 0 Then
                 matresefic(i, 3) = matresefic(i, 3) + mateficpros(i + 1, j + 1)
              End If
              contar3 = contar3 + 1
          Next j
          efecpros = efecpros + matresefic(i, 3)
      Next i
      efecpros = efecpros / contar3
      MsgBox "La efectividad de la operacion " & coperacion & " fue del " & Format(efecpros, "##0.00 %")
      Call IniciarConexOracle(conAdo2, BDIKOS)
      Call GuardaResEfiPros(fecha, coperacion, efecpros, conAdo2)
      conAdo2.Close
      ReDim mata(1 To 1, 1 To 2) As Variant
      mata(1, 1) = coperacion
      mata(1, 2) = matpos(1).HoraRegOp
      Call ImpReporteEfProsSwap(fecha, matresefic, coperacion, efecpros)
'se debe de crear el archivo que contenga el resumen de la eficiencia de la cobertura
      txtnomarch = DirResVaR & "\Resultados eficiencia prospectiva operacion " & matpos(1).c_operacion & "  " & Format(fecha, "yyyymmdd") & ".txt"
      frmCalVar.CommonDialog1.FileName = txtnomarch
      frmCalVar.CommonDialog1.ShowSave
      txtnomarch = frmCalVar.CommonDialog1.FileName
      Call VerificarSalidaArchivo(txtnomarch, 5, exitoarch)
      If exitoarch Then
      For r = 1 To contart
          Print #5, matrest(r)
      Next r
      Close #5
      End If
      Call ValidarOperaciones2(mata, finicio, hinicio, Date, Time)
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      exito = False
      txtmsg = txtmsg2
   End If
Else
  MsgBox "No hay registros en la posicion simulada"
  exito = False
End If
Exit Sub
hayerror:
    MsgBox "CEficProsSwapsPort " & error(Err())
  exito = False
End Sub

Sub CEficProsSwapsPos(ByVal fecha As Date, _
                      ByRef matpos() As propPosRiesgo, _
                      ByRef matposmd() As propPosMD, _
                      ByRef matposdiv() As propPosDiv, _
                      ByRef matposswaps() As propPosSwaps, _
                      ByRef matposfwd() As propPosFwd, _
                      ByRef matflswap() As estFlujosDeuda, _
                      ByRef matposdeuda() As propPosDeuda, _
                      ByRef matfldeuda() As estFlujosDeuda, _
                      ByVal htiempo As Integer, ByRef txtmsg As String, ByRef exito As Boolean)

Dim matresefic() As Variant
Dim mateficpros() As Variant
Dim parval1(1 To 2) As Variant
Dim parval2(1 To 4) As Variant
Dim matpact() As Variant
Dim matppas() As Variant
Dim matprimact() As Variant
Dim matprimpas() As Variant
Dim matidact() As Variant
Dim matidpas() As Variant
Dim matidprimact() As Variant
Dim matidprimpas() As Variant
Dim matres() As Variant
Dim noreg As Integer
Dim r As Integer
Dim i As Integer
Dim coperacion As String
Dim fcurva As Date
Dim fvenc As Date
Dim valora As Double
Dim valorb As Double
Dim valorc As Double
Dim valord As Double
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim valor4 As Double
Dim teficiencia As Integer
Dim mattxt() As String
Dim matfactestres() As Double
Dim nofval As Integer
Dim noescen As Integer
Dim contart As Integer
Dim contar0 As Integer
Dim jj As Integer
Dim j As Integer
Dim contar3 As Integer
Dim efecpros As Double
Dim txtnomarch As String
Dim mata() As Variant
Dim hinicio As Date
Dim finicio As Date
Dim fechax As Date
Dim matfr() As Variant
Dim matx() As Variant
Dim matx1() As Double
Dim matfechas() As Date
Dim matrends() As Double
Dim matinctasa() As Double
Dim indswap As Long
Dim indppact As Long
Dim indpppas As Long
Dim exitoarch As Boolean

finicio = Date
hinicio = Time

'se debe de saber que tipo de operaciones y como empatan los flujos
noreg = UBound(matpos, 1)
If noreg <> 0 Then
   'matpos = RutinaOrden(matpos, CCOperacion, 3)
   Call DeterminaIndEficCob(matpos, indswap, indppact, indpppas)
   teficiencia = DetTEfec(matpos, indppact, indpppas)
   If teficiencia = 0 Then
      txtmsg = "No es una relacion valida de efectividad"
      MsgBox txtmsg
      exito = False
      Exit Sub
   End If
'FECHA       fecha del analisis
'matpos      posicion de swaps a analizar
'mateficpros contiene el desglose de los calculos realizados por fecha y escenario
'matresefic  contiene el resumen de los calculos

'se definen las fechas donde se realizara el calculo de valores de efectividad
'ojo si MatFlujosSwaps es la tabla completa se debe de cortar el no de
'simulaciones hasta la fecha de vencimiento del swap
'como generalmente la eficiencia prospectiva es de un solo swap
'se obtiene la fecha de
Dim MatFechasEf() As Date
fvenc = matposswaps(1).FvencSwap
fcurva = PBD1(fecha, 1, "MX")
coperacion = matpos(1).c_operacion
MatFechasEf = DefFechasEfCob(fecha, fvenc) 'se toma la fecha del swap
nofval = UBound(MatFechasEf, 1)

'se obtienen las emisiones de los swaps
'en matresefefcob
noescen = 10
ReDim matfechasext(1 To noescen) As Date
'escenarios negativos
matfechasext(1) = #3/11/2011#
matfechasext(2) = #9/2/2011#
matfechasext(3) = #9/15/2011#
matfechasext(4) = #9/12/2011#
matfechasext(5) = #4/16/2012#
'escenarios positivos
matfechasext(6) = #10/30/2012#
matfechasext(7) = #8/27/2012#
matfechasext(8) = #9/10/2012#
matfechasext(9) = #9/11/2012#
matfechasext(10) = #4/27/2012#

MatFactR1 = CargaFR1Dia(fcurva, exito)
ReDim matfactestres(1 To 10, 1 To NoFactores) As Double
For i = 1 To noescen
    fechax = DetFechaFNoEsc(matfechasext(i), htiempo + 1)
    Call CrearMatFRiesgo2(fechax, matfechasext(i), matfr, "", exito)
    matx = ExtraerSMatFR(htiempo + 1, htiempo + 1, matfr, True, SiFactorRiesgo)
    matfechas = ConvArVtDT(ExtraeSubMatrizV(matx, 1, 1, 1 + htiempo, UBound(matx, 1)))
    matx1 = ConvArVtDbl(ExtraeSubMatrizV(matx, 2, UBound(matx, 2), 1, UBound(matx, 1)))
    matrends = GenRends(matx1, htiempo, matfechas)
    ReDim matinctasa(1 To NoFactores, 1 To 1) As Double
    For j = 1 To NoFactores
         matinctasa(j, 1) = Abs(MatFactR1(j, 1)) * matrends(1, j)
         matfactestres(i, j) = MatFactR1(j, 1) + matinctasa(j, 1)
         DoEvents
    Next j
Next i

ReDim mateficpros(1 To nofval, 1 To noescen + 1) As Variant
ReDim matrest(1 To 1) As Variant
ReDim res_efi(1 To nofval) As New resEficSwap
contart = 0
For r = 1 To noescen    'el no de escenarios de estres
    Call CalcEfProspSwap(r, matfactestres, MatFechasEf, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, teficiencia, res_efi, matres, indswap, indppact, indpppas)
    ReDim Preserve matrest(1 To contart + 1) As Variant
    matrest(contart + 1) = "Escenario " & matfechasext(r)
    contart = contart + 1
    contar0 = UBound(matres, 1)
    ReDim Preserve matrest(1 To contart + contar0) As Variant
    For jj = 1 To contar0
        matrest(contart + jj) = matres(jj)
    Next jj
    For i = 2 To nofval
        mateficpros(i, 1) = MatFechasEf(i)
        mateficpros(i, r + 1) = 0
        If res_efi(i).val_efect >= 0.8 And res_efi(i).val_efect <= 1.25 Then
           mateficpros(i, r + 1) = mateficpros(i, r + 1) + 1
        Else
           mateficpros(i, r + 1) = mateficpros(i, r + 1) + 0
        End If
        matrest(contart + i + 1) = matrest(contart + i + 1) & Chr(9) & res_efi(i).val_efect
    Next i
    contart = contart + contar0
Next r
Close #5
'en esta matriz se ponen los resultados de la eficiencia prospectiva
ReDim matresefic(1 To nofval - 1, 1 To 3) As Variant
contar3 = 0
efecpros = 0
For i = 1 To nofval - 1
    matresefic(i, 1) = MatFechasEf(i + 1)
    matresefic(i, 2) = noescen
    matresefic(i, 3) = 0
    For j = 1 To noescen
    If mateficpros(i + 1, j + 1) <> 0 Then
          matresefic(i, 3) = matresefic(i, 3) + mateficpros(i + 1, j + 1)
        End If
        contar3 = contar3 + 1
    Next j
    efecpros = efecpros + matresefic(i, 3)
Next i
efecpros = efecpros / contar3
MsgBox "La efectividad de la operacion " & coperacion & " fue del " & Format(efecpros, "##0.00 %")
'Call IniciarConexOracle(conAdo2, BDIKOS)
'Call GuardaResEfiPros(fecha, coperacion, efecpros, conAdo2)
'conAdo2.Close
ReDim mata(1 To 1, 1 To 2) As Variant
mata(1, 1) = coperacion
mata(1, 2) = matpos(1).HoraRegOp
'Call ImpReporteEfProsSwap(fecha, matresefic, coperacion, efecpros)
'se debe de crear el archivo que contenga el resumen de la eficiencia de la cobertura
txtnomarch = DirResVaR & "\Resultados eficiencia prospectiva operacion " & matpos(1).c_operacion & "  " & Format(fecha, "yyyymmdd") & ".txt"
frmCalVar.CommonDialog1.FileName = txtnomarch
frmCalVar.CommonDialog1.ShowSave
txtnomarch = frmCalVar.CommonDialog1.FileName
Call VerificarSalidaArchivo(txtnomarch, 5, exitoarch)
If exitoarch Then
For r = 1 To contart
 Print #5, matrest(r)
Next r
Close #5
End If
  txtmsg = "El proceso finalizo correctamente"
  exito = True
Else
  exito = False
End If

End Sub

Sub DeterminaIndEficCob(ByRef matpos() As propPosRiesgo, ByRef indswap As Long, ByRef indppact As Long, ByRef indpppas As Long)
Dim i As Long
For i = 1 To UBound(matpos, 1)
    If matpos(i).fValuacion = "SWAP" Then
      indswap = i
    End If
    If matpos(i).fValuacion = "DEUDA" And matpos(i).Signo_Op = 1 Then
      indppact = i
    End If
    If matpos(i).fValuacion = "DEUDA" And matpos(i).Signo_Op = -1 Then
       indpppas = i
    End If
Next i
End Sub

Sub CalcEfProspSwap(ByVal indice As Integer, _
                    ByRef matfr() As Double, _
                    ByRef matfechas() As Date, _
                    ByRef matpos() As propPosRiesgo, _
                    ByRef matposmd() As propPosMD, _
                    ByRef matposdiv() As propPosDiv, _
                    ByRef matposswaps() As propPosSwaps, _
                    ByRef matposfwd() As propPosFwd, _
                    ByRef matflswap() As estFlujosDeuda, _
                    ByRef matposdeuda() As propPosDeuda, _
                    ByRef matfldeuda() As estFlujosDeuda, _
                    ByRef tcalculoef As Integer, ByRef res_efi() As resEficSwap, ByRef matres() As Variant, ByVal indswap As Long, ByVal indppact As Long, ByVal indpppas As Long)

Dim noreg As Integer
Dim sicargarf As Boolean
Dim MDblTasas1() As Double
Dim parval As ParamValPos
Dim mrvalflujo() As resValFlujo
Dim matfact() As Variant
Dim exito As Boolean
Dim contar As Integer
Dim nofval As Integer
Dim i As Integer
Dim j As Integer
Dim nodim1 As Integer
Dim matv() As New resValIns
Dim txtcadena As String
Dim coperacion As String
Dim exito3 As Boolean
Dim txtmsg3 As String

contar = 0
'se realiza el calculo de la eficiencia prospectiva para una sola fecha
'se toma la matriz de riesgo
noreg = UBound(matpos, 1)

sicargarf = True
ValExacta = False
nofval = UBound(matfechas, 1)
ReDim matres(1 To nofval + 1) As Variant  'aqui se guardan los resultados
Set parval = DeterminaPerfilVal("EFECTIVIDAD")
'lo primero es establecer las fechas en las que se va a valuar el swap
'se pone primero la fecha en cada caso
ReDim res_efis(1 To nofval) As New resEficSwap

'se cargan los escenarios de estres en una matriz de datos
For i = 1 To nofval
'se procede a estresar el escenario tabla
    ReDim MDblTasas1(1 To NoFactores, 1 To 1) As Double
    For j = 1 To NoFactores
        MDblTasas1(j, 1) = matfr(indice, j)
    Next j
'realiza la valuacion de los flujos
    res_efis(i).val_efect = 0
    matv = CalcValuacion(matfechas(i), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MDblTasas1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
'en matv estan todas las valuaciones
    Call CalculosEfecSwap(tcalculoef, res_efis, matv, i, indswap, indppact, indpppas)
    MensajeProc = "Procesando el día " & matfechas(i)
    DoEvents
Next i
noreg = UBound(matpos, 1)
coperacion = matpos(1).c_operacion
'se imprimen los resultados
'primero los encabezados
txtcadena = "Fecha " & Chr(9)
If tcalculoef = 1 Then
   txtcadena = txtcadena & coperacion & " SWAP activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " SWAP pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " PRIMARIA activa*" & Chr(9)
   txtcadena = txtcadena & coperacion & " PRIMARIA pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int. Dev activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int. Dev pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int. Dev p activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int. Dev p pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Efectividad periodo"
ElseIf tcalculoef = 2 Then
   txtcadena = txtcadena & coperacion & " SWAP pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " PRIMARIA activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev p activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Efectividad periodo"
ElseIf tcalculoef = 3 Then
   txtcadena = txtcadena & coperacion & " SWAP activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " SWAP pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " PRIMARIA activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " PRIMARIA pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev p activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev p pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Efectividad periodo"
End If
matres(1) = txtcadena
'las valuaciones
For i = 1 To nofval
    txtcadena = matfechas(i) & Chr(9)
    If tcalculoef = 1 Then
       txtcadena = txtcadena & res_efis(i).v_swapact & Chr(9)
       txtcadena = txtcadena & res_efis(i).v_swappas & Chr(9)
       txtcadena = txtcadena & res_efis(i).v_primact & Chr(9)
       txtcadena = txtcadena & res_efis(i).v_primpas & Chr(9)
       txtcadena = txtcadena & res_efis(i).id_swapact & Chr(9)
       txtcadena = txtcadena & res_efis(i).id_swappas & Chr(9)
       txtcadena = txtcadena & res_efis(i).id_primact & Chr(9)
       txtcadena = txtcadena & res_efis(i).id_primpas & Chr(9)
    ElseIf tcalculoef = 2 Then
       txtcadena = txtcadena & res_efis(i).v_swappas & Chr(9)
       txtcadena = txtcadena & res_efis(i).v_primact & Chr(9)
       txtcadena = txtcadena & res_efis(i).id_swappas & Chr(9)
       txtcadena = txtcadena & res_efis(i).id_primact & Chr(9)
    ElseIf tcalculoef = 3 Then
       txtcadena = txtcadena & res_efis(i).v_swapact & Chr(9)
       txtcadena = txtcadena & res_efis(i).v_swappas & Chr(9)
       txtcadena = txtcadena & res_efis(i).v_primact & Chr(9)
       txtcadena = txtcadena & res_efis(i).v_primpas & Chr(9)
       txtcadena = txtcadena & res_efis(i).id_swapact & Chr(9)
       txtcadena = txtcadena & res_efis(i).id_swappas & Chr(9)
       txtcadena = txtcadena & res_efis(i).id_primact & Chr(9)
       txtcadena = txtcadena & res_efis(i).id_primpas & Chr(9)
    End If
    matres(i + 1) = txtcadena
Next i
 res_efi = res_efis
End Sub

Sub CalcEfProspSwap2(ByVal indice As Integer, _
                     ByRef matfr() As Double, _
                     ByRef matfechas() As Date, _
                     ByRef matpos() As propPosRiesgo, _
                     ByRef matposmd() As propPosMD, _
                     ByRef matposdiv() As propPosDiv, _
                     ByRef matposswaps() As propPosSwaps, _
                     ByRef matposfwd() As propPosFwd, _
                     ByRef matflswap() As estFlujosDeuda, _
                     ByRef matposdeuda() As propPosDeuda, _
                     ByRef matfldeuda() As estFlujosDeuda, _
                     ByRef tcalculoef As Integer, ByRef res_efi() As resEficSwap, ByRef matres() As Variant)

If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim noreg As Integer
Dim MDblTasas1() As Double
Dim mrvalflujo() As resValFlujo
Dim matfact() As Variant
Dim nofval As Integer
Dim i As Integer
Dim j As Integer
Dim matv() As New resValIns
Dim txtcadena As String
Dim coperacion As String
Dim parval As New ParamValPos
Dim indswap As Long
Dim indpact As Long
Dim indppas As Long
Dim res_efi1() As New resEficSwap
Dim exito3 As Boolean
Dim txtmsg3 As String


'se realiza el calculo de la eficiencia prospectiva para una sola fecha
'se toma la matriz de riesgo
noreg = UBound(matpos, 1)
ValExacta = False
nofval = UBound(matfechas, 1)
ReDim matres(1 To nofval + 1) As Variant  'aqui se guardan los resultados
Set parval = DeterminaPerfilVal("EFECTIVIDAD")
Call DetermIndEfecCob(matpos, indswap, indpact, indppas)
'lo primero es establecer las fechas en las que se va a valuar el swap
'se pone primero la fecha en cada caso
ReDim res_efi1(1 To nofval) As New resEficSwap

'se cargan los escenarios de estres en una matriz de datos
For i = 1 To nofval
'se procede a estresar el escenario tabla
    ReDim MDblTasas1(1 To NoFactores, 1 To 1) As Double
    For j = 1 To NoFactores
        MDblTasas1(j, 1) = matfr(indice, j)
    Next j
'realiza la valuacion de los flujos
    res_efi1(i).val_efect = 0
    matv = CalcValuacion(matfechas(i), matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MDblTasas1, MatCurvasT, parval, mrvalflujo, txtmsg3, exito3)
'en matv estan todas las valuaciones
    Call CalculosEfecSwap(tcalculoef, res_efi1, matv, i, indswap, indpact, indppas)
    MensajeProc = "Procesando el día " & matfechas(i)
    DoEvents
Next i
coperacion = matpos(indswap).c_operacion
'se imprimen los resultados
'primero los encabezados
txtcadena = "Fecha " & Chr(9)
If tcalculoef = 1 Then
   txtcadena = txtcadena & coperacion & " SWAP activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " SWAP pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " PRIMARIA activa*" & Chr(9)
   txtcadena = txtcadena & coperacion & " PRIMARIA pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int. Dev activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int. Dev pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int. Dev p activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int. Dev p pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Efectividad periodo"
ElseIf tcalculoef = 2 Then
   txtcadena = txtcadena & coperacion & " SWAP pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " PRIMARIA activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev p activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Efectividad periodo"
ElseIf tcalculoef = 3 Then
   txtcadena = txtcadena & coperacion & " SWAP activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " SWAP pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " PRIMARIA activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " PRIMARIA pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev p activa" & Chr(9)
   txtcadena = txtcadena & coperacion & " Int Dev p pasiva" & Chr(9)
   txtcadena = txtcadena & coperacion & " Efectividad periodo"
End If
matres(1) = txtcadena
'las valuaciones
For i = 1 To nofval
    txtcadena = matfechas(i) & Chr(9)
    If tcalculoef = 1 Then
       txtcadena = txtcadena & res_efi1(i).v_swapact & Chr(9)
       txtcadena = txtcadena & res_efi1(i).v_swappas & Chr(9)
       txtcadena = txtcadena & res_efi1(i).v_primact & Chr(9)
       txtcadena = txtcadena & res_efi1(i).v_primpas & Chr(9)
       txtcadena = txtcadena & res_efi1(i).id_swapact & Chr(9)
       txtcadena = txtcadena & res_efi1(i).id_swappas & Chr(9)
       txtcadena = txtcadena & res_efi1(i).id_primact & Chr(9)
       txtcadena = txtcadena & res_efi1(i).id_primpas & Chr(9)
    ElseIf tcalculoef = 2 Then
       txtcadena = txtcadena & res_efi1(i).v_swappas & Chr(9)
       txtcadena = txtcadena & res_efi1(i).v_primact & Chr(9)
       txtcadena = txtcadena & res_efi1(i).id_swappas & Chr(9)
       txtcadena = txtcadena & res_efi1(i).id_primact & Chr(9)
    ElseIf tcalculoef = 3 Then
       txtcadena = txtcadena & res_efi1(i).v_swapact & Chr(9)
       txtcadena = txtcadena & res_efi1(i).v_swappas & Chr(9)
       txtcadena = txtcadena & res_efi1(i).v_primact & Chr(9)
       txtcadena = txtcadena & res_efi1(i).v_primpas & Chr(9)
       txtcadena = txtcadena & res_efi1(i).id_swapact & Chr(9)
       txtcadena = txtcadena & res_efi1(i).id_swappas & Chr(9)
       txtcadena = txtcadena & res_efi1(i).id_primact & Chr(9)
       txtcadena = txtcadena & res_efi1(i).id_primpas & Chr(9)
    End If
    matres(i + 1) = txtcadena
Next i
res_efi = res_efi1
On Error GoTo 0
Exit Sub
hayerror:
MsgBox "CalcEfProspSwap2 " & error(Err())
End Sub

Sub GuardaResEfiPros(ByVal fecha As Date, ByVal coperacion As String, ByVal eficpros As Double, ByRef obj1 As ADODB.Connection)
On Error GoTo hayerror
Dim txtfecha As String
Dim txtfiltro As String
Dim noreg As Integer
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT COUNT(*) from " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha & "AND CLAVE_SWAP = '" & coperacion & "'"
rmesa.Open txtfiltro, obj1
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 txtcadena = "UPDATE " & TablaEficienciaCob & " SET EFIC_PRO = " & 100 * eficpros & ", EFIC_RETRO = " & 100 * eficpros & " where (FECHA = " & txtfecha & " AND CLAVE_SWAP = '" & coperacion & "')"
 obj1.Execute txtcadena
Else
 txtcadena = "INSERT INTO " & TablaEficienciaCob & " VALUES("
 txtcadena = txtcadena & txtfecha & ","                        'fecha del calculo de la eficiencia
 txtcadena = txtcadena & txtfecha & ","                        'fecha de inicio del analisis
 txtcadena = txtcadena & txtfecha & ","                        'fecha final del analisis
 txtcadena = txtcadena & "'" & coperacion & "',"    '          la Clave de operación
 'If MatResEficSwaps(i, 12) <> 0 Then
 txtcadena = txtcadena & "0,"                                 'val p activa t0
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 13) <> 0 Then
 txtcadena = txtcadena & "0,"                              'val p pasiva t0
' Else
 ' txtcadena = txtcadena & "0,"
 'End If
 'If MatResEficSwaps(i, 16) <> 0 Then
  txtcadena = txtcadena & "0,"                             'val p activa t1
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 17) <> 0 Then
  txtcadena = txtcadena & "0,"                              'val p pasiva t1
' Else
'  txtcadena = txtcadena & "0,"
' End If
 
 'If MatResEficSwaps(i, 11) <> 0 Then
  txtcadena = txtcadena & "0,"                              'val p primaria1 t0
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 15) <> 0 Then
  txtcadena = txtcadena & "0,"                               'val p primaria1 t1
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 14) <> 0 Then
  txtcadena = txtcadena & "0,"                              'val p primaria2 t0
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 18) <> 0 Then
  txtcadena = txtcadena & "0,"                               'val p primaria2 t1
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 19) <> 0 Then
 txtcadena = txtcadena & "0,"                                'ajuste primaria1
' Else
' txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 20) <> 0 Then
  txtcadena = txtcadena & "0,"                               'ajuste activa
' Else
'  txtcadena = txtcadena & "0,"
' End If
 'If MatResEficSwaps(i, 21) <> 0 Then
  txtcadena = txtcadena & "0,"                                 'ajuste pasiva
 'Else
 ' txtcadena = txtcadena & "0,"
 'End If
 'If MatResEficSwaps(i, 22) <> 0 Then
  txtcadena = txtcadena & "0,"        'ajuste primaria2
 'Else
 ' txtcadena = txtcadena & "0,"
 'End If
  txtcadena = txtcadena & Format(eficpros * 100, "###.00") & ","    'eficiencia retrospectiva
  txtcadena = txtcadena & Format(eficpros * 100, "###.00") & ")"    'eficiencia pro
  obj1.Execute txtcadena
End If
On Error GoTo 0
Exit Sub
hayerror:
MsgBox "guardaresefipros " & error(Err())
End Sub

Sub GuardaResEficRetro(ByVal fecha As Date, ByVal coperacion As String, ByVal efic As Double, ByRef obj1 As ADODB.Connection)
If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim txtfecha As String
Dim txtfiltro As String
Dim noreg As Integer
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT COUNT(*) from " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha & "AND CLAVE_SWAP = '" & coperacion & "'"
rmesa.Open txtfiltro, obj1
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
 txtcadena = "UPDATE " & TablaEficienciaCob & " SET EFIC_PRO = " & 100 * efic & ", EFIC_RETRO = " & 100 * efic & " where (FECHA = " & txtfecha & " AND CLAVE_SWAP = '" & coperacion & "')"
 obj1.Execute txtcadena
Else
 txtcadena = "INSERT INTO " & TablaEficienciaCob & " VALUES("
 txtcadena = txtcadena & txtfecha & ","                        'fecha del calculo de la eficiencia
 txtcadena = txtcadena & txtfecha & ","                        'fecha de inicio del analisis
 txtcadena = txtcadena & txtfecha & ","                        'fecha final del analisis
 txtcadena = txtcadena & "'" & coperacion & "',"    '          la Clave de operación
 'If MatResEficSwaps(i, 12) <> 0 Then
 txtcadena = txtcadena & "0,"                                 'val p activa t0
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 13) <> 0 Then
 txtcadena = txtcadena & "0,"                              'val p pasiva t0
' Else
 ' txtcadena = txtcadena & "0,"
 'End If
 'If MatResEficSwaps(i, 16) <> 0 Then
  txtcadena = txtcadena & "0,"                             'val p activa t1
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 17) <> 0 Then
  txtcadena = txtcadena & "0,"                              'val p pasiva t1
' Else
'  txtcadena = txtcadena & "0,"
' End If
 
 'If MatResEficSwaps(i, 11) <> 0 Then
  txtcadena = txtcadena & "0,"                              'val p primaria1 t0
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 15) <> 0 Then
  txtcadena = txtcadena & "0,"                               'val p primaria1 t1
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 14) <> 0 Then
  txtcadena = txtcadena & "0,"                              'val p primaria2 t0
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 18) <> 0 Then
  txtcadena = txtcadena & "0,"                               'val p primaria2 t1
' Else
'  txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 19) <> 0 Then
 txtcadena = txtcadena & "0,"                                'ajuste primaria1
' Else
' txtcadena = txtcadena & "0,"
' End If
' If MatResEficSwaps(i, 20) <> 0 Then
  txtcadena = txtcadena & "0,"                               'ajuste activa
' Else
'  txtcadena = txtcadena & "0,"
' End If
 'If MatResEficSwaps(i, 21) <> 0 Then
  txtcadena = txtcadena & "0,"                                 'ajuste pasiva
 'Else
 ' txtcadena = txtcadena & "0,"
 'End If
 'If MatResEficSwaps(i, 22) <> 0 Then
  txtcadena = txtcadena & "0,"        'ajuste primaria2
 'Else
 ' txtcadena = txtcadena & "0,"
 'End If
  txtcadena = txtcadena & Format(efic * 100, "###.00") & ","    'eficiencia retrospectiva
  txtcadena = txtcadena & Format(efic * 100, "###.00") & ")"    'eficiencia pro
  obj1.Execute txtcadena
End If
On Error GoTo 0
Exit Sub
hayerror:
MsgBox "guardaresefipros " & error(Err())
End Sub


Function LeerArchCurvas(ByVal nomarch As String) As Variant()
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
 
 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
 Set registros1 = base1.OpenRecordset("Sheet1$", dbOpenDynaset, dbReadOnly)
 
 'se revisa si hay registros en la tabla
If registros1.RecordCount <> 0 Then
 registros1.MoveLast
 noreg = registros1.RecordCount
 registros1.MoveFirst
 nocampos = registros1.Fields.Count
 ReDim mata(1 To noreg, 1 To nocampos) As Variant
 For i = 1 To noreg
  For j = 1 To nocampos
  mata(i, j) = LeerTAccess(registros1, j - 1, i)
  Next j
  registros1.MoveNext
  AvanceProc = i / noreg
  MensajeProc = "Leyendo las curvas del dia " & Format(AvanceProc, "##0.00 %")
  DoEvents
 Next i
 registros1.Close
 LeerArchCurvas = mata
End If
End Function

Function LeerFactoresExcel(fecha1 As Date, fecha2 As Date) As Variant()
Dim exito As Boolean
Dim mfrexcel() As Variant
Dim noreg As Long
Dim noreg1 As Long
Dim noreg2 As Long
Dim noreg3 As Long
Dim noreg4 As Long
Dim i As Long
Dim j As Long
Dim mat1() As propCurva
Dim mat2() As propCurva
Dim mat3() As propCurva
Dim mat4() As propCurva


'se leen los factores de riesgo asociados a esta matriz
Call CrearMatFRiesgo2(fecha1, fecha2, mfrexcel, "", exito)
noreg = UBound(mfrexcel, 1)
' se deben de obtener las curvas de
If noreg <> 0 Then
 mat1 = CrearCurvaNodos("DESC IRS", 1, mfrexcel)
 noreg1 = UBound(mat1, 1)
 mat2 = CrearCurvaNodos("CETES IMP", 1, mfrexcel)
 noreg2 = UBound(mat2, 1)
 mat3 = CrearCurvaNodos("LIBOR", 1, mfrexcel)
 noreg3 = UBound(mat3, 1)
 mat4 = CrearCurvaNodos("BONDES D", 1, mfrexcel)
 noreg4 = UBound(mat4, 1)
 ReDim matr(1 To noreg, 1 To noreg1 + noreg2 + noreg3 + noreg4 + 1) As Variant
 For i = 1 To noreg
     matr(i, 1) = CDbl(mfrexcel(i, 1))
     mat1 = CrearCurvaNodos("DESC IRS", i, mfrexcel)
     noreg1 = UBound(mat1, 1)
     mat2 = CrearCurvaNodos("CETES IMP", i, mfrexcel)
     noreg2 = UBound(mat2, 1)
     mat3 = CrearCurvaNodos("LIBOR", i, mfrexcel)
     noreg3 = UBound(mat3, 1)
     mat4 = CrearCurvaNodos("BONDES D", i, mfrexcel)
     noreg4 = UBound(mat4, 1)
     For j = 1 To noreg1
         matr(i, j + 1) = mat1(j, 1)
     Next j
     For j = 1 To noreg2
         matr(i, j + noreg1 + 1) = mat2(j, 1)
     Next j
     For j = 1 To noreg3
         matr(i, j + noreg1 + noreg2 + 1) = mat3(j, 1)
     Next j
     For j = 1 To noreg4
         matr(i, j + noreg1 + noreg2 + noreg3 + 1) = mat4(j, 1)
     Next j
     Call MostrarMensajeSistema("Leyendo las curvas del dia " & Format(AvanceProc, "#,##0.00 %"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
     DoEvents
 Next i
Else
     ReDim matr(0 To 0, 0 To 0) As Variant
End If
LeerFactoresExcel = matr
End Function

Sub ActFactoresExcel(ByRef mata() As Variant, ByVal nomtabla As String, ByRef conex As ADODB.Connection, ByRef rbase As ADODB.recordset)
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim noreg1 As Long
Dim txtfiltro As String
Dim txtcadena As String

'se leen los factores de riesgo asociados a esta matriz
' se deben de obtener las curvas de
noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
If UBound(mata, 1) > 0 Then
   For i = 1 To UBound(mata, 1)
       txtfiltro = "SELECT COUNT(*) FROM [" & nomtabla & "] WHERE FECHA = " & CLng(mata(i, 1))
       rbase.Open txtfiltro, conex
       noreg1 = rbase.Fields(0)
       rbase.Close
       If noreg1 <> 0 Then
       
       Else
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & CLng(mata(i, 1)) & ","
          For j = 2 To nocampos - 1
              txtcadena = txtcadena & Val(mata(i, j)) & ","
          Next j
          txtcadena = txtcadena & Val(mata(i, nocampos)) & ")"
          conex.Execute txtcadena
       End If

   Next i
End If
End Sub

Sub ExportaBackExcel(ByVal fecha As Date, ByRef mata() As Double, ByVal nomtabla As String, conex, rbase)
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Integer
Dim j As Integer
Dim noreg1 As Integer

noreg = UBound(mata, 1)
If noreg > 0 Then
       txtfiltro = "SELECT COUNT(*) FROM [" & nomtabla & "] WHERE FECHA = " & CLng(fecha)
       rbase.Open txtfiltro, conex
       noreg1 = rbase.Fields(0)
       rbase.Close
       If noreg1 <> 0 Then
       
       Else
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & CLng(fecha) & ","
          For j = 1 To noreg - 1
              txtcadena = txtcadena & Val(mata(j)) & ","
          Next j
          txtcadena = txtcadena & Val(mata(noreg)) & ")"
          conex.Execute txtcadena
       End If
End If
End Sub

Function LeerEscHistRE(ByVal fecha As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtsubport As String, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim valor As String
Dim mata() As String
Dim matv() As Variant
Dim noreg As Integer
Dim noreg1 As Integer
Dim i As Integer
Dim indice As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & txtsubport & "'"
txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo
txtfiltro1 = "SELECT COUNT(*) FROM  (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
   mata = EncontrarSubCadenas(valor, ",")
   noreg1 = UBound(mata, 1)
   ReDim matv(1 To noreg1, 1 To 2) As Variant
   indice = BuscarValorArray(fecha, MatFechasVaR, 1)
   For i = 1 To noreg1
       matv(i, 1) = MatFechasVaR(indice - noreg1 + i, 1)
       matv(i, 2) = CDbl(mata(i))
   Next i
   rmesa.Close
Else
   ReDim matv(0 To 0, 0 To 0) As Variant
End If
LeerEscHistRE = matv
End Function

Sub GuardaEscHistExcel(ByVal fecha As Date, ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef mata() As Variant, ByVal nomtabla As String, ByVal noesc As Integer, conex, rbase)
Dim noreg As Integer
Dim noreg1 As Integer
Dim nocampos As Integer
Dim txtfiltro As String
Dim txtcadena As String
Dim j As Integer

noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
If noreg > 0 Then
       txtfiltro = "SELECT COUNT(*) FROM [" & nomtabla & "] WHERE FECHA = " & CLng(fecha)
       rbase.Open txtfiltro, conex
       noreg1 = rbase.Fields(0)
       rbase.Close
       If noreg1 <> 0 Then
       
       Else
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & CLng(fecha) & ","
          txtcadena = txtcadena & CLng(fecha1) & ","
          txtcadena = txtcadena & CLng(fecha2) & ","
          For j = 1 To noesc
              txtcadena = txtcadena & CLng(mata(j, 1)) & ","
          Next j
          For j = 1 To noesc - 1
              txtcadena = txtcadena & Val(mata(j, 2)) & ","
          Next j
          txtcadena = txtcadena & Val(mata(noesc, 2)) & ")"
          conex.Execute txtcadena
       End If
End If
End Sub

Sub GuardaValContraparte(mata, nomtabla, conex, rbase)
Dim noreg As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String

noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
If noreg > 0 Then
   For i = 1 To noreg
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & "'" & mata(i, 1) & "',"
          For j = 2 To nocampos - 1
              txtcadena = txtcadena & Val(mata(i, j)) & ","
          Next j
          txtcadena = txtcadena & Val(mata(i, j)) & ")"
          conex.Execute txtcadena
          DoEvents
   Next i
End If
End Sub

Sub GuardaReEscEstresExcel(mata, nomtabla, conex, rbase)
Dim noreg As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String

noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
If noreg > 0 Then
   For i = 1 To noreg
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & "'" & mata(i, 1) & "',"
          txtcadena = txtcadena & mata(i, 2) & ","
          txtcadena = txtcadena & mata(i, 3) & ","
          txtcadena = txtcadena & "'" & mata(i, 4) & "',"
          For j = 5 To nocampos - 1
              txtcadena = txtcadena & Val(mata(i, j)) & ","
          Next j
          txtcadena = txtcadena & Val(mata(i, j)) & ")"
          conex.Execute txtcadena
          DoEvents
   Next i
End If
End Sub

Sub GuardaResEstruct1(mata, nomtabla, conex, rbase)
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String

noreg = UBound(mata, 1)

If noreg > 0 Then
   For i = 1 To noreg
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & "'" & mata(i, 1) & "',"
          txtcadena = txtcadena & ConvValor(mata(i, 2)) & ","
          txtcadena = txtcadena & ConvValor(mata(i, 3)) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 4), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 5), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 6), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 7), 0) & ")"
          conex.Execute txtcadena
          DoEvents
   Next i
End If
End Sub

Sub GuardaResEstruct2(mata, nomtabla, conex, rbase)
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String

noreg = UBound(mata, 1)

If noreg > 0 Then
   For i = 1 To noreg
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & "'" & mata(i, 1) & "',"
          txtcadena = txtcadena & ConvValor(mata(i, 2)) & ","
          txtcadena = txtcadena & ConvValor(mata(i, 3)) & ")"
          conex.Execute txtcadena
          DoEvents
   Next i
End If
End Sub


Sub GuardaResPIDVDer(mata, nomtabla, conex, rbase)
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim txtcadena As String

noreg = UBound(mata, 1)

If noreg > 0 Then
   For i = 1 To noreg
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & "'" & mata(i, 1) & "',"
          If Not EsVariableVacia(mata(i, 3)) Then
           txtcadena = txtcadena & "'" & mata(i, 2) & "',"
          Else
           txtcadena = txtcadena & "' ',"
          End If
          If Not EsVariableVacia(mata(i, 3)) Then
           txtcadena = txtcadena & "''" & mata(i, 3) & "',"
          Else
           txtcadena = txtcadena & "' ',"
          End If
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 4), 0) & ","
          txtcadena = txtcadena & ReemplazaVacioValor(mata(i, 5), 0) & ")"
          conex.Execute txtcadena
          DoEvents
   Next i
End If
End Sub

Sub GuardaResCVaRMarginal(mata, nomtabla, conex, rbase)
Dim noreg As Integer
Dim i As Integer
Dim txtcadena As String

noreg = UBound(mata, 1)
If noreg > 0 Then
   For i = 1 To noreg
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & CLng(mata(i, 1)) & ","
          txtcadena = txtcadena & mata(i, 2) & ","
          txtcadena = txtcadena & mata(i, 3) & ","
          txtcadena = txtcadena & mata(i, 4) & ","
          txtcadena = txtcadena & mata(i, 5) & ","
          txtcadena = txtcadena & mata(i, 6) & ")"
          conex.Execute txtcadena
          DoEvents
   Next i
End If
End Sub




Sub GuardaResPosExcel(ByVal fecha As Date, ByRef mata() As Variant, ByVal nomtabla As String, ByRef conex As ADODB.Connection, ByRef rbase As ADODB.recordset)
Dim txtcadena As String
Dim txtfiltro As String
Dim i As Integer
Dim j As Integer
Dim noreg As Integer
Dim noreg1 As Integer
Dim nocampos As Integer

noreg = UBound(mata, 1)
nocampos = UBound(mata, 2)
If noreg <> 0 Then
       txtfiltro = "SELECT COUNT(*) FROM [" & nomtabla & "] WHERE FECHA = " & CLng(fecha)
       rbase.Open txtfiltro, conex
       noreg1 = rbase.Fields(0)
       rbase.Close
       If noreg1 <> 0 Then
       
       Else
          txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
          txtcadena = txtcadena & CLng(fecha) & ","
          For i = 1 To noreg - 1
          For j = 1 To nocampos
              txtcadena = txtcadena & Val(mata(i, j)) & ","
          Next j
          Next i
          For j = 1 To nocampos - 1
              txtcadena = txtcadena & Val(mata(noreg, j)) & ","
          Next j
          txtcadena = txtcadena & Val(mata(noreg, nocampos)) & ")"
          conex.Execute txtcadena
       End If

End If
End Sub

Sub ActPosPortExcel(ByVal fecha As Date, ByRef mata() As Variant, ByVal nomtabla As String, ByVal txtport As String, ByRef objeto1 As DAO.Database)
Dim noreg As Long
Dim i As Long
Dim registros1 As DAO.recordset

 noreg = UBound(mata, 1)
If noreg <> 0 Then
 Set registros1 = objeto1.OpenRecordset(nomtabla, dbOpenDynaset)
 registros1.FindFirst "FECHA = " & CDbl(fecha)
 If registros1.NoMatch Then
  registros1.AddNew
 Else
  registros1.Edit
 End If
 For i = 1 To noreg
 If mata(i, 3) = txtport And mata(i, 4) = "TPA" Then
  Call GrabarTAccess(registros1, 1, 0, i)            'posicion activa
  Call GrabarTAccess(registros1, 1, mata(i, 5), i)   'posicion activa
 End If
 If mata(i, 3) = txtport And mata(i, 4) = "TPP" Then
  Call GrabarTAccess(registros1, 2, 0, i)   'posicion pasiva
  Call GrabarTAccess(registros1, 2, mata(i, 5), i)   'posicion pasiva
 End If
 If mata(i, 3) = txtport And mata(i, 4) = "MM" Then
  Call GrabarTAccess(registros1, 3, 0, i)  'marca mercado
  Call GrabarTAccess(registros1, 3, mata(i, 5), i)  'marca mercado
 End If
Next i
MensajeProc = "Actualizando la posicion de " & txtport & " del " & Format(fecha, "dd/mm/yyyy")
registros1.Update
registros1.Close
End If
End Sub

Function LeerValDerivadosExcel(ByVal fecha As Date, id_val As Integer) As Double()

'se exporta la valuacion del portafolio completo asi como la valuacion de los
'swaps del estado de mexico y el swap de yenes
Dim matport(1 To 6) As String
Dim i As Integer
Dim matv() As Double
Dim mata() As String

matport(1) = "DERIV SECT FINANCIERO"
matport(2) = "DERIV SECT NO FINANCIERO"
matport(3) = "DERIVADOS DE COBERTURA"
matport(4) = "DERIVADOS DE NEGOCIACION"
matport(5) = "DERIVADOS ESTRUCTURALES"
matport(6) = "DERIVADOS NEGOCIACION RECLASIFICACION"

Dim matvalor1(1 To 6) As Double
Dim matconcepto(1 To 17) As String
matconcepto(1) = "38"
matconcepto(2) = "57"
matconcepto(3) = "58"
matconcepto(4) = "59"
matconcepto(5) = "60"
matconcepto(6) = "62"
matconcepto(7) = "64"
matconcepto(8) = "65"
matconcepto(9) = "66"
matconcepto(10) = "67"
matconcepto(11) = "68"
matconcepto(12) = "69"
matconcepto(13) = "71"
matconcepto(14) = "72"
matconcepto(15) = "73"
matconcepto(16) = "77"
matconcepto(17) = "78"
For i = 1 To 6
   matv = LeerResValPort(fecha, txtportCalc1, matport(i), id_val)
   If UBound(matv, 1) <> 0 Then
      matvalor1(i) = matv(1)
   Else
      matvalor1(i) = 0
   End If
Next i
Dim matvalor2(1 To 17) As Double
For i = 1 To 17
   matv = LeerResValOper(fecha, txtportCalc1, 4, matconcepto(i), id_val)
   matvalor2(i) = matv(1)
Next i
Dim mats(1 To 31) As Double
For i = 1 To 6
    mats(i) = matvalor1(i)
Next i
For i = 1 To 17
    mats(i + 6) = matvalor2(i)
Next i
For i = 1 To 6
    mats(i + 23) = contarOperDeriv(fecha, matport(i))
Next i
   mata = ObtenerContrapFinDer(fecha)
   mats(30) = UBound(mata, 1)
   mata = ObtenerContrapNoFinDer(fecha)
   mats(31) = UBound(mata, 1)
LeerValDerivadosExcel = mats
End Function

Function contarOperDeriv(ByVal fecha As Date, ByVal txtport As String)
Dim txtfiltro As String
Dim txtfecha As String
Dim rmesa As New ADODB.recordset
txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT COUNT(*) FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
txtfiltro = txtfiltro & " AND PORTAFOLIO ='" & txtport & "'"
rmesa.Open txtfiltro, ConAdo
  contarOperDeriv = rmesa.Fields(0)
rmesa.Close
End Function

Sub GuardaValDerivadosExcel(ByVal fecha As Date, ByRef mata() As Double, ByVal nomtabla As String, conex, rbase)
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim j As Integer

noreg = UBound(mata, 1)
If noreg <> 0 Then
   txtfiltro = "SELECT COUNT(*) FROM [" & nomtabla & "] WHERE FECHA = " & CLng(fecha)
   rbase.Open txtfiltro, conex
   noreg1 = rbase.Fields(0)
   rbase.Close
   If noreg1 <> 0 Then
       
   Else
      txtcadena = "INSERT INTO [" & nomtabla & "] VALUES("
      txtcadena = txtcadena & CLng(fecha) & ","
      For j = 1 To noreg - 1
          txtcadena = txtcadena & Val(mata(j)) & ","
      Next j
      txtcadena = txtcadena & Val(mata(noreg)) & ")"
      conex.Execute txtcadena
   End If
End If
End Sub

Function LeerDuracionPos(ByVal fecha As Date) As Variant()
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim contar As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
 txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfiltro2 = "SELECT * from " & TablaValPos & " WHERE FECHAP = " & txtfecha
 txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtportCalc1 & "'"
 txtfiltro2 = txtfiltro2 & " AND ESC_FR = 'Normal'"
 txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = 1"
 txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
 rmesa.Open txtfiltro1, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
 If noreg <> 0 Then
    rmesa.Open txtfiltro2, ConAdo
    ReDim matres(1 To noreg, 1 To 5) As Variant
    contar = 0
    For i = 1 To noreg
        matres(i, 1) = rmesa.Fields("CPOSICION")          'clave de posicion
        matres(i, 2) = rmesa.Fields("COPERACION")         'clave de operacion
        matres(i, 3) = rmesa.Fields("T_OPERACION")        'tipo operacion
        matres(i, 4) = rmesa.Fields("DUR_ACT")            'duracion activa
        matres(i, 5) = rmesa.Fields("DUR_PAS")            'duracion pasiva
        rmesa.MoveNext
        MensajeProc = "Leyendo las duraciones de la posicion del dia " & fecha
    Next i
rmesa.Close

Else
ReDim matres(0 To 0, 0 To 0) As Variant
End If
LeerDuracionPos = matres
End Function

Function LeerResValOper(ByVal fecha As Date, ByVal txtport As String, ByVal id_posicion As Long, ByVal id_operacion As String, ByVal tval As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Long
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * from " & TablaValPos & " WHERE FECHAP = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FECHAFR = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FECHAV = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & id_posicion
txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & id_operacion & "'"
txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = " & tval
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
Dim mata(1 To 3) As Double
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(1) = rmesa.Fields(11)      'MTM
   mata(2) = rmesa.Fields(12)      'VAL ACTIVA
   mata(3) = rmesa.Fields(13)      'VAL PASIVA
   rmesa.Close
End If
LeerResValOper = mata
End Function

Function LeerResValDeriv(ByVal fecha As Date, ByVal txtport As String, ByVal id_val As Integer, ByRef contar As Long) As Variant()
'objetivo de la funcion: leer los resultados del calculo de la valuacion de la posicion
'de derivados y colocarlos en un array para su impresion en un archivo de texto

'datos de entrada:
'fecha    -   fecha del proceso
'txtport  -   nombre del portafolio sobre el que se realizaron los calculos
'id_val   -   clave de valuacion de la posicion
'

Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset
txtfiltro2 = CrearCadSQLValPosDeriv(fecha, txtport, ClavePosDeriv, id_val)
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mata(1 To noreg, 1 To 17) As Variant
   rmesa.Open txtfiltro2, ConAdo
   contar = 0
   For i = 1 To noreg
        mata(i, 1) = rmesa.Fields("COPERACION")     'clave de operacion
        mata(i, 2) = rmesa.Fields("VAL_ACT_S")      'VAL ACTIVA SIVARMER
        mata(i, 3) = rmesa.Fields("VAL_PAS_S")      'VAL PASIVA SIVARMER
        mata(i, 4) = rmesa.Fields("MTM_S")          'mtm sivarmer
        mata(i, 5) = rmesa.Fields("VAL_ACT_IKOS")   'val activa ikos derivados
        mata(i, 6) = rmesa.Fields("VAL_PAS_IKOS")   'val pasiva ikos derivados
        mata(i, 7) = rmesa.Fields("MTM_IKOS")       'mtm ikos derivados
        mata(i, 8) = mata(i, 5) - mata(i, 2)        'dif pata activa
        mata(i, 9) = mata(i, 6) - mata(i, 3)        'dif pata pasiva
        mata(i, 10) = mata(i, 7) - mata(i, 4)       'dif mtm
        If mata(i, 4) <> 0 Then mata(i, 11) = mata(i, 10) / mata(i, 4) 'diferencia porcentual
        mata(i, 12) = rmesa.Fields("FINICIO")       'FECHA DE INICIO
        mata(i, 13) = rmesa.Fields("FVENCIMIENTO")  'FECHA DE VENCIMIENTO
        mata(i, 14) = rmesa.Fields("INTENCION")     'intencion
        mata(i, 15) = rmesa.Fields("CPRODUCTO")     'CLAVE DEL PRODUCTO
        mata(i, 16) = rmesa.Fields("FVALUACION")    'GRUPO DE PRODUCTO
        mata(i, 17) = rmesa.Fields("ID_CONTRAP")    'clave de contraparte
        If Abs(mata(i, 10)) > 100 Then
          contar = contar + 1
        End If
        rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If

LeerResValDeriv = mata
End Function


Function LeerResValPort(ByVal fecha As Date, ByVal txtport As String, ByVal txtsubport As String, ByVal tval As Integer) As Double()
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim contar As Integer
Dim noreg As Long
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * from " & TablaValPosPort & " WHERE FECHAP = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FECHAFR = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FECHAV = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & txtsubport & "'"
txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = '" & tval & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
ReDim mata(1 To 3) As Double
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   contar = 0
    mata(1) = rmesa.Fields(7)       'MTM
    mata(2) = rmesa.Fields(8)       'val activa
    mata(3) = rmesa.Fields(9)       'val pasiva
    rmesa.Close
Else
   ReDim mata(0 To 0) As Double
End If
LeerResValPort = mata
End Function

Function LeerResPosMD(ByVal fecha As Date, ByVal txtgrupoport As String) As Variant()
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim noport As Integer
Dim txtport As String
Dim valactiva As Double
Dim valpasiva As Double
Dim dxvactiva As Integer
Dim dxvpasiva As Integer
Dim dxvcactiva As Integer
Dim dxvcpasiva As Integer
Dim ntactiva As Double
Dim ntpasiva As Double
Dim duractiva As Double
Dim durpasiva As Double
Dim dv01activa As Double
Dim dv01pasiva As Double
Dim rmesa As New ADODB.recordset

noport = 15
txtport = txtportCalc1
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
ReDim mats(1 To noport) As Variant
'ReDim matc1(1 To 5) As Variant
'ReDim matc2(1 To 5) As Variant
'ReDim matc3(1 To 5) As Variant
'se definen las claves a buscar en la tabla de datos
mats(1) = "MDT DIRECTO BONDES D"
mats(2) = "MDT DIRECTO BONOS M"
mats(3) = "MDT DIRECTO BONOS USD"
mats(4) = "MDT DIRECTO CBICS"
mats(5) = "MDT DIRECTO CERTIFICADOS BURSATILES"
mats(6) = "MDT DIRECTO CETES"
mats(7) = "MDT DIRECTO IPAB IM"
mats(8) = "MDT DIRECTO IPAB IQ"
mats(9) = "MDT DIRECTO IPAB IS"
mats(10) = "MDT DIRECTO PRLV"
mats(11) = "MDT DIRECTO UDIBONOS"
mats(12) = "MDT REPORTOS COMPRA"
mats(13) = "MDT REPORTOS VENTA"
mats(14) = "MDT DIRECTO"
mats(15) = "MERCADO DE DINERO"
ReDim mata(1 To noport, 1 To 6) As Variant
For i = 1 To noport
For j = 1 To 6
mata(i, j) = 0
Next j
Next i
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
For i = 1 To noport
    txtfiltro2 = "SELECT * from " & TablaValPosPort & " WHERE FECHAP = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT = '" & mats(i) & "' AND ID_VALUACION = 1"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg
           valactiva = rmesa.Fields(8)        'valuacion activa sucia
           valpasiva = rmesa.Fields(9)        'valuacion pasiva sucia
           dxvactiva = rmesa.Fields(21)       'dxv activa
           dxvpasiva = rmesa.Fields(22)       'dxv pasiva
           If Abs(valactiva - valpasiva) > 0.0001 Then
              mata(i, 1) = (dxvactiva * valactiva - dxvpasiva * valpasiva) / (valactiva - valpasiva)
           Else
              mata(i, 1) = 0
           End If
           dxvcactiva = rmesa.Fields(19)      'dxv cupon activa
           dxvcpasiva = rmesa.Fields(20)      'dxv cupon activa
           If Abs(valactiva - valpasiva) > 0.0001 Then
              mata(i, 2) = (dxvcactiva * valactiva - dxvcpasiva * valpasiva) / (valactiva - valpasiva)
           Else
              mata(i, 2) = 0
           End If
           ntactiva = rmesa.Fields(13)
           ntpasiva = rmesa.Fields(14)
           mata(i, 3) = ntactiva - ntpasiva  'total de titulos
           mata(i, 4) = rmesa.Fields(7)       'marca a mercado sucio
           duractiva = rmesa.Fields(15)
           durpasiva = rmesa.Fields(16)
           If Abs(valactiva - valpasiva) > 0.0001 Then
              mata(i, 5) = (duractiva * valactiva - durpasiva * valpasiva) / (valactiva - valpasiva)
           Else
              mata(i, 5) = 0
           End If
           dv01activa = rmesa.Fields(17)
           dv01pasiva = rmesa.Fields(18)
           mata(i, 6) = dv01activa - dv01pasiva
           rmesa.MoveNext
       Next j
       rmesa.Close
     End If
   Next i
LeerResPosMD = mata
End Function

Function LeerResPosInversion(ByVal fecha As Date) As Variant()
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Integer
Dim j As Integer
Dim noport As Integer
Dim txtport As String
Dim valactiva As Double
Dim valpasiva As Double
Dim dxvactiva As Integer
Dim dxvpasiva As Integer
Dim dxvcactiva As Integer
Dim dxvcpasiva As Integer
Dim ntactiva As Double
Dim ntpasiva As Double
Dim duractiva As Double
Dim durpasiva As Double
Dim dv01activa As Double
Dim dv01pasiva As Double
Dim rmesa As New ADODB.recordset

noport = 8
txtport = txtportCalc1
ReDim matsubport(1 To noport) As Variant
'ReDim matc1(1 To 5) As Variant
'ReDim matc2(1 To 5) As Variant
'ReDim matc3(1 To 5) As Variant
'se definen las claves a buscar en la tabla de datos
 matsubport(1) = "PICV DIR CB 90"
 matsubport(2) = "PICV DIR CB 91"
 matsubport(3) = "PICV DIR CB 92"
 matsubport(4) = "PICV DIR CB 93"
 matsubport(5) = "PICV DIR CB 94"
 matsubport(6) = "PICV DIR CB 95"
 matsubport(7) = "PICV DIR F"
 matsubport(8) = "PI CONSERVADOS A VENCIMIENTO"
ReDim mata(1 To noport, 1 To 6) As Variant
For i = 1 To noport
For j = 1 To 6
mata(i, j) = 0
Next j
Next i
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
For i = 1 To noport
    txtfiltro2 = "SELECT * from " & TablaValPosPort & " WHERE FECHAP = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT = '" & matsubport(i) & "' AND ID_VALUACION = 1"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       For j = 1 To noreg
           valactiva = rmesa.Fields(8)        'valuacion activa sucia
           valpasiva = rmesa.Fields(9)        'valuacion pasiva sucia
           dxvactiva = rmesa.Fields(21)       'dxv activa
           dxvpasiva = rmesa.Fields(22)       'dxv pasiva
           If Abs(valactiva - valpasiva) > 0.0001 Then
              mata(i, 1) = (dxvactiva * valactiva - dxvpasiva * valpasiva) / (valactiva - valpasiva)
           Else
              mata(i, 1) = 0
           End If
           dxvcactiva = rmesa.Fields(19)      'dxv cupon activa
           dxvcpasiva = rmesa.Fields(20)      'dxv cupon activa
           If Abs(valactiva - valpasiva) > 0.0001 Then
              mata(i, 2) = (dxvcactiva * valactiva - dxvcpasiva * valpasiva) / (valactiva - valpasiva)
           Else
              mata(i, 2) = 0
           End If
           ntactiva = rmesa.Fields(13)
           ntpasiva = rmesa.Fields(14)
           mata(i, 3) = ntactiva - ntpasiva  'total de titulos
           mata(i, 4) = rmesa.Fields(7)       'marca a mercado sucio
           duractiva = rmesa.Fields(15)
           durpasiva = rmesa.Fields(16)
           If Abs(valactiva - valpasiva) > 0.0001 Then
              mata(i, 5) = (duractiva * valactiva - durpasiva * valpasiva) / (valactiva - valpasiva)
           Else
              mata(i, 5) = 0
           End If
           dv01activa = rmesa.Fields(17)
           dv01pasiva = rmesa.Fields(18)
           mata(i, 6) = dv01activa - dv01pasiva
           rmesa.MoveNext
       Next j
       rmesa.Close
     End If
   Next i
LeerResPosInversion = mata
End Function


Function CProbIncUMS(ByVal fecha As Date, ByVal fvenums33 As Date, ByVal fvenust10 As Date, ByVal fvenust30 As Date, ByVal tcums As Double, ByVal rums As Double, ByVal rust10 As Double, ByVal rust30 As Double) As Variant()
Dim pc As Long
Dim dxvums33 As Long
Dim dxvust10 As Long
Dim noreg As Long
Dim i As Long
'rutina para calcular la probabilidad de incumplimiento del bono ums33

'rinterpol=

pc = 180
dxvums33 = -Int(-(fvenums33 - fecha) / pc) / 2
dxvust10 = -Int(-(fvenust10 - fecha) / pc) / 2
noreg = dxvums33
ReDim mata(1 To noreg, 1 To 20) As Variant
For i = 1 To noreg
mata(i, 1) = fecha          'fecha de corte
mata(i, 2) = noreg - i      'semestres por vencer
mata(i, 3) = 100 * tcums * 180 / 360 'cupon
mata(i, 4) = mata(i, 3) * 1 / (1 + rums * 180 / 360) ^ mata(i, 2)
mata(i, 5) = mata(i, 3) * 1 / (1 + rums * 180 / 360) ^ mata(i, 2)
Next i

CProbIncUMS = mata
End Function

Sub GuardaCurvasTotal2(ByVal fecha As Date, ByVal direc As String, ByVal nr As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim noreg As Integer
Dim nomarch As String
Dim sihayarch As Boolean
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim txtfecha As String
Dim txtborra As String
Dim contar As Integer
Dim txtcadena As String
Dim indice As Long
Dim largo As Long
Dim tamanoseg As Long
Dim nosegmentos As Long
Dim residuo As Long
Dim txttexto As String
Dim valor As Variant
Dim rmesa As New ADODB.recordset

 nomarch = direc & "\" & "CURVAS" & Format(fecha, "yyyymmdd") & ".XLS"
 sihayarch = VerifAccesoArch(nomarch)
 If sihayarch Then

    Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
    Set registros1 = base1.OpenRecordset("Sheet1$", dbOpenDynaset, dbReadOnly)
    registros1.MoveLast
    noreg = registros1.RecordCount
    nocampos = registros1.Fields.Count
    If noreg <> 12000 Then
       registros1.Close
       base1.Close
       txtmsg = "El archivo de curvas " & Format(fecha, "dd-mm-yyyy") & " no tiene los 12,000 registros requeridos"
       exito = False
       Exit Sub
    End If
    ReDim mata(1 To noreg, 1 To nocampos) As Double
    registros1.MoveFirst
    For i = 1 To noreg
        For j = 1 To nocampos
            valor = LeerTAccess(registros1, j - 1, i)
            If IsNumeric(valor) Then
               mata(i, j) = valor
            Else
              registros1.Close
              base1.Close
              mata(i, j) = 0
              txtmsg = "Hay datos no validos en el archivo"
              exito = False
              Exit Sub
            End If
        Next j
        registros1.MoveNext
        AvanceProc = i / noreg
        MensajeProc = "Leyendo el archivo de curvas del " & fecha & " " & Format(AvanceProc, "#00.00 %")
        DoEvents
    Next i
    registros1.Close
    base1.Close
 'se procede a guardar toda la informacion de las curvas en una tabla de datos
    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtborra = "DELETE FROM " & TablaCurvas & " WHERE FECHA = " & txtfecha
    ConAdo.Execute txtborra
    rmesa.Open "select * from " & TablaCurvas, ConAdo, 1, 3
    contar = 0
    
    For i = 1 To UBound(MatCatCurvas, 1)
        indice = DetColCurva(MatCatCurvas(i, 1), fecha)
        If indice <> 0 Then
           txtcadena = ""
           For j = 1 To noreg - 1
               txtcadena = txtcadena & mata(j, indice + 1) & ","
           Next j
           txtcadena = txtcadena & mata(noreg, indice + 1)
           rmesa.AddNew
           rmesa.Fields(0) = CLng(fecha)                 'fecha del archivo
           rmesa.Fields(1) = MatCatCurvas(i, 1)          'clave de la curva
           Call GuardarElementoClob(txtcadena, rmesa, "CURVA")
           rmesa.Fields(3) = NomUsuario                 'usuario
           rmesa.Update
           contar = contar + 1
        Else
           MensajeProc = "No hay datos para la curva"
        End If
        AvanceProc = i / UBound(MatCatCurvas, 1)
        MensajeProc = "Guardando las curvas del " & fecha & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
    Next i
    nr = contar
    rmesa.Close
    txtmsg = "El proceso finalizo correctamente"
    exito = True
 Else
    MensajeProc = "No hay acceso al archivo " & nomarch
    txtmsg = MensajeProc
    exito = False
 End If
End Sub

Sub GuardaCurvaInd(ByVal fecha As Date, ByVal direc As String, txtcurva As String, ByVal nr As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim noreg As Integer
Dim nomarch As String
Dim sihayarch As Boolean
Dim registros1 As DAO.recordset
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim txtfecha As String
Dim txtborra As String
Dim contar As Integer
Dim txtcadena As String
Dim indice As Long
Dim largo As Long
Dim tamanoseg As Long
Dim nosegmentos As Long
Dim residuo As Long
Dim txttexto As String
Dim valor As Variant
Dim rmesa As New ADODB.recordset

 nomarch = direc & "\" & "CURVAS" & Format(fecha, "yyyymmdd") & ".XLS"
 sihayarch = VerifAccesoArch(nomarch)
 If sihayarch Then
    Dim base1 As DAO.Database
    Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
    Set registros1 = base1.OpenRecordset("Sheet1$", dbOpenDynaset, dbReadOnly)
    registros1.MoveLast
    noreg = registros1.RecordCount
    nocampos = registros1.Fields.Count
    If noreg <> 12000 Then
       registros1.Close
       base1.Close
       txtmsg = "El archivo de curvas " & Format(fecha, "dd-mm-yyyy") & " no tiene los 12,000 registros requeridos"
       exito = False
       Exit Sub
    End If
    ReDim mata(1 To noreg, 1 To nocampos) As Double
    registros1.MoveFirst
    For i = 1 To noreg
        For j = 1 To nocampos
            valor = LeerTAccess(registros1, j - 1, i)
            If IsNumeric(valor) Then
               mata(i, j) = LeerTAccess(registros1, j - 1, i)
            Else
              registros1.Close
              base1.Close
              mata(i, j) = 0
              txtmsg = "Hay datos incorrectos en el archivo de curvas"
              exito = False
              Exit Sub
            End If
        Next j
        registros1.MoveNext
        AvanceProc = i / noreg
        MensajeProc = "Leyendo el archivo de curvas del " & fecha & " " & Format(AvanceProc, "#00.00 %")
        DoEvents
    Next i
    registros1.Close
    base1.Close
 'se procede a guardar toda la informacion de las curvas en una tabla de datos
    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"

    contar = 0
    For i = 1 To UBound(MatCatCurvas, 1)
        If MatCatCurvas(i, 2) = txtcurva Then
           indice = DetColCurva(MatCatCurvas(i, 1), fecha)
           If indice <> 0 Then
              txtborra = "DELETE FROM " & TablaCurvas & " WHERE FECHA = " & txtfecha & " AND IDCURVA= " & MatCatCurvas(i, 1)
              ConAdo.Execute txtborra
              rmesa.Open "select * from " & TablaCurvas, ConAdo, 1, 3
              txtcadena = ""
              For j = 1 To noreg - 1
                  txtcadena = txtcadena & mata(j, indice + 1) & ","
              Next j
              txtcadena = txtcadena & mata(noreg, indice + 1)
              rmesa.AddNew
              rmesa.Fields(0) = CLng(fecha)                 'fecha del archivo
              rmesa.Fields(1) = MatCatCurvas(i, 1)          'clave de la curva
              Call GuardarElementoClob(txtcadena, rmesa, "CURVA")
              rmesa.Fields(3) = NomUsuario                 'usuario
              rmesa.Update
              rmesa.Close
              contar = contar + 1
           Else
              MensajeProc = "No hay datos para la curva"
           End If
           Exit For
        End If
        AvanceProc = i / UBound(MatCatCurvas, 1)
        MensajeProc = "Guardando las curvas del " & fecha & " " & Format(AvanceProc, "##0.00 %")
        DoEvents
    Next i
    nr = contar
    txtmsg = "El proceso finalizo correctamente"
    exito = True
 Else
    MensajeProc = "No hay acceso al archivo " & nomarch
    txtmsg = MensajeProc
    exito = False
 End If
End Sub


Function DetColCurva(ByVal idcurva As Integer, ByVal fecha As Date)
Dim noreg As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Integer
Dim RInterfIKOS As New ADODB.recordset

txtfiltro1 = "SELECT * FROM " & PrefijoBD & TablaHistCurvas & " WHERE ID_CURVA = " & idcurva & " ORDER BY FECHA"
txtfiltro2 = "SELECT COUNT(*) FROM (" & txtfiltro1 & ")"
RInterfIKOS.Open txtfiltro2, ConAdo
noreg = RInterfIKOS.Fields(0)
RInterfIKOS.Close
If noreg <> 0 Then
   RInterfIKOS.Open txtfiltro1, ConAdo
   ReDim mata(1 To noreg, 1 To 2) As Variant
   For i = 1 To noreg
       mata(i, 1) = RInterfIKOS.Fields(2)   'fecha
       mata(i, 2) = RInterfIKOS.Fields(3)   'columna
       RInterfIKOS.MoveNext
   Next i
   RInterfIKOS.Close
   If fecha < mata(1, 1) Then
      DetColCurva = 0
      Exit Function
   ElseIf fecha >= mata(noreg, 1) Then
      DetColCurva = mata(noreg, 2)
      Exit Function
   Else
      For i = 1 To noreg - 1
          If fecha >= mata(i, 1) And fecha < mata(i + 1, 1) Then
             DetColCurva = mata(i, 2)
             Exit Function
          End If
      Next i
   End If
Else
   DetColCurva = 0
End If
End Function

Function LeerCurvaCompleta(ByVal fecha As Date, ByRef exito As Boolean) As Variant()
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim matc() As String
Dim rmesa As New ADODB.recordset

'rutina que lee todo el archivo de curvas
'1 fecha
'2 plazo
'3 contenido del archivo de curvas

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro1 = "SELECT * from " & TablaCurvas & " WHERE FECHA = " & txtfecha & " ORDER BY IDCURVA"
txtfiltro2 = "SELECT COUNT(*) from (" & txtfiltro1 & ")"
rmesa.Open txtfiltro2, ConAdo
 noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mata(1 To 12001, 1 To noreg) As Variant
   ReDim matb(1 To 3) As Variant
   rmesa.Open txtfiltro1, ConAdo
   For i = 1 To noreg
       matb(1) = rmesa.Fields(1) 'id curva
       matb(2) = rmesa.Fields(2).GetChunk(rmesa.Fields(2).ActualSize)
       matc = EncontrarSubCadenas(matb(2), ",")
       mata(1, i) = matb(1)
       For j = 1 To 12000
           If j <= UBound(matc, 1) Then
              mata(j + 1, i) = CDbl(matc(j))
           Else
              mata(j + 1, i) = 0
           End If
       Next j
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo las curvas del " & fecha & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
 rmesa.Close
 exito = True
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
   exito = False
End If

LeerCurvaCompleta = mata
End Function

Function LeerNodosCurvaO(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtcurva As String, ByRef matx() As Variant, ByRef exito As Boolean) As Variant()
Dim idcurva As Long
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim nocampos As Long
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim matc() As String
Dim rmesa As New ADODB.recordset

'rutina que lee la historia de una curva entre la fecha1 y fecha2 en funcion de una lista de nodos
'1 fecha
'2 plazo
'3 contenido del archivo de curvas
idcurva = 0
For i = 1 To UBound(MatCatCurvas, 1)
    If MatCatCurvas(i, 2) = txtcurva Then
       idcurva = MatCatCurvas(i, 1)
       Exit For
    End If
Next i
If idcurva <> 0 Then
   txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro1 = "SELECT * from " & TablaCurvas & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & " AND IDCURVA = " & idcurva & " ORDER BY FECHA"
   txtfiltro2 = "SELECT COUNT(*) from (" & txtfiltro1 & ")"
   rmesa.Open txtfiltro2, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      nocampos = UBound(matx, 1)
      ReDim mata(1 To noreg, 1 To nocampos + 1) As Variant
      ReDim matb(1 To 2) As Variant
      rmesa.Open txtfiltro1, ConAdo
      For i = 1 To noreg
          matb(1) = rmesa.Fields(0)  'fecha
          matb(2) = rmesa.Fields(2).GetChunk(rmesa.Fields(2).ActualSize)   'curva
          matc = EncontrarSubCadenas(matb(2), ",")
          mata(i, 1) = matb(1)
          For j = 1 To nocampos
              If matx(j) <= UBound(matc, 1) Then
                 mata(i, j + 1) = CDbl(matc(matx(j)))
              Else
                 mata(i, j + 1) = 0
              End If
          Next j
          rmesa.MoveNext
          AvanceProc = i / noreg
          MensajeProc = "Leyendo la curva " & txtcurva & " del " & mata(i, 1) & " " & Format(AvanceProc, "##0.00 %")
          DoEvents
      Next i
      rmesa.Close
   End If
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerNodosCurvaO = mata
End Function

Sub ImpPosPrimSwapsArch(ByVal fecha As Date, ByVal nomarch As String, ByRef nr As Long, ByRef exito As Boolean)
Dim matpos2() As propPosDeuda
Dim matfl2() As Variant
Dim sihayarch As Boolean
Dim exito2 As Boolean
Dim i As Integer
'la matriz de caracteristicas
'aun asi se tiene que leer las valuaciones del dia para adjuntarlas en la tabla de datos
sihayarch = VerifAccesoArch(nomarch)
 'la matriz de flujos
If sihayarch Then
'se lee la fuente de datos para riesgos
'se lee la tabla de datos de oracle u excel  no aplica ningun filtro ya que espera que todos los
'registros sean del mismo grupo
  matpos2 = LeerCaractDeudaExcel(nomarch)
  matfl2 = LeerFlujosDExcel(nomarch, exito2)
  matfl2 = RutinaOrden(matfl2, UBound(matfl2, 2), SRutOrden)
  If UBound(matpos2, 1) <> 0 Then
    For i = 1 To UBound(matpos2, 1)
        matpos2(i).fechareg = fecha
    Next i
     Call GuardarPosDeuda(1, "Real", "000000", matpos2, matfl2)
    ' Call ActualizarPosDeuda(matpos2, matfl2, 1, "Real", nr, exito)
  Else
     MsgBox "hay errores en la tabla de datos"
  End If
End If
End Sub

Sub DetermFValDeuda(ByVal coperacion As String, ByRef fvalua As String, ByVal toper As Integer)
Dim coper2 As String
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim j As Integer
Dim rmesa As New ADODB.recordset
Dim matoper() As Variant
    coper2 = DetCoper(coperacion)
    
          txtfiltro2 = "SELECT * FROM " & TablaPosSwaps & " WHERE COPERACION = '" & coper2 & "' AND TIPOPOS = 1 ORDER BY FECHAREG"
          txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
          rmesa.Open txtfiltro1, ConAdo
          noreg = rmesa.Fields(0)
          rmesa.Close
          If noreg <> 0 Then
             rmesa.Open txtfiltro2, ConAdo
              ReDim matoper(1 To noreg, 1 To 2) As Variant
              For j = 1 To noreg
                  matoper(j, 1) = rmesa.Fields("coperacion")
                  matoper(j, 2) = rmesa.Fields("fechareg")
                  rmesa.MoveNext
              Next j
              rmesa.Close
              For j = 1 To noreg
                 frmListaOpR.Combo1.AddItem matoper(j, 1) & " " & matoper(j, 2)
              Next j
              IndOperR = 0
              Do While IndOperR = 0
                 frmListaOpR.Show 1
              Loop
              fvalua = DetFValxSwapAsociado(1, matoper(IndOperR, 2), matoper(IndOperR, 1), "000000", toper)
          Else
             fvalua = ""
          End If
End Sub


Sub ImpPosSwapsArch(ByVal fecha As Date, ByVal nomarch As String, ByRef nr As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim matcar() As New propPosSwaps
Dim matfl() As Variant
Dim sihayarch As Boolean
Dim exito2 As Boolean
'la matriz de caracteristicas
'aun asi se tiene que leer las valuaciones del dia para adjuntarlas en la tabla de datos
sihayarch = VerifAccesoArch(nomarch)
 'la matriz de flujos
If sihayarch Then
'se lee la fuente de datos para riesgos
'se lee la tabla de datos de oracle u excel  no aplica ningun filtro ya que espera que todos los
'registros sean del mismo grupo
   matcar = LeerCaractSwapExcel(nomarch)
   matfl = LeerFlujosSwapExcel(nomarch, exito2)
   If UBound(matcar, 1) <> 0 And UBound(matfl, 1) <> 0 Then
      Call GuardarPosSwaps(fecha, 2, "Previa", "000000", matcar, matfl, exito2)
   End If
End If
End Sub

Function LeerFlujosPrimSwapsExcel(ByVal nomarch As String, ByRef exito As Boolean) As Variant()
 Dim base1 As DAO.Database
 Dim registros1 As DAO.recordset
 Dim i As Long
 Dim j As Long
 Dim noreg As Long
 Dim nocampos As Long
 
 Set base1 = OpenDatabase(nomarch, False, "S", VersExcel)
 Set registros1 = base1.OpenRecordset("Flujos$", dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
 If registros1.RecordCount <> 0 Then
    registros1.MoveLast
    noreg = registros1.RecordCount
    registros1.MoveFirst
    nocampos = registros1.Fields.Count
    exito = True
    ReDim matfl(1 To noreg, 1 To nocampos + 2) As Variant
    For i = 1 To noreg
       For j = 1 To nocampos
          matfl(i, j) = LeerTAccess(registros1, j - 1, i)
          If Len(Trim(matfl(i, j))) = 0 Then exito = False
       Next j
       matfl(i, nocampos + 1) = "S " & matfl(i, 1) & " P " & matfl(i, 2)           'clave de la pata a la que pertenece el flujo
       matfl(i, nocampos + 2) = matfl(i, nocampos + 1) & CLng(matfl(i, 4))         'clave de ordenacion del flujo
       registros1.MoveNext
       Call MostrarMensajeSistema("Leyendo los flujos de los swaps del archivo de excel " & i & "/" & noreg & " " & Format(AvanceProc, "###0.00"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
       AvanceProc = i / noreg
       DoEvents
    Next i
    registros1.Close
    base1.Close
    matfl = RutinaOrden(matfl, UBound(matfl, 2), SRutOrden)
 Else
    ReDim matfl(0 To 0, 0 To 0) As Variant
 End If
 LeerFlujosPrimSwapsExcel = matfl
End Function

Function LeerFlujosSwapExcel(ByVal nomarch As String, ByRef exito As Boolean) As Variant()
 Dim base1 As DAO.Database
 Dim registros1 As DAO.recordset
 Dim noreg As Long
 Dim nocampos As Long
 Dim i As Long
 Dim j As Long
 
 Set base1 = OpenDatabase(nomarch, False, True, VersExcel)
 Set registros1 = base1.OpenRecordset("Flujos$", dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
 If registros1.RecordCount <> 0 Then
    registros1.MoveLast
    noreg = registros1.RecordCount
    registros1.MoveFirst
    nocampos = registros1.Fields.Count
    exito = True
    ReDim matfl(1 To noreg, 1 To nocampos + 2) As Variant
    For i = 1 To noreg
       For j = 1 To nocampos
          matfl(i, j) = LeerTAccess(registros1, j - 1, i)
          If EsVariableVacia(matfl(i, j)) Then exito = False
       Next j
       matfl(i, nocampos + 1) = "S " & matfl(i, 1) & " P " & matfl(i, 2)           'clave de la pata a la que pertenece el flujo
       matfl(i, nocampos + 2) = matfl(i, nocampos + 1) & CLng(matfl(i, 4))         'clave de ordenacion del flujo
       registros1.MoveNext
       Call MostrarMensajeSistema("Leyendo los flujos de los swaps del archivo de excel " & i & "/" & noreg & " " & Format(AvanceProc, "###0.00"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
       AvanceProc = i / noreg
       DoEvents
    Next i
    registros1.Close
    base1.Close
    matfl = RutinaOrden(matfl, UBound(matfl, 2), SRutOrden)
 Else
    ReDim matfl(0 To 0, 0 To 0) As Variant
 End If
 LeerFlujosSwapExcel = matfl
End Function

Function LeerFlujosDExcel(ByVal nomarch As String, ByRef exito As Boolean) As Variant()
 Dim base1 As DAO.Database
 Dim registros1 As DAO.recordset
 Dim noreg As Long
 Dim nocampos As Integer
 Dim i As Long
 Dim j As Integer
 
 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, True, VersExcel)
 Set registros1 = base1.OpenRecordset("Flujos$", dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
 If registros1.RecordCount <> 0 Then
    registros1.MoveLast
    noreg = registros1.RecordCount
    registros1.MoveFirst
    nocampos = registros1.Fields.Count
    exito = True
    ReDim matfl(1 To noreg, 1 To nocampos + 2) As Variant
    For i = 1 To noreg
       For j = 1 To nocampos
           matfl(i, j) = LeerTAccess(registros1, j - 1, i)
           If Len(Trim(matfl(i, j))) = 0 Then exito = False
       Next j
       matfl(i, nocampos + 1) = "S " & matfl(i, 1)                               'clave de la pata a la que pertenece el flujo
       matfl(i, nocampos + 2) = matfl(i, nocampos + 1) & " " & CLng(matfl(i, 2)) 'clave de ordenacion del flujo
       registros1.MoveNext
       Call MostrarMensajeSistema("Leyendo los flujos de los swaps del archivo de excel " & i & "/" & noreg & " " & Format(AvanceProc, "###0.00"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
       AvanceProc = i / noreg
       DoEvents
    Next i
    registros1.Close
    base1.Close
    matfl = RutinaOrden(matfl, UBound(matfl, 2), SRutOrden)
 Else
    ReDim matfl(0 To 0, 0 To 0) As Variant
 End If
 LeerFlujosDExcel = matfl
End Function

Function LeerCarPrimSwapsExcel(ByVal nomarch As String) As Variant()
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim nocampos As Long

 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
 Set registros1 = base1.OpenRecordset("Caract$", dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
 If registros1.RecordCount <> 0 Then
 registros1.MoveLast
 noreg = registros1.RecordCount
 registros1.MoveFirst
 nocampos = registros1.Fields.Count
 ReDim matpos(1 To noreg, 1 To nocampos) As Variant
 For i = 1 To noreg
  For j = 1 To nocampos
   matpos(i, j) = LeerTAccess(registros1, j - 1, i)
  Next j
 registros1.MoveNext
 AvanceProc = i / noreg
 DoEvents
 Next i
 registros1.Close
 base1.Close
 End If
 matpos = RutinaOrden(matpos, 1, SRutOrden)
LeerCarPrimSwapsExcel = matpos
End Function

Function LeerCaractSwapExcel(ByVal nomarch As String)
On Error GoTo hayerror
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim i As Long
Dim nocampos As Integer
Dim toper As String

 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
 Set registros1 = base1.OpenRecordset("Caract$", dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
 If registros1.RecordCount <> 0 Then
 registros1.MoveLast
 noreg = registros1.RecordCount
 registros1.MoveFirst
 nocampos = registros1.Fields.Count
 ReDim matpos(1 To noreg) As New propPosSwaps
 For i = 1 To noreg
     matpos(i).c_operacion = LeerTAccess(registros1, 0, i)
     matpos(i).C_Posicion = LeerTAccess(registros1, 1, i)
     toper = LeerTAccess(registros1, 2, i)
     If toper = "B" Or toper = "A" Then
        matpos(i).Tipo_Mov = 1
     ElseIf toper = "C" Or toper = "P" Then
        matpos(i).Tipo_Mov = 4
     End If
     
     matpos(i).intencion = LeerTAccess(registros1, 3, i)
     matpos(i).EstructuralSwap = LeerTAccess(registros1, 4, i)
     matpos(i).FCompraSwap = LeerTAccess(registros1, 5, i)
     matpos(i).FvencSwap = LeerTAccess(registros1, 6, i)
     matpos(i).IntercIFSwap = LeerTAccess(registros1, 7, i)
     matpos(i).IntercFFSwap = LeerTAccess(registros1, 8, i)
     matpos(i).RIntAct = LeerTAccess(registros1, 9, i)
     matpos(i).RIntPas = LeerTAccess(registros1, 10, i)
     matpos(i).TCActivaSwap = LeerTAccess(registros1, 11, i)
     matpos(i).TCPasivaSwap = LeerTAccess(registros1, 12, i)
     matpos(i).STActiva = LeerTAccess(registros1, 13, i)
     matpos(i).STPasiva = LeerTAccess(registros1, 14, i)
     matpos(i).ConvIntAct = LeerTAccess(registros1, 15, i)
     matpos(i).ConvIntPas = LeerTAccess(registros1, 16, i)
     matpos(i).cProdSwapGen = LeerTAccess(registros1, 17, i)
     matpos(i).ID_ContrapSwap = LeerTAccess(registros1, 18, i)
     registros1.MoveNext
     AvanceProc = i / noreg
     DoEvents
 Next i
 registros1.Close
 base1.Close
 End If
LeerCaractSwapExcel = matpos
Exit Function
hayerror:
MsgBox error(Err())
End Function

Function LeerCaractDeudaExcel(ByVal nomarch As String)
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim i As Long
Dim toper As String
Dim fvalua As String

 Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
 Set registros1 = base1.OpenRecordset("Caract$", dbOpenDynaset, dbReadOnly)
 'se revisa si hay registros en la tabla
 If registros1.RecordCount <> 0 Then
 registros1.MoveLast
 noreg = registros1.RecordCount
 registros1.MoveFirst
 ReDim matpos(1 To noreg) As New propPosDeuda
 For i = 1 To noreg
     matpos(i).C_Posicion = LeerTAccess(registros1, 0, i)
     matpos(i).c_operacion = LeerTAccess(registros1, 1, i)
     toper = LeerTAccess(registros1, 2, i)
     If toper = "A" Then
        matpos(i).Tipo_Mov = 1
     ElseIf toper = "P" Then
        matpos(i).Tipo_Mov = 4
     End If
     matpos(i).FinicioDeuda = LeerTAccess(registros1, 3, i)
     matpos(i).FVencDeuda = LeerTAccess(registros1, 4, i)
     matpos(i).InteriDeuda = LeerTAccess(registros1, 5, i)
     matpos(i).InterfDeuda = LeerTAccess(registros1, 6, i)
     matpos(i).RintDeuda = LeerTAccess(registros1, 7, i)
     matpos(i).TRefDeuda = LeerTAccess(registros1, 8, i)
     matpos(i).SpreadDeuda = LeerTAccess(registros1, 9, i)
     matpos(i).ConvIntDeuda = LeerTAccess(registros1, 10, i)
     'matpos(i).FValuadeuda = LeerTAccess(registros1, 16, i)
     'matpos(i).ID_ContrapSwap = LeerTAccess(registros1, 17, i)
     Call DetermFValDeuda(matpos(i).c_operacion, fvalua, matpos(i).Tipo_Mov)
     If Not EsVariableVacia(fvalua) Then
        matpos(i).fValuacion = fvalua
     Else
       MsgBox "no se definio con que operación se debe asociar"
     End If
     registros1.MoveNext
     AvanceProc = i / noreg
     DoEvents
 Next i
 registros1.Close
 base1.Close
 End If
LeerCaractDeudaExcel = matpos
End Function



Sub LeerPosPrimSwapsArch(ByVal fecha As Date, ByVal nomarch As String, ByRef noreg As Integer, ByRef exito As Boolean)
Dim matpos1() As propPosDeuda
Dim matfl1() As Variant
Dim matpos2() As Variant
Dim matfl2() As Variant
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim siesfv As Boolean
Dim exito2 As Boolean

txtfiltro1 = TablaPosDeuda & " order by COPERACION"
txtfiltro2 = TablaFlujosDeudaO
matpos1 = LeerTablaDeuda(txtfiltro1)
matfl1 = LeerFlujosPrimSwaps(txtfiltro2)
matpos2 = LeerCarPrimSwapsExcel(nomarch)
matfl2 = LeerFlujosPrimSwapsExcel(nomarch, exito2)
siesfv = EsFechaVaR(fecha)
If siesfv Then
   Call GenPosPrimSwaps(fecha, matpos1, matfl1, noreg, exito)
Else
   MsgBox "Falta la fecha en la tabla escenarios de VaR"
End If

End Sub

Sub GenPosHistO(ByVal fecha As Date, ByRef matfl() As Variant, ByRef matpos() As Variant, ByVal noreg As Long, ByRef exito As Boolean)
Dim matc() As Variant
Dim i As Long
Dim contar As Long
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtcadena As String
Dim contar11 As Long
Dim contar21 As Long
Dim nocampos1 As Long
Dim mats() As Variant
Dim txtfecha As String


'se filtran los swaps vigentes
mats = FiltrarSwapsVigentes(fecha, matpos)
 If UBound(mats, 1) > 0 Then
 matc = FiltrarFlujosSwaps(matfl, mats)
 nocampos1 = UBound(matc, 2)
 'se genera una clave que identifica a cada flujo de manera exclusiva
 '1  clave de operación ikos
 '2  posicion: banco ,cliente o posicion primaria
 '3  fecha inicio periodo
 '4  fecha final periodo
 '5  fecha pago intereses
 '6  saldo o monto nocional
 '7  amortizacion
 '8  tasa (var o fija)
 '9  spread
 '10 periodo cupon
 'se genera 2 claves que indentifica a las distintas patas de swaps
 'clave 1: no de swap, posicion
 'clave 1: no de swap, posicion, fecha de inicio de flujo
 For i = 1 To UBound(matc, 1)
  matc(i, nocampos1 - 1) = "S " & matc(i, 1) & "P " & matc(i, 2) 'clave para identificar la pata
 'clave para identificar el flujo
  matc(i, nocampos1) = "S " & matc(i, 1) & "P " & matc(i, 2) & " F " & Format(matc(i, 4), "#######")
 Next i
'se ordenan los campos de acuerdo a la segunda clave de ordenacion
matc = RutinaOrden(matc, nocampos1, SRutOrden)
'se borran registros previos
txtfecha = Format(fecha, "dd/mm/yyyy")
ConAdo.Execute "DELETE FROM " & TablaFlujosSwapsO & " WHERE FECHA = '" & txtfecha & "'"
contar = 0
For i = 1 To UBound(matc, 1)
'se crea el registro a insertar en la tabla de flujos
  txtfecha1 = Format(fecha, "dd/mm/yyyy")
  txtfecha2 = "to_date('" & Format(matc(i, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtfecha3 = "to_date('" & Format(matc(i, 5), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtcadena = "INSERT INTO " & TablaFlujosSwapsO & " VALUES("
  txtcadena = txtcadena & "'" & txtfecha1 & "',"                    'fecha o nombre de los flujos
  txtcadena = txtcadena & "'" & matc(i, 1) & "',"                   'Clave de operación en ikos
  txtcadena = txtcadena & "'" & matc(i, 2) & "',"                   'tipo de pata
  txtcadena = txtcadena & "'" & matc(i, 3) & "',"                   'intencion
  txtcadena = txtcadena & txtfecha2 & ","                           'fecha inicio flujo
  txtcadena = txtcadena & txtfecha3 & ","                           'fecha fin flujo
  txtcadena = txtcadena & "'" & matc(i, 6) & "',"                   'pago de intereses
  txtcadena = txtcadena & "'" & matc(i, 7) & "',"                   'reinversion intereses
  txtcadena = txtcadena & matc(i, 8) & ","                          'saldo
  txtcadena = txtcadena & matc(i, 9) & ","                          'amortizacion
  If Val(Trim(matc(i, 10))) <> 0 Then
  txtcadena = txtcadena & "'" & Val(Trim(matc(i, 10))) / 100 & "',"  'tasa cupon
  Else
  txtcadena = txtcadena & "'" & Trim(matc(i, 10)) & "',"             'tasa cupon
  End If
  txtcadena = txtcadena & Val(Trim(matc(i, 11))) / 100 & ")"         'spread
' MsgBox txtcadena
  ConAdo.Execute txtcadena
  contar = contar + 1
  AvanceProc = i / UBound(matc, 1)
  MensajeProc = "Guardando los flujos de los swaps del " & fecha & " " & i & "/" & UBound(matc, 1) & " " & Format(AvanceProc, "##0.00 %")
  DoEvents
Next i

If contar <> UBound(matc, 1) Then MsgBox "No se realizo el proceso en su totalidad"
'se borran los registros de swaps existentes
txtfecha = Format(fecha, "dd/mm/yyyy")
ConAdo.Execute "DELETE FROM " & TablaPosSwaps & " WHERE FECHA = '" & txtfecha & "'"
'AHORA SE CAPTURA LA referencia que nos indica la naturaleza del swap y algunos parametros
contar21 = 0
contar11 = 0
If UBound(mats, 1) <> 0 Then
noreg = UBound(mats, 1)
For i = 1 To noreg
  contar21 = contar21 + 1
  txtfecha1 = Format(fecha, "dd/mm/yyyy")
  txtfecha2 = "to_date('" & Format(mats(i, 3), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtfecha3 = "to_date('" & Format(mats(i, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
  txtcadena = "INSERT INTO " & TablaPosSwaps & " VALUES("
  txtcadena = txtcadena & "'" & txtfecha1 & "',"             'fecha de la posicion
  txtcadena = txtcadena & "'" & mats(i, 1) & "',"            'Clave de operación ikos
  txtcadena = txtcadena & "1,"                               'tipo de operacion
  txtcadena = txtcadena & "'" & mats(i, 2) & "',"            'intencion
  txtcadena = txtcadena & txtfecha2 & ","                    'fecha inicio swap
  txtcadena = txtcadena & txtfecha3 & ","                    'fecha venc swap
  txtcadena = txtcadena & "'" & mats(i, 5) & "',"            'interi
  txtcadena = txtcadena & "'" & mats(i, 6) & "',"            'interf
  txtcadena = txtcadena & "'" & mats(i, 7) & "',"            'clave de valuacion
  txtcadena = txtcadena & "'" & mats(i, 8) & "',"            'contraparte
  If mats(i, 9) Then
   contar11 = contar11 + 1
   txtcadena = txtcadena & "'55')"                           'derivados estrucurales
  Else
   txtcadena = txtcadena & "'4')"                            'derivados neg
  End If
  ConAdo.Execute txtcadena
  AvanceProc = i / noreg
  MensajeProc = "Guardando la posición de swaps del " & fecha & " " & Format(AvanceProc, "##0.00 %")
 DoEvents
Next i
MensajeProc = "Se guardaron " & contar21 & " registros de la posición de swaps del " & fecha
 If contar21 = noreg Then
     exito = True
     Call MostrarMensajeSistema("Se guardaron " & contar & " flujos de swaps del " & fecha, frmProgreso.Label2, 1, Date, Time, NomUsuario)
     Call MostrarMensajeSistema("Se guardaron " & contar21 & " registros de la posición de swaps del " & fecha, frmProgreso.Label2, 1, Date, Time, NomUsuario)
 Else
  MsgBox "No se encontro toda la información de las operaciones del dia"
 End If
Else
 MensajeProc = "Atencion. Faltan registros de la posicion de swaps para el " & fecha
 MsgBox MensajeProc
 Call MostrarMensajeSistema(MensajeProc, frmProgreso.Label2, 2, Date, Time, NomUsuario)
End If
End If
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub ActualizarPosSwaps(ByVal fecha As Date, ByRef matcar2() As propPosSwaps, tipopos, ByVal txtnompos As String, ByVal noreg As Integer, ByRef exito As Boolean)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim matcara() As New propPosSwaps
Dim matcarb() As New propPosSwaps
Dim matfla() As New estFlujosDeuda
Dim matflb() As Variant
Dim contar As Long
Dim i As Long
Dim c_operacion As String
Dim estaln As Boolean
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim iguales1 As Boolean
Dim iguales2 As Boolean

Dim exito1 As Boolean

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'matcar1 Y matflswap1 son los datos de riesgos
contar = 0
exito = False
For i = 1 To UBound(matcar2, 1)
    c_operacion = matcar2(i).c_operacion
    estaln = EnBlackList(c_operacion, MBList)
    If Not estaln Then
       matcarb = FiltrarCaracSwaps(matcar2, i)
       matflb = LeerFlujosSwapsIKOS(fecha, c_operacion, conAdoBD)      'los flujos de los swaps
       txtfecha1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfiltro1 = "SELECT * FROM " & TablaPosSwaps & " where TIPOPOS = " & tipopos
       txtfiltro1 = txtfiltro1 & " AND (TIPOPOS,COPERACION,FECHAREG) IN "
       txtfiltro1 = txtfiltro1 & "(SELECT TIPOPOS,COPERACION,MAX(FECHAREG) AS FECHAREG"
       txtfiltro1 = txtfiltro1 & " FROM " & TablaPosSwaps
       txtfiltro1 = txtfiltro1 & " WHERE COPERACION = '" & c_operacion & "'"
       txtfiltro1 = txtfiltro1 & " AND FECHAREG <=" & txtfecha1
       txtfiltro1 = txtfiltro1 & " AND TIPOPOS = " & tipopos
       txtfiltro1 = txtfiltro1 & " GROUP BY TIPOPOS,COPERACION)"
       matcara = LeerTablaSwaps(txtfiltro1)
       If UBound(matcara, 1) > 0 Then
          txtfecha2 = "to_date('" & Format(matcara(1).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtfiltro2 = "SELECT * FROM " & TablaFlujosSwapsO & " where TIPOPOS = " & tipopos
          txtfiltro2 = txtfiltro2 & " AND COPERACION = '" & c_operacion & "'"
          txtfiltro2 = txtfiltro2 & " AND FECHAREG = " & txtfecha2 & " ORDER BY TPATA,FINICIO"
          matfla = LeerFlujosSwaps(txtfiltro2, True)
          If fecha >= matcarb(1).FCompraSwap Then  'si la fecha de registro es mayor o igual que la fecha de inicio
             Call CompararFlujos(matfla, matflb, iguales1)
             Call CompararCSwaps(matcara, matcarb, iguales2)
             If Not iguales1 Or Not iguales2 Then
         'es una nueva version de la informacion, se agrega
                MensajeProc = "Se guarda una nueva version del swap " & c_operacion
                Call GuardarPosSwaps(fecha, tipopos, txtnompos, "000000", matcarb, matflb, exito1)
                contar = contar + 1
             End If
             MensajeProc = "Se comparo el swap " & c_operacion & " con registros previos"
          Else
             MensajeProc = "No se puede registrar una operacion antes de su fecha de inicio"
          End If
       Else
'no se encontro la informacion en la tabla de datos, se agrega la informacion
'la fecha de registro sera la fecha de inicio de la operacion
          Call GuardarPosSwaps(fecha, tipopos, txtnompos, "000000", matcarb, matflb, exito1)
          contar = contar + 1
          MensajeProc = "Guardando una operacion nueva"
       End If
       DoEvents
       MensajeProc = "Avance del proceso " & Format(i / UBound(matcar2, 1), "##0.00 %")
    End If
Next i
exito = True
noreg = contar
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function FiltrarPosDeuda1(ByRef matpos() As propPosDeuda, ByVal indice As Long)


ReDim matb(1 To 1) As propPosDeuda
 Set matb(1) = matpos(indice)

FiltrarPosDeuda1 = matb
End Function

Function FiltrarFlujosDeuda1(ByRef matfl() As Variant, ByVal copera As String) As Variant()
On Error GoTo hayerror

Dim noreg As Long
Dim nocampos As Long
Dim contar As Long
Dim i As Long
Dim j As Long


noreg = UBound(matfl, 1)
nocampos = UBound(matfl, 2)
ReDim matb(1 To nocampos, 1 To 1) As Variant
contar = 0
For i = 1 To noreg
    If matfl(i, 1) = copera Then
       contar = contar + 1
  ReDim Preserve matb(1 To nocampos, 1 To contar) As Variant
       For j = 1 To nocampos
           matb(j, contar) = matfl(i, j)
       Next j
    End If
Next i
FiltrarFlujosDeuda1 = MTranV(matb)

On Error GoTo 0
Exit Function
hayerror:
MsgBox "filtrarflujosdeuda1 " & error(Err())
End Function

Sub ActualizarPosDeuda(ByRef matpos2() As propPosDeuda, ByRef matfl2() As Variant, ByVal tipopos As Integer, ByVal txtnompos As String, ByRef noreg As Long, ByRef exito As Boolean)
Dim matpos1() As Variant
Dim matcar() As propPosDeuda
Dim matfl3() As Variant
Dim matfl1() As estFlujosDeuda
Dim iguales As Boolean
Dim sicomparar As Boolean
Dim contar As Long
Dim i As Long
Dim txtcoperacion As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim fregistro1 As Date

If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
'MATPOS1 Y matflswap1 ES LA BASE DE ORACLE
sicomparar = True
contar = 0
For i = 1 To UBound(matpos2, 1)
    txtcoperacion = "" & matpos2(i, 2) & ""
    matcar = FiltrarPosDeuda1(matpos2, i)
    matfl3 = FiltrarFlujosDeuda1(matfl2, txtcoperacion)
    txtfiltro1 = "SELECT * FROM " & TablaPosDeuda & " where TIPOPOS = " & tipopos & " and COPERACION = '" & txtcoperacion & "'"
    txtfiltro2 = "SELECT * FROM " & TablaFlujosDeudaO & " where TIPOPOS = " & tipopos & " AND COPERACION = '" & txtcoperacion & "'  ORDER BY FINICIO"
    matpos1 = LeerTablaDeuda(txtfiltro1)
    If UBound(matpos1, 1) > 0 Then
      If sicomparar Then
          matfl1 = LeerFlujosDeuda(txtfiltro2, True)
          If matpos1(i, 3) >= matpos2(i, 4) Then
     'se busca la fecha de registro adecuada para comparar con la nueva posicion
     'ahora hay que comparar las versiones correctas de informacion
             Call CompararFlujosD(matfl1, matfl3, iguales)
             'Call CompararCSwaps(matpos1, matpos2, iguales2)
             If Not iguales Then
          'es una nueva version de la informacion, se agrega
                MensajeProc = "Se guarda una nueva version del swap " & txtcoperacion
               Call GuardarPosDeuda(tipopos, txtnompos, "000000", matcar, matfl3)
               contar = contar + 1
            End If
          Else
           MsgBox "No se puede registrar una operacion antes de su fecha de inicio"
          End If
      End If
    Else
'no se encontro la informacion en la tabla de datos, se agrega la informacion
'la fecha de registro sera la fecha de inicio de la operacion
       Call GuardarPosDeuda(tipopos, txtnompos, "000000", matcar, matfl3)
       contar = contar + 1
    End If
   MensajeProc = "Comparando la operación " & txtcoperacion
   Call MostrarMensajeSistema(MensajeProc, frmProgreso.Label2, 0, Date, Time, NomUsuario)
   DoEvents
Next i
noreg = contar
sicomparar = False
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Sub GenPosPrimSwaps(ByVal fecha As Date, ByRef matpos2() As propPosDeuda, ByRef matfl2() As Variant, ByRef noreg As Integer, ByRef exito As Boolean)
Dim matpos1() As Variant
Dim matfl1() As estFlujosDeuda
Dim iguales As Boolean
Dim i As Long
Dim cswap1 As String
Dim indice As Long
Dim fregistro1 As Date


If ActivarControlErrores Then
On Error GoTo ControlErrores
End If

For i = 1 To UBound(matpos2, 1)
If UBound(matpos1, 1) > 0 Then
    cswap1 = matpos2(i, 1)
    indice = BuscarValorArray(cswap1, matpos1, 2)
    If indice <> 0 Then
     If fecha >= matpos1(i, 1) Then
     'se encontro la informacion en la posicion guardada
     'ahora hay que comparar las versiones correctas de informacion
     Call CompararFlujos(matfl1, matfl2, iguales)
        If Not iguales Then
        'es una nueva version de la informacion, se agrega
          MsgBox "Se guarda una nueva version del swap " & cswap1
          Call GuardarPosDeuda(1, "Real", "000000", matpos2, matfl2)
        End If
     Else
       MsgBox "No se puede registrar una operacion antes de su fecha de inicio"
     End If
    Else
       'no se encontro la informacion en la tabla de datos, se agrega la informacion
        Call GuardarPosDeuda(1, "Real", "000000", matpos2, matfl2)
    End If
Else
' la tabla de datos esta vacia, se agrega la informacion sin ninguna restriccion
  Call GuardarPosDeuda(1, "Real", "000000", matpos2, matfl2)
End If
Next i


On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function FiltrarCaracSwaps(matcar, i)
 ReDim matcarb(1 To 1) As New propPosSwaps
 Set matcarb(1) = matcar(i)
FiltrarCaracSwaps = matcarb
End Function

Sub CompararCSwaps(ByRef matcar1() As propPosSwaps, ByRef matcar2() As propPosSwaps, ByRef iguales As Boolean)
'matcar1 son los datos de riesgos
'matcar2 son los datos de la interfaz
'MsgBox matcar1(1, 8)
iguales = True
   If matcar1(1).intencion <> matcar2(1).intencion Then         'intencion
      iguales = False
      Exit Sub
   End If
   If matcar1(1).FCompraSwap <> matcar2(1).FCompraSwap Then     'fecha de inicio
      iguales = False
      Exit Sub
   End If
   If matcar1(1).FvencSwap <> matcar2(1).FvencSwap Then         'fecha de vencimiento
      iguales = False
      Exit Sub
   End If
   If matcar1(1).IntercIFSwap <> matcar2(1).IntercIFSwap Then    'intercambio inicial de flujos
      iguales = False
      Exit Sub
   End If
   If matcar1(1).IntercFFSwap <> matcar2(1).IntercFFSwap Then    'intercambio intermedio y final de flujos
      iguales = False
      Exit Sub
   End If
   If matcar1(1).RIntAct <> matcar2(1).RIntAct Then              'reinversion int activo
      iguales = False
      Exit Sub
   End If
   If matcar1(1).RIntPas <> matcar2(1).RIntPas Then              'reinversion int pasivo
      iguales = False
      Exit Sub
   End If
   If matcar1(1).TCActivaSwap <> matcar2(1).TCActivaSwap Then    'tasa interes activa
      iguales = False
      Exit Sub
   End If
   If matcar1(1).TCPasivaSwap <> matcar2(1).TCPasivaSwap Then    'tasa interes pasiva
      iguales = False
      Exit Sub
   End If
   If Abs(matcar1(1).STActiva - matcar2(1).STActiva) > 0.0000001 Then 'sobretasa activa
      iguales = False
      Exit Sub
   End If
   If Abs(matcar1(1).STPasiva - matcar2(1).STPasiva) > 0.0000001 Then 'sobretasa pasiva
      iguales = False
      Exit Sub
   End If
   If matcar1(1).ConvIntAct <> matcar2(1).ConvIntAct Then 'convencion de intereses activa
      iguales = False
      Exit Sub
   End If
   If matcar1(1).ConvIntPas <> matcar2(1).ConvIntPas Then 'convencion de intereses pasiva
      iguales = False
      Exit Sub
   End If
   If matcar1(1).ClaveProdSwap <> matcar2(1).ClaveProdSwap Then    'clave producto
      iguales = False
      Exit Sub
   End If
   If matcar1(1).ID_ContrapSwap <> matcar2(1).ID_ContrapSwap Then       'contraparte
      iguales = False
      Exit Sub
   End If
   If matcar1(1).EstructuralSwap <> matcar2(1).EstructuralSwap Then     'estructural
      iguales = False
      Exit Sub
   End If
End Sub

Sub CompararFlujos(ByRef matfl1() As estFlujosDeuda, ByRef matfl2() As Variant, ByRef iguales As Boolean)
Dim i As Long

'se filtran los flujos en la tabla de datos
iguales = True
'primer criterio: no tengan el mismo numero de flujos
If UBound(matfl1, 1) <> UBound(matfl2, 1) Then
  iguales = False
Else
'segundo criterio: comparar cada registro para encontrar alguna diferencia
For i = 1 To UBound(matfl1, 1)
   If matfl1(i).coperacion <> "" & matfl2(i, 1) & "" Then   'clave de la operacion
    iguales = False
    Exit Sub
   End If
   If matfl1(i).tpata <> matfl2(i, 2) Then             'sentido de la pata
    iguales = False
    Exit Sub
   End If
   If matfl1(i).finicio <> matfl2(i, 3) Then             'fecha de inicio del flujo
    iguales = False
    Exit Sub
   End If
   If matfl1(i).ffin <> matfl2(i, 4) Then             'fecha final del flujo
    iguales = False
    Exit Sub
   End If
   If matfl1(i).fpago <> matfl2(i, 5) Then             'fecha de descuento
    iguales = False
    Exit Sub
   End If
   If matfl1(i).pago_int <> matfl2(i, 6) Then            'pago de intereses
    iguales = False
    Exit Sub
   End If
   If matfl1(i).int_t_saldo <> matfl2(i, 7) Then            'saldo*intereses
    iguales = False
    Exit Sub
   End If
   If Abs(matfl1(i).saldo - matfl2(i, 8)) > 0.001 Then  'saldo
      iguales = False
      Exit Sub
   End If
   If Abs(matfl1(i).amort - matfl2(i, 9)) > 0.001 Then  'amortizacion
      iguales = False
      Exit Sub
   End If
Next i
End If
End Sub

Sub CompararFlujosD(ByRef matfl1() As estFlujosDeuda, ByRef matfl2() As Variant, ByRef iguales As Boolean)
Dim i As Long

'se filtran los flujos en la tabla de datos
iguales = True
'primer criterio: no tengan el mismo numero de flujos
If UBound(matfl1, 1) <> UBound(matfl2, 1) Then
  iguales = False
Else
'segundo criterio: comparar cada registro para encontrar alguna diferencia
For i = 1 To UBound(matfl1, 1)
   If matfl1(i).coperacion <> "" & matfl2(i, 1) & "" Then   'clave de la operacion
    iguales = False
    Exit Sub
   End If
   If matfl1(i).finicio <> matfl2(i, 3) Then             'fecha de inicio del flujo
    iguales = False
    Exit Sub
   End If
   If matfl1(i).ffin <> matfl2(i, 4) Then             'fecha final del flujo
    iguales = False
    Exit Sub
   End If
   If matfl1(i).fpago <> matfl2(i, 5) Then             'fecha de descuento
    iguales = False
    Exit Sub
   End If
   If matfl1(i).pago_int <> matfl2(i, 6) Then            'pago de intereses
    iguales = False
    Exit Sub
   End If
   If matfl1(i).int_t_saldo <> matfl2(i, 7) Then            'saldo*intereses
    iguales = False
    Exit Sub
   End If
   If Abs(matfl1(i).saldo - matfl2(i, 8)) > 0.001 Then  'saldo
      iguales = False
      Exit Sub
   End If
   If Abs(matfl1(i).amort - matfl2(i, 9)) > 0.001 Then  'amortizacion
      iguales = False
      Exit Sub
   End If
Next i
End If
End Sub


Function FiltrarFlujos1(ByVal fecha As Date, ByVal coperacion As String, ByRef matfl() As Variant)
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim kk As Long
Dim contar As Long

noreg = UBound(matfl, 1)
nocampos = UBound(matfl, 2)
ReDim mata(1 To nocampos, 1 To 1) As Variant
contar = 0
For i = 1 To noreg
    If fecha = matfl(i, 2) And coperacion = matfl(i, 5) Then
       contar = contar + 1
ReDim Preserve mata(1 To nocampos, 1 To contar) As Variant
       For kk = 1 To nocampos
           mata(kk, contar) = matfl(i, kk)
       Next kk
    End If
Next i
If contar <> 0 Then
   FiltrarFlujos1 = MTranV(mata)
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
   FiltrarFlujos1 = mata
End If

End Function

Function FiltrarFlujos2(coperacion, ByRef matfl() As Variant)
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim contar As Long
Dim kk As Long

noreg = UBound(matfl, 1)
nocampos = UBound(matfl, 2)
ReDim mata(1 To nocampos + 1, 1 To 1) As Variant
contar = 0
For i = 1 To noreg
If coperacion = "" & matfl(i, 1) & "" Then
contar = contar + 1
ReDim Preserve mata(1 To nocampos + 1, 1 To contar) As Variant
For kk = 1 To nocampos
 mata(kk, contar) = matfl(i, kk)
Next kk
mata(nocampos + 1, contar) = "0000000"
End If
Next i
FiltrarFlujos2 = MTranV(mata)
End Function

Sub GuardarPosSwaps(ByVal fregistro As Date, ByVal tipopos As Integer, ByVal txtpossim As String, ByVal horareg As String, ByRef matcar() As propPosSwaps, ByRef matfl() As Variant, ByRef exito As Boolean)
If ActivarControlErrores Then
   On Error GoTo hayerror
End If
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtfecha4 As String
Dim txtfecha5 As String
Dim txtborra1 As String
Dim txtborra2 As String
Dim txtcadena As String
Dim i As Integer
Dim j As Integer
Dim C_Posicion As Integer

  'se agregan los flujos del swap
  txtfecha1 = "to_date('" & Format(fregistro, "dd/mm/yyyy") & "','dd/mm/yyyy')"
  For i = 1 To UBound(matcar, 1)
     txtborra1 = "DELETE FROM " & TablaPosSwaps & " WHERE TIPOPOS = " & tipopos & " AND FECHAREG = " & txtfecha1 & " AND HORAREG=  '" & horareg & "' and COPERACION = '" & matcar(i).c_operacion & "' AND CPOSICION = " & matcar(i).C_Posicion
     txtborra2 = "DELETE FROM " & TablaFlujosSwapsO & " WHERE TIPOPOS = " & tipopos & " AND FECHAREG = " & txtfecha1 & " AND HORAREG =  '" & horareg & "' and COPERACION = '" & matcar(i).c_operacion & "' AND CPOSICION = " & matcar(i).C_Posicion
     ConAdo.Execute txtborra1
     ConAdo.Execute txtborra2
     txtfecha1 = "to_date('" & Format(fregistro, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfecha2 = "to_date('" & Format(matcar(i).FCompraSwap, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfecha3 = "to_date('" & Format(matcar(i).FvencSwap, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtcadena = "INSERT INTO " & TablaPosSwaps & " VALUES("
     txtcadena = txtcadena & tipopos & ","                                'tipo de posicion: real sim intradia
     txtcadena = txtcadena & txtfecha1 & ","                              'fecha de registro
     txtcadena = txtcadena & "'" & txtpossim & "',"                       'nombre de la posicion
     txtcadena = txtcadena & "'" & horareg & "',"                         'hora de la posicion
     txtcadena = txtcadena & "'" & matcar(i).intencion & "',"           'intencion
     txtcadena = txtcadena & "'" & matcar(i).EstructuralSwap & "',"     'estructural
     txtcadena = txtcadena & matcar(i).C_Posicion & ","                  'clave de posicion
     txtcadena = txtcadena & "'" & matcar(i).c_operacion & "',"          'Clave de operación ikos
     txtcadena = txtcadena & matcar(i).Tipo_Mov & ","                    'tipo de operacion
     txtcadena = txtcadena & txtfecha2 & ","                              'fecha inicio swap
     txtcadena = txtcadena & txtfecha3 & ","                              'fecha venc swap
     txtcadena = txtcadena & "'" & matcar(i).IntercIFSwap & "',"        'interi
     txtcadena = txtcadena & "'" & matcar(i).IntercFFSwap & "',"        'interf
     txtcadena = txtcadena & "'" & matcar(i).RIntAct & "',"             'acumula int activa
     txtcadena = txtcadena & "'" & matcar(i).RIntPas & "',"             'acumula int pasiva
     txtcadena = txtcadena & "'" & matcar(i).TCActivaSwap & "',"        'tasa cupon activa
     txtcadena = txtcadena & "'" & matcar(i).TCPasivaSwap & "',"        'tasa cupon pasiva
     txtcadena = txtcadena & matcar(i).STActiva & ","                   'st activa
     txtcadena = txtcadena & matcar(i).STPasiva & ","                   'st pasiva
     txtcadena = txtcadena & "'" & matcar(i).ConvIntAct & "',"          'conv int activa
     txtcadena = txtcadena & "'" & matcar(i).ConvIntPas & "',"          'conv int pasiva
     txtcadena = txtcadena & "'" & matcar(i).ClaveProdSwap & "',"       'clave de producto
     txtcadena = txtcadena & "'" & matcar(i).cProdSwapGen & "',"        'clave de valuacion
     txtcadena = txtcadena & Val(matcar(i).ID_ContrapSwap) & ","        'contraparte
     txtcadena = txtcadena & "'',"                                       'pidv
     txtcadena = txtcadena & "'')"                                       'ce_pidv
 
     ConAdo.Execute txtcadena
     MensajeProc = "Guardando caracteristicas del swap " & matcar(i).c_operacion
  Next i
  For i = 1 To UBound(matfl, 1)
     For j = 1 To UBound(matcar, 1)
         If matfl(i, 1) = matcar(j).c_operacion Then
            C_Posicion = matcar(j).C_Posicion
            Exit For
         End If
     Next j
     txtfecha1 = "to_date('" & Format(fregistro, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfecha3 = "to_date('" & Format(matfl(i, 3), "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfecha4 = "to_date('" & Format(matfl(i, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfecha5 = "to_date('" & Format(matfl(i, 5), "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtcadena = "INSERT INTO " & TablaFlujosSwapsO & " VALUES("
     txtcadena = txtcadena & tipopos & ","                     'tipo de posicion
     txtcadena = txtcadena & txtfecha1 & ","                   'fecha de registro
     txtcadena = txtcadena & "'" & txtpossim & "',"            'nombre de la posicion
     txtcadena = txtcadena & "'" & horareg & "',"              'hora de registro
     txtcadena = txtcadena & C_Posicion & ","                  'CLAVE DE LA POSICION
     txtcadena = txtcadena & "'" & matfl(i, 1) & "',"          'Clave de operación en ikos
     txtcadena = txtcadena & "'" & matfl(i, 2) & "',"          'tipo de pata
     txtcadena = txtcadena & txtfecha3 & ","                   'fecha inicio flujo
     txtcadena = txtcadena & txtfecha4 & ","                   'fecha fin flujo
     txtcadena = txtcadena & txtfecha5 & ","                   'fecha descuento flujo
     txtcadena = txtcadena & "'" & matfl(i, 6) & "',"          'paga intereses
     txtcadena = txtcadena & "'" & matfl(i, 7) & "',"          'saldo*intereses
     txtcadena = txtcadena & matfl(i, 8) & ","                 'saldo
     txtcadena = txtcadena & matfl(i, 9) & ","                 'amortizacion
     txtcadena = txtcadena & matfl(i, 10) & ")"                'tasa cupon observada
     ConAdo.Execute txtcadena
     AvanceProc = i / UBound(matfl, 1)
     MensajeProc = "Guardando los flujos de la operación " & matfl(i, 1) & " " & Format(AvanceProc, "##0.00 %")
     DoEvents
  Next i
  exito = True
On Error GoTo 0
Exit Sub
hayerror:
MsgBox error(Err())
exito = False
End Sub

Sub GuardarPosDeuda(ByVal tipopos As Integer, ByVal txtnompos As String, ByVal horareg As String, ByRef matcar() As propPosDeuda, ByRef matfl() As Variant)
Dim txtfecha As String
Dim txtborra1 As String
Dim txtborra2 As String
Dim i As Integer
Dim j As Integer
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtfecha4 As String
Dim txtfecha5 As String
Dim txtcadena As String
Dim cposicion As Integer
Dim coperacion As String
Dim fvalua As String
Dim toper As Integer
Dim fecha As Date

  'se agregan los flujos del swap

  For i = 1 To UBound(matcar, 1)
    txtfecha1 = "to_date('" & Format(matcar(i).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtborra1 = "DELETE FROM " & TablaPosDeuda & " WHERE TIPOPOS = " & tipopos & " AND FECHAREG = " & txtfecha1 & " AND NOMPOS = '" & txtnompos & "' AND HORAREG = '" & horareg & "'  and CPOSICION = " & matcar(i).C_Posicion & " AND COPERACION = '" & matcar(i).c_operacion & "'"
    txtborra2 = "DELETE FROM " & TablaFlujosDeudaO & " WHERE TIPOPOS = " & tipopos & " AND FECHAREG = " & txtfecha1 & " AND NOMPOS = '" & txtnompos & "' AND HORAREG = '" & horareg & "'   AND CPOSICION = " & matcar(i).C_Posicion & " AND COPERACION = '" & matcar(i).c_operacion & "'"
    ConAdo.Execute txtborra1
    ConAdo.Execute txtborra2
    txtfecha1 = "to_date('" & Format(matcar(i).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(matcar(i).FinicioDeuda, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha3 = "to_date('" & Format(matcar(i).FVencDeuda, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtcadena = "INSERT INTO " & TablaPosDeuda & " VALUES("
     txtcadena = txtcadena & "'" & tipopos & "',"                        'tipo de posicion: real sim intradia
     txtcadena = txtcadena & txtfecha1 & ","                             'fecha de registro
     txtcadena = txtcadena & "'" & txtnompos & "',"                      'nombre de la posicion si es simulada
     txtcadena = txtcadena & "'" & horareg & "',"                        'hora de la posicion
     txtcadena = txtcadena & matcar(i).C_Posicion & ","                  'clave de posicion
     txtcadena = txtcadena & "'" & matcar(i).c_operacion & "',"          'Clave de operación
     txtcadena = txtcadena & matcar(i).Tipo_Mov & ","                    'tipo de operacion
     txtcadena = txtcadena & txtfecha2 & ","                             'fecha inicio swap
     txtcadena = txtcadena & txtfecha3 & ","                             'fecha venc swap
     txtcadena = txtcadena & "'" & matcar(i).InteriDeuda & "',"          'interi
     txtcadena = txtcadena & "'" & matcar(i).InterfDeuda & "',"          'interf
     txtcadena = txtcadena & "'" & matcar(i).RintDeuda & "',"            'reinvertir intereses
     txtcadena = txtcadena & "'" & Trim(matcar(i).TRefDeuda) & "',"      'tasa cupon
     txtcadena = txtcadena & matcar(i).SpreadDeuda & ","                 'st activa
     txtcadena = txtcadena & "'" & matcar(i).ConvIntDeuda & "',"         'conv calc INTERESES
     txtcadena = txtcadena & "'" & matcar(i).fValuacion & "',"           'tipo de instrumento
     txtcadena = txtcadena & "'N')"                                      'posicion relacionada
     ConAdo.Execute txtcadena
     MensajeProc = "Guardando caracteristicas de la operacion " & matcar(i).c_operacion
  Next i
  
  For i = 1 To UBound(matfl, 1)
      For j = 1 To UBound(matcar, 1)
         If matcar(j).c_operacion = matfl(i, 1) Then
            fecha = matcar(j).fechareg
            cposicion = matcar(j).C_Posicion
            Exit For
         End If
      Next j
     txtfecha1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfecha3 = "to_date('" & Format(matfl(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfecha4 = "to_date('" & Format(matfl(i, 3), "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfecha5 = "to_date('" & Format(matfl(i, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtcadena = "INSERT INTO " & TablaFlujosDeudaO & " VALUES("
     txtcadena = txtcadena & tipopos & ","                     'tipo de posicion
     txtcadena = txtcadena & txtfecha1 & ","                   'fecha de registro
     txtcadena = txtcadena & "'" & txtnompos & "',"            'nombre de la posicion
     txtcadena = txtcadena & "'" & horareg & "',"              'hora de la posicion
     txtcadena = txtcadena & cposicion & ","                   'Clave de la posicion
     txtcadena = txtcadena & "'" & matfl(i, 1) & "',"          'Clave de operación en ikos
     txtcadena = txtcadena & txtfecha3 & ","                   'fecha inicio flujo
     txtcadena = txtcadena & txtfecha4 & ","                   'fecha fin flujo
     txtcadena = txtcadena & txtfecha5 & ","                   'fecha de descuento de flujos
     txtcadena = txtcadena & "'" & matfl(i, 5) & "',"          'paga intereses en el periodo
     txtcadena = txtcadena & "'" & matfl(i, 6) & "',"          'int por saldo
     txtcadena = txtcadena & matfl(i, 7) & ","                 'saldo
     txtcadena = txtcadena & matfl(i, 8) & ","                 'amortizacion
     txtcadena = txtcadena & matfl(i, 9) & ")"                 'tasa cupon observada
     ConAdo.Execute txtcadena
     AvanceProc = i / UBound(matfl, 1)
     MensajeProc = "Guardando los flujos de la operación " & matfl(i, 1) & " " & Format(AvanceProc, "##0.00 %")
     DoEvents
  Next i
End Sub

Function FiltrarSwapsVigentes(ByVal fecha As Date, ByRef matb() As Variant) As Variant()
Dim nocampos As Long
Dim contar2 As Long
Dim i As Long
Dim nocampos2 As Long
Dim j As Long

nocampos2 = UBound(matb, 2)
ReDim matd(1 To nocampos2, 1 To 1) As Variant
contar2 = 0
For i = 1 To UBound(matb, 1)
If (matb(i, 3) <= fecha And fecha <= matb(i, 4)) Then
 contar2 = contar2 + 1
 ReDim Preserve matd(1 To nocampos2, 1 To contar2) As Variant
 For j = 1 To nocampos2
 matd(j, contar2) = matb(i, j)
 Next j
End If
Next i
FiltrarSwapsVigentes = RutinaOrden(MTranV(matd), 1, SRutOrden)
End Function

Function FiltrarFlujosSwaps(ByRef matfl() As Variant, ByRef matcar() As Variant) As Variant()
Dim nocampos1 As Long
Dim i As Long
Dim indice As Long
Dim contar1 As Long

nocampos1 = UBound(matfl, 2) + 5
ReDim matc(1 To nocampos1, 1 To 1) As Variant
contar1 = 0
'se filtran los flujos de los swaps vigentes
For i = 1 To UBound(matfl, 1)
 indice = BuscarValorArray(matfl(i, 1), matcar, 1)
 If indice <> 0 Then
  contar1 = contar1 + 1
  ReDim Preserve matc(1 To nocampos1, 1 To contar1) As Variant
  matc(1, contar1) = matfl(i, 1)        'emision
  matc(2, contar1) = matfl(i, 2)        'pata
  matc(3, contar1) = matcar(indice, 2)  'intencion
  matc(4, contar1) = matfl(i, 3)        'f inicio
  matc(5, contar1) = matfl(i, 4)        'f final
  matc(6, contar1) = matfl(i, 5)        'pago intereses
  matc(7, contar1) = matfl(i, 6)        'reinversion intereses
  matc(8, contar1) = matfl(i, 7)        'saldo
  matc(9, contar1) = matfl(i, 8)        'amortizacion
  matc(10, contar1) = matfl(i, 9)       'tasa cupon
  matc(11, contar1) = matfl(i, 10)      'spread
 End If
 AvanceProc = i / UBound(matfl, 1)
 MensajeProc = "Filtrando los flujos validos de la posicion " & Format(AvanceProc, "##0.00 %")
Next i
FiltrarFlujosSwaps = MTranV(matc)
End Function

Function FiltrarFlujosSwaps2(ByRef matfl() As Variant, ByRef matcar() As Variant) As Variant()
Dim noreg As Long
Dim nocampos1 As Long
Dim i As Long
Dim contar1 As Long
Dim indice As Long


noreg = UBound(matcar, 1)
nocampos1 = UBound(matfl, 2) + 2
ReDim matcar1(1 To noreg, 1 To 1) As Variant
For i = 1 To UBound(matcar, 1)
matcar1(i, 1) = matcar(i, 1) & matcar(i, 2)
Next i
matcar1 = RutinaOrden(matcar1, 1, SRutOrden)
ReDim matc(1 To nocampos1, 1 To 1) As Variant
contar1 = 0
'se filtran los flujos de los swaps vigentes
For i = 1 To UBound(matfl, 1)
 indice = BuscarValorArray(matfl(i, 1) & matfl(i, 2), matcar1, 1)
 If indice <> 0 Then
  contar1 = contar1 + 1
  ReDim Preserve matc(1 To nocampos1, 1 To contar1) As Variant
  matc(1, contar1) = matfl(i, 2)             'clave ikos
  matc(2, contar1) = matfl(i, 3)             'pata
  matc(3, contar1) = matfl(i, 4)             'f inicio
  matc(4, contar1) = matfl(i, 5)             'f final
  matc(5, contar1) = matfl(i, 6)             'f c intereses
  matc(6, contar1) = matfl(i, 7)             'acumula int
  matc(7, contar1) = matfl(i, 8)             'saldo
  matc(8, contar1) = matfl(i, 9)             'amortizacion
  If Val(matfl(i, 10)) <> 0 Then
  matc(9, contar1) = Val(matfl(i, 10))       'tasa cupon
  Else
  matc(9, contar1) = Trim(matfl(i, 10))      'tasa cupon
  End If
  matc(10, contar1) = matfl(i, 11)           'spread
  matc(11, contar1) = "S " & matfl(i, 2) & " P " & matfl(i, 3)    'clave de la pata
  matc(12, contar1) = matc(11, contar1) & " " & CLng(matfl(i, 4)) 'clave de ordenacion de los flujos
 End If
  AvanceProc = i / UBound(matfl, 1)
  MensajeProc = "Filtrando los flujos validos de la posicion " & Format(AvanceProc, "##0.00 %")

Next i
FiltrarFlujosSwaps2 = RutinaOrden(MTranV(matc), 12, SRutOrden)
End Function

Function CrearFiltroFlujos1(ByVal tipopos As Integer, ByVal fecha As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String, ByVal tpata As String) As String
On Error GoTo hayerror

Dim fregistro As String
Dim txtfiltro As String
   fregistro = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro = "select * from " & TablaFlujosSwapsO & " where TIPOPOS = " & tipopos
   txtfiltro = txtfiltro & " AND FECHAREG = " & fregistro
   txtfiltro = txtfiltro & " AND NOMPOS = '" & txtnompos & "'"
   txtfiltro = txtfiltro & " AND HORAREG = '" & horareg & "'"
   txtfiltro = txtfiltro & " AND CPOSICION = " & cposicion
   txtfiltro = txtfiltro & " AND COPERACION = '" & coperacion & "'"
   txtfiltro = txtfiltro & " AND TPATA = '" & tpata & "' ORDER BY  FINICIO"
CrearFiltroFlujos1 = txtfiltro
On Error GoTo 0
Exit Function
hayerror:
MsgBox "crearfiltroflujos1 " & error(Err())
End Function

Function CrearFiltroDeudaEsp(ByVal tipopos As Integer, ByVal fecha As Date, ByVal txtnompos As String, ByVal horareg As String, ByVal cposicion As Integer, ByVal coperacion As String) As String
Dim fregistro As String
Dim txtfiltro As String
   fregistro = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro = "select * from " & TablaFlujosDeudaO & " where TIPOPOS = " & tipopos
   txtfiltro = txtfiltro & " AND FECHAREG = " & fregistro
   txtfiltro = txtfiltro & " AND NOMPOS = '" & txtnompos & "'"
   txtfiltro = txtfiltro & " AND HORAREG = '" & horareg & "'"
   txtfiltro = txtfiltro & " AND CPOSICION = " & cposicion
   txtfiltro = txtfiltro & " AND COPERACION = '" & coperacion & "'"
   txtfiltro = txtfiltro & " ORDER BY FINICIO"
CrearFiltroDeudaEsp = txtfiltro
End Function

Function FiltrarFlujosSwaps3(ByRef matpos() As propPosSwaps, ByVal simav As Boolean, ByRef exito As Boolean) As estFlujosDeuda()
If ActivarControlErrores Then
On Error GoTo hayerror
End If
Dim noreg As Long
Dim noreg1 As Long
Dim noreg2 As Long

Dim nocampos1 As Long
Dim ncol1 As Long
Dim ncol2 As Long
Dim contar1 As Long
Dim i As Long
Dim j As Long
Dim kk As Long
Dim tiempo1 As Date
Dim tiempo2 As Date
Dim tiempo As Date
Dim tpromedio As Date
Dim coperacion As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim matflswap1() As New estFlujosDeuda
Dim matflujos2() As New estFlujosDeuda
Dim matc() As New estFlujosDeuda
exito = True
noreg = UBound(matpos, 1)
'se filtran los flujos de los swaps vigentes
contar1 = 0
For i = 1 To noreg
'la llave de filtrado es fecha de registro
tiempo1 = Time
   If Not EsVariableVacia(matpos(i).c_operacion) Then
        txtfiltro1 = CrearFiltroFlujos1(matpos(i).tipopos, matpos(i).fechareg, matpos(i).nompos, matpos(i).HoraRegOp, matpos(i).C_Posicion, matpos(i).c_operacion, "B")
        txtfiltro2 = CrearFiltroFlujos1(matpos(i).tipopos, matpos(i).fechareg, matpos(i).nompos, matpos(i).HoraRegOp, matpos(i).C_Posicion, matpos(i).c_operacion, "C")
        matflswap1 = LeerFlujosSwaps(txtfiltro1, False)
        matflujos2 = LeerFlujosSwaps(txtfiltro2, False)
        noreg1 = UBound(matflswap1, 1)
        If noreg1 <> 0 Then
           matpos(i).IFlujoActSwap = contar1 + 1
           contar1 = contar1 + noreg1
           matpos(i).FFlujoActSwap = contar1
           ReDim Preserve matc(1 To contar1)
           For j = 1 To noreg1
               Set matc(contar1 + j - noreg1) = matflswap1(j)
           Next j
        Else
           exito = False
        End If
        noreg2 = UBound(matflujos2, 1)
        If noreg2 <> 0 Then
           matpos(i).IFlujoPasSwap = contar1 + 1
           contar1 = contar1 + noreg2
           matpos(i).FFlujoPasSwap = contar1
           ReDim Preserve matc(1 To contar1)
           For j = 1 To noreg2
               Set matc(contar1 + j - noreg2) = matflujos2(j)
           Next j
        Else
           exito = False
        End If
       

   End If
   tiempo2 = Time
   tiempo = tiempo2 - tiempo1
   tpromedio = tpromedio + tiempo
   If simav Then
      AvanceProc = i / UBound(matpos, 1)
      MensajeProc = "Cargando los flujos de la posicion de swaps " & Format(AvanceProc, "###0.00 %")
   End If
   DoEvents
Next i
FiltrarFlujosSwaps3 = matc
On Error GoTo 0
Exit Function
hayerror:
 MsgBox "filtrarflujosswaps3 " & error(Err())
End Function

Function FiltrarFlujosDeuda(ByRef matpos() As propPosDeuda, ByRef exito As Boolean) As estFlujosDeuda()
Dim noreg As Long
Dim noreg1 As Long
Dim contar1 As Long
Dim i As Long
Dim j As Long
Dim txtfiltro As String
Dim matflujos() As New estFlujosDeuda
Dim matc() As New estFlujosDeuda
noreg = UBound(matpos, 1)
'se filtran los flujos de los swaps vigentes
contar1 = 0
exito = True
For i = 1 To noreg
        txtfiltro = CrearFiltroDeudaEsp(matpos(i).tipopos, matpos(i).fechareg, matpos(i).nompos, matpos(i).HoraRegOp, matpos(i).C_Posicion, matpos(i).c_operacion)
        matflujos = LeerFlujosDeuda(txtfiltro, False)
        noreg1 = UBound(matflujos, 1)
        If noreg1 <> 0 Then
           matpos(i).IFlujoDeuda = contar1 + 1
           contar1 = contar1 + noreg1
           matpos(i).FFlujoDeuda = contar1
           ReDim Preserve matc(1 To contar1)
           For j = 1 To UBound(matflujos, 1)
               Set matc(contar1 + j - noreg1) = matflujos(j)
           Next j
        Else
            exito = False
        End If
   AvanceProc = i / UBound(matpos, 1)
   MensajeProc = "Cargando los flujos validos de la posicion de deuda " & Format(AvanceProc, "###0.00 %")
   DoEvents
Next i
FiltrarFlujosDeuda = matc
End Function

Function DetermSiEmFlujos(ByVal tv As String, ByVal emision As String, ByVal serie As String)
Dim i As Integer
Dim siencontro As Boolean
For i = 1 To UBound(MatTValSTCupon, 1)
    siencontro = DetermSiEstaTablaVal(tv, emision, serie, i, MatTValSTCupon)
    If siencontro Then
       DetermSiEmFlujos = True
       Exit Function
    End If
Next i
For i = 1 To UBound(MatTValSTDesc, 1)
    siencontro = DetermSiEstaTablaVal(tv, emision, serie, i, MatTValSTDesc)
    If siencontro Then
       DetermSiEmFlujos = True
       Exit Function
    End If
Next i
For i = 1 To UBound(MatTValBonos, 1)
    siencontro = DetermSiEstaTablaVal(tv, emision, serie, i, MatTValBonos)
    If siencontro Then
       DetermSiEmFlujos = True
       Exit Function
    End If
Next i
DetermSiEmFlujos = False
End Function

Function DetermSiEstaTablaVal(ByVal tv As String, ByVal emision As String, ByVal serie As String, ByVal indice As Integer, ByRef mata() As Variant)
If mata(indice, 1) <> "*" And mata(indice, 2) <> "*" And mata(indice, 3) <> "*" Then
   If tv = mata(indice, 1) And emision = mata(indice, 2) And serie = mata(indice, 3) Then DetermSiEstaTablaVal = True
ElseIf mata(indice, 1) <> "*" And mata(indice, 2) <> "*" And mata(indice, 3) = "*" Then
   If tv = mata(indice, 1) And emision = mata(indice, 2) Then DetermSiEstaTablaVal = True
ElseIf mata(indice, 1) <> "*" And mata(indice, 2) = "*" And mata(indice, 3) <> "*" Then
   If tv = mata(indice, 1) And serie = mata(indice, 3) Then DetermSiEstaTablaVal = True
ElseIf mata(indice, 1) = "*" And mata(indice, 2) <> "*" And mata(indice, 3) <> "*" Then
   If emision = mata(indice, 2) And serie = mata(indice, 3) Then DetermSiEstaTablaVal = True
ElseIf mata(indice, 1) <> "*" And mata(indice, 2) = "*" And mata(indice, 3) = "*" Then
   If tv = mata(indice, 1) Then DetermSiEstaTablaVal = True
ElseIf mata(indice, 1) = "*" And mata(indice, 2) <> "*" And mata(indice, 3) = "*" Then
   If emision = mata(indice, 2) Then DetermSiEstaTablaVal = True
ElseIf mata(indice, 1) = "*" And mata(indice, 2) = "*" And mata(indice, 3) <> "*" Then
   If serie = mata(indice, 3) Then DetermSiEstaTablaVal = True
End If
End Function


Sub GuardaResEfRetroSwaps(ByVal fecha As Date, ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef matefic() As Variant, ByRef obj1 As ADODB.Connection)
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtcadena As String
Dim noreg As Long
Dim i As Long
Dim jj As Long
Dim j As Long

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
obj1.Execute "DELETE FROM " & TablaEficSwaps & " WHERE FECHA = " & txtfecha
obj1.Execute "DELETE FROM " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha
noreg = UBound(matefic, 1)
For i = 1 To noreg
    If Not IsNull(matefic(i, 27)) Then
       txtcadena = "INSERT INTO " & TablaEficSwaps & " VALUES("
       txtcadena = txtcadena & txtfecha & ","                     'fecha del calculo de la eficiencia
       txtcadena = txtcadena & txtfecha1 & ","                    'fecha de inicio del analisis
       txtcadena = txtcadena & txtfecha2 & ","                    'fecha final del analisis
       txtcadena = txtcadena & "'" & matefic(i, 1) & "',"         'la Clave de operación
       txtcadena = txtcadena & "'" & matefic(i, 2) & "',"         'pos primaria activa
       txtcadena = txtcadena & "'" & matefic(i, 3) & "',"         'pos primaria pasiva
       txtcadena = txtcadena & "'" & matefic(i, 4) & "',"         '
       txtcadena = txtcadena & "'" & matefic(i, 5) & "',"         'FECHA DE VENCIMIENTO
       txtcadena = txtcadena & "'" & matefic(i, 6) & "',"         'tipo de eficiencia
       For jj = 7 To 27
           If matefic(i, jj) <> 0 Then
              txtcadena = txtcadena & matefic(i, jj) & ","
           Else
              txtcadena = txtcadena & "0,"
           End If
       Next jj
       txtcadena = txtcadena & "0)"                                'eficiencia prospectiva
       obj1.Execute txtcadena
End If
   txtcadena = "INSERT INTO " & TablaEficienciaCob & " VALUES("
   txtcadena = txtcadena & txtfecha & ","                     'fecha del calculo de la eficiencia
   txtcadena = txtcadena & txtfecha1 & ","                    'fecha de inicio del analisis
   txtcadena = txtcadena & txtfecha2 & ","                    'fecha final del analisis
   txtcadena = txtcadena & "'" & matefic(i, 1) & "',"         'la Clave de operación
   For j = 1 To 12
       txtcadena = txtcadena & "0,"
   Next j
 txtcadena = txtcadena & "'" & matefic(i, 27) & "',"         'tipo de eficiencia
 txtcadena = txtcadena & "0)"                                'tipo de eficiencia
 obj1.Execute txtcadena
MensajeProc = "Guardando los resultados del calculo de efectividad"
DoEvents
Next i
MensajeProc = "Se transfirieron " & noreg & " registros"
End Sub

Sub GuardaResEfRetroFwds(ByVal fecha As Date, ByRef matefic() As Variant, ByRef obj1 As ADODB.Connection)
Dim txtfecha As String
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim txtcadena As String

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
noreg = UBound(matefic, 1)
For i = 1 To noreg
If Not EsVariableVacia(matefic(i, 12)) Then
 obj1.Execute "DELETE FROM " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha & " AND CLAVE_SWAP = '" & matefic(i, 2) & "'"
 txtcadena = "INSERT INTO " & TablaEficienciaCob & " VALUES("
 txtcadena = txtcadena & txtfecha & ","                     'fecha del calculo de la eficiencia
 txtcadena = txtcadena & txtfecha & ","                     'fecha de inicio del analisis
 txtcadena = txtcadena & txtfecha & ","                     'fecha final del analisis
 txtcadena = txtcadena & "'" & matefic(i, 2) & "',"         'la clave de la operacion
 For j = 1 To 12
     txtcadena = txtcadena & "0,"                           '
 Next j
txtcadena = txtcadena & Format(100 * matefic(i, 14), "###,##0.00") & ","     'eficiencia retrospectiva
If Not IsNull(matefic(i, 14)) = 0 Then
   txtcadena = txtcadena & matefic(i, 14) & ")"           'eficiencia pro
Else
   txtcadena = txtcadena & "0)"                          'eficiencia pro
End If
obj1.Execute txtcadena
End If
Next i
MensajeProc = "Se transfirieron " & noreg & " registros"
End Sub

Sub GuardaMatArchTexto(mata, ByVal txtarch As String)
Dim noreg As Integer
Dim nocols As Integer
Dim txtsalida As String
Dim i As Integer
Dim j As Integer
Dim exitoarch As Boolean

noreg = UBound(mata, 1)
nocols = UBound(mata, 2)
frmCalVar.CommonDialog1.FileName = txtarch
frmCalVar.CommonDialog1.ShowSave
txtarch = frmCalVar.CommonDialog1.FileName
Call VerificarSalidaArchivo(txtarch, 1, exitoarch)
If exitoarch Then
For i = 1 To noreg
    txtsalida = ""
    For j = 1 To nocols
        txtsalida = txtsalida & mata(i, j) & Chr(9)
    Next j
    Print #1, txtsalida
Next i
Close #1
End If
End Sub

Sub VerificarSalidaArchivo(ByVal nomarch As String, ByVal txtclave As String, ByRef exito As Boolean)

On Error GoTo corregir
Open nomarch For Output As #txtclave
Close #txtclave
Dim txtnomarch As String
'si es exitosa la apertura se vuelve a abrir
Open nomarch For Output As #txtclave
exito = True
Exit Sub
corregir:
'como ocurrio un error de lectura

If Err() = 52 Then
 MsgBox "No existe la ruta de escritura"
 exito = False
ElseIf Err() = 70 Then
 MsgBox "Acceso denegado"
 exito = False
Else
 MsgBox "VerificarSalidaArchivo:" & error(Err())
 exito = False
End If
End Sub

Sub CrearMatFRiesgo(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef matfr() As Variant, ByRef txtmsg As String, ByRef exito As Boolean)
'esta es la tabla de factores de riesgo con la nueva estructura
'se cuentan los registros
Dim mfechas() As Date
Dim nodays As Long
Dim i As Long
Dim j As Long
Dim esnulo As Boolean
Dim mata() As Variant

exito = True
mfechas = LeerFechasVaR(fecha1, fecha2)
nodays = UBound(mfechas, 1)
If nodays <> 0 Then
   ReDim matfr(1 To nodays, 1 To NoFactores + 1) As Variant
   For i = 1 To nodays
       matfr(i, 1) = mfechas(i, 1)
   Next i
   For i = 1 To NoFactores
       mata = Leer1FRiesgoxVaR(fecha1, fecha2, MatCaracFRiesgo(i).nomFactor, MatCaracFRiesgo(i).plazo, esnulo)
       For j = 1 To nodays
           matfr(j, i + 1) = mata(j, 2)
       Next j
       AvanceProc = i / NoFactores
       MensajeProc = "Leyendo los factores de riesgo " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   exito = True
   txtmsg = "El proceso finalizo correctamente"
Else
    ReDim matfr(0 To 0, 1 To NoFactores + 1) As Variant
    MensajeProc = "No hay datos en la tabla de factores de riesgo para el rango de fechas"
    txtmsg = MensajeProc
    MsgBox txtmsg
    exito = False
End If

End Sub

Function BuscarFechaNBFR(ByVal fecha As Date) As Date
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim valor As Double
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT MAX(FECHA) FROM " & TablaFRiesgoO & " WHERE FECHA <= " & txtfecha
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   valor = rmesa.Fields(0)
   rmesa.Close
Else
   valor = 0
End If
BuscarFechaNBFR = valor
End Function

Function CalcFlujosEmision(ByVal fecha1 As Date, ByVal f_vence As Date, ByVal vn As Double, ByVal pcupon As Long) As Variant()
Dim fechax As Date
Dim inicio As Long
Dim dxv As Long
Dim nc As Long
Dim i As Long

nc = (f_vence - fecha1) / pcupon
ReDim matf(1 To nc, 1 To 4) As Variant
matf(1, 1) = fecha1
For i = 1 To nc
    If i <> 1 Then
       matf(i, 1) = matf(i - 1, 2)
    End If
    matf(i, 2) = matf(i, 1) + pcupon
    matf(i, 3) = vn
    If i <> nc Then
       matf(i, 4) = 0
    Else
       matf(i, 4) = vn
    End If
Next i
CalcFlujosEmision = matf
End Function

Sub GuardaFlujosMD(ByVal txtemision As String, ByVal fregistro As Date, ByRef mata() As Variant)
Dim noreg As Long
Dim i As Long
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtcadena As String

noreg = UBound(mata, 1)
ConAdo.Execute "DELETE FROM " & TablaFlujosMD & " where EMISION = '" & txtemision & "'"
For i = 1 To noreg
    txtfecha = "to_date('" & Format(fregistro, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha1 = "to_date('" & Format(mata(i, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(mata(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "INSERT INTO " & TablaFlujosMD & " VALUES("
    txtcadena = txtcadena & "'" & txtemision & "',"    'EMISION
    txtcadena = txtcadena & txtfecha & ","             'fecha de registro
    txtcadena = txtcadena & txtfecha1 & ","            'INICIO DEL FLUJO
    txtcadena = txtcadena & txtfecha2 & ","            'FIN DEL FLUJO
    txtcadena = txtcadena & mata(i, 3) & ","           'NOCIONAL
    txtcadena = txtcadena & mata(i, 4) & ","           'AMORTIZACION
    txtcadena = txtcadena & "0," & mata(i, 2) - mata(i, 1) & ")"                   'tasa y periodo del cupon
    ConAdo.Execute txtcadena
    Call MostrarMensajeSistema("Guardando los flujos de la emision " & txtemision & " " & Format(AvanceProc, "###0.00"), frmProgreso.Label2, 0, Date, Time, NomUsuario)
    DoEvents
Next i
End Sub

Sub GenerarFlujosMD(ByVal fecha As Date, ByRef noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim matem() As String
Dim MatVPrecios() As New propVecPrecios
Dim mindvp() As Variant
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim rmesa As New ADODB.recordset

exito = False
MatVPrecios = LeerVPrecios(fecha, mindvp) 'leer el vector de precios
'se lee la posicion del dia
If UBound(MatVPrecios, 1) <> 0 Then
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro2 = "SELECT C_EMISION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1"
   txtfiltro2 = txtfiltro2 & " AND (TOPERACION = 1 OR TOPERACION = 4)"
   txtfiltro2 = txtfiltro2 & " GROUP BY C_EMISION ORDER BY C_EMISION"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      ReDim matem(1 To noreg) As String
      For i = 1 To noreg
          matem(i) = rmesa.Fields(0)
          rmesa.MoveNext
      Next i
      rmesa.Close
      Call GenerarFlujosMD2(matem, MatVPrecios, mindvp, exito)
      If exito Then
         txtmsg = "El proceso finalizo correctamente"
      End If
   Else
      exito = False
   End If
Else
   exito = False
End If

End Sub

Function BuscarFlujosBD(ByVal cemision As String) As Long
Dim txtfiltro As String
Dim rmesa As New ADODB.recordset

    txtfiltro = "SELECT COUNT(*) FROM " & TablaFlujosMD & " WHERE EMISION = '" & cemision & "'"
    rmesa.Open txtfiltro, ConAdo
    BuscarFlujosBD = rmesa.Fields(0)
    rmesa.Close
End Function

Sub GenerarFlujosMD2(ByRef matem() As String, ByRef MatVPrecios() As propVecPrecios, ByRef mindvp() As Variant, ByRef exito As Boolean)
Dim matpos1() As propPosMD
Dim matp() As Variant
Dim noreg As Long
Dim noregvp As Long
Dim contar As Long
Dim i As Long
Dim nemision As String
Dim indice As Long
Dim indice1 As Long
Dim indice2 As Long
Dim tv As String
Dim txtemision As String
Dim emision As String
Dim serie As String
Dim pc As Integer
Dim vn As Double
Dim finicio As Date
Dim ffinal As Date

Dim matfl1() As Variant

exito = False

If UBound(matem, 1) > 0 Then
noreg = UBound(matem, 1)
noregvp = UBound(MatVPrecios, 1)
If noregvp <> 0 Then
contar = 0
For i = 1 To noreg
    nemision = matem(i)
    noreg = BuscarFlujosBD(nemision)
    If noreg = 0 Then
       indice = BuscarValorArray(nemision, mindvp, 1)
       If indice <> 0 Then
          indice1 = mindvp(indice, 2)
          tv = MatVPrecios(indice1).tv
          emision = MatVPrecios(indice1).emision
          serie = MatVPrecios(indice1).serie
          txtemision = MatVPrecios(indice1).c_emision
          pc = MatVPrecios(indice1).pcupon
          vn = MatVPrecios(indice1).vnominal
          finicio = MatVPrecios(indice1).femision
          ffinal = MatVPrecios(indice1).fvenc
          indice2 = EncuentraEmisionBase(tv, emision, serie)
          If indice2 <> 0 Then
             If MatParamEmisiones(indice2, 5) <> 0 Then
                matfl1 = CalcFlujosEmision(finicio, ffinal, vn, MatParamEmisiones(indice2, 5))
             Else
                matfl1 = CalcFlujosEmision(finicio, ffinal, vn, pc)
             End If
             Call GuardaFlujosMD(txtemision, finicio, matfl1)
             contar = contar + 1
          Else
             If pc <> 0 Then
                MensajeProc = "no hay informacion para la emision " & tv & emision & serie & " en la base de datos"
                exito = False
             End If
          End If
       Else
          MensajeProc = "Falta la emision " & nemision & " en el vector del día"
       End If
    End If
Next i
noreg = contar
exito = True
Else
 MensajeProc = "No hay Vector de precios para esta fecha"
 Call MostrarMensajeSistema(MensajeProc, frmProgreso.Label2, 0, Date, Time, NomUsuario)
 exito = False
End If
Else

End If
End Sub

Function FiltrarPosDirectoMD(ByRef matp() As propPosMD)
Dim nocols As Integer
Dim contar As Integer
Dim i As Integer
Dim j As Integer
ReDim matp1(1 To 1) As New propPosMD
contar = 0
For i = 1 To UBound(matp, 1)
If matp(i).Tipo_Mov = 1 Or matp(i).Tipo_Mov = 4 Then
   contar = contar + 1
 ReDim Preserve matp1(1 To contar) As New propPosMD
   Set matp1(contar) = matp(i)
End If
Next i
FiltrarPosDirectoMD = matp1
End Function

Sub TratamientoErrores(clave)
If clave = -2147467259 Then
 Call reiniciarConex
Else
  MsgBox error(clave)
End If
End Sub

Sub reiniciarConex()
On Error Resume Next
   Call CerrarTablas
   ConAdo.Close
   conAdoBD.Close
   Call IniciarConexOracle(ConAdo, OpcionBDatos)
   Call IniciarConexOracle(conAdoBD, BDIKOS)
   Call AbrirTablas
On Error GoTo 0
End Sub

Function FiltrarRelPrimFecha(ByVal fecha As Date, ByRef matr() As propRelSwapPrim) As propRelSwapPrim()
'se filtran las relaciones de posicion primarias vigentes
Dim contar As Long
Dim i As Long
Dim j As Long
Dim mats() As New propRelSwapPrim
contar = 0
ReDim mats(1 To 1)
For i = 1 To UBound(matr, 1)
    If fecha >= matr(i).finicio And fecha < matr(i).ffin Then
       contar = contar + 1
       ReDim Preserve mats(1 To contar)
       Set mats(contar) = matr(i)
    End If
Next i
FiltrarRelPrimFecha = mats
End Function




Function CrearEstRepEmisiones(ByRef matpos() As Variant) As Variant()
Dim noreg As Long
Dim i As Long
Dim matc() As Variant
Dim j As Long
Dim noreg1 As Long

noreg = UBound(matpos, 1)
ReDim matb(1 To noreg, 1 To 1)
For i = 1 To noreg
    matb(i, 1) = matpos(i).Tipo_Mov & matpos(i).cEmisionMD
Next i
matc = ObtFactUnicos(matb, 1)
noreg1 = UBound(matc, 1)
ReDim matd(1 To noreg1, 1 To 4) As Variant
For i = 1 To noreg1
For j = 1 To noreg
If matc(i, 1) = matb(j, 1) Then
   matd(i, 2) = matb(j, 1)
   matd(i, 2) = matpos(j).cEmisionMD
   matd(i, 3) = matpos(j).Tipo_Mov
   matd(i, 4) = matpos(j).C_Posicion
   Exit For
End If
Next j
Next i
CrearEstRepEmisiones = matd
End Function


Sub AnalisisBack(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim indice1 As Integer
Dim indice2 As Integer
Dim noreg As Integer
Dim i As Integer
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtcad As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim contar As Long
Dim probe As Double
Dim nivconf As Double
Dim estadis As Double
Dim rmesa As New ADODB.recordset

 indice1 = BuscarValorArray(fecha1, MatFechasVaR, 1)
 indice2 = BuscarValorArray(fecha2, MatFechasVaR, 1)
 noreg = indice2 - indice1 + 1
 ReDim matback1(1 To noreg, 1 To 10) As Variant
 For i = 1 To noreg
  matback1(i, 1) = MatFechasVaR(indice1 + i - 1, 1)
 Next i
'ahora se procede a cargar los resultados guardados en la tabla de datos
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtcad = "(fecha >= " & txtfecha1 & " and fecha <= " & txtfecha2 & ") AND MPOSICION = '" & txtportBanobras & "' AND PORTAFOLIO = 'TOTALES' ORDER BY FECHA"
txtfiltro1 = "select count(*) from " & TablaBackPort & " WHERE " & txtcad
rmesa.Open txtfiltro1, ConAdo
noreg1 = rmesa.Fields(0)
rmesa.Close
If noreg1 <> 0 Then
 ReDim mata(1 To noreg1, 1 To 6) As Variant
 txtfiltro2 = "select * from " & TablaBackPort & " WHERE " & txtcad
 rmesa.Open txtfiltro2, ConAdo
 rmesa.MoveFirst
 For i = 1 To noreg1
  mata(i, 1) = rmesa.Fields(0)
  mata(i, 2) = rmesa.Fields(5)
  mata(i, 3) = rmesa.Fields(6)
  mata(i, 4) = mata(i, 3) - mata(i, 2)
  rmesa.MoveNext
 Next i
 rmesa.Close
End If
txtcad = "(fecha >= " & txtfecha1 & " and fecha <= " & txtfecha2 & ") AND MPOSICION = '" & txtportBanobras & "' AND PORTAFOLIO = 'TOTALES' AND TVAR = 'CVAR VOL CONST P' AND NIV_CONF = .97 ORDER BY FECHA"
txtfiltro1 = "select count(*) from " & TablaResVaR & " WHERE " & txtcad
rmesa.Open txtfiltro1, ConAdo
noreg2 = rmesa.Fields(0)
rmesa.Close
If noreg2 <> 0 Then
   ReDim matb(1 To noreg2, 1 To 3) As Variant
   txtfiltro2 = "select * from " & TablaResVaR & " WHERE " & txtcad
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg2
       matb(i, 1) = rmesa.Fields(0)
       matb(i, 2) = rmesa.Fields(6)
       matb(i, 3) = rmesa.Fields(7)
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
contar = 0
For i = 1 To noreg
indice1 = BuscarValorArray(matback1(i, 1), mata, 1)
indice2 = BuscarValorArray(matback1(i, 1), matb, 1)
If indice1 <> 0 Then
 matback1(i, 2) = mata(indice1, 4)
End If
If indice2 <> 0 Then
 matback1(i, 3) = matb(indice2, 2)
End If
If matback1(i, 3) < matback1(i, 2) Then
matback1(i, 4) = 1
Else
matback1(i, 4) = 0
End If
contar = contar + matback1(i, 4)
Next i
probe = contar / noreg
nivconf = 0.97
estadis = -2 * Log((nivconf ^ contar * (1 - nivconf) ^ (noreg - contar)) / ((probe) ^ contar * (1 - probe) ^ (noreg - contar)))
End Sub

Function ObtenerHistEficiencia(ByVal noswap As String) As Variant()
Dim txtfiltro As String
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim rmesa As New ADODB.recordset

txtfiltro = "SELECT COUNT(*) from " & TablaEficienciaCob & " WHERE CLAVE_SWAP = '" & noswap & "'"
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
  txtfiltro = "SELECT * from " & TablaEficienciaCob & " WHERE CLAVE_SWAP = '" & noswap & "' ORDER BY FECHA2"
  rmesa.Open txtfiltro, ConAdo
  rmesa.MoveFirst
  ReDim mata(1 To noreg, 1 To 18) As Variant
  For i = 1 To noreg
   For j = 1 To 18
    mata(i, j) = rmesa.Fields(j - 1)
   Next j
   rmesa.MoveNext
  Next i
  rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
ObtenerHistEficiencia = mata
End Function

Function LeerVaR100MDResumen(ByVal fecha As Date, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal valposdiv As Double) As Variant()
Dim contar As Integer
Dim noreg As Integer
Dim suma As Double
Dim i As Integer
Dim mata() As Variant
Dim matb() As Variant
'exporta el VaR al 100% al archivo del resumen ejecutivo
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
noreg = UBound(MatGruposPortPos, 1)
suma = 0
contar = 0
ReDim matv(1 To 2, 1 To 1) As Variant
For i = 1 To noreg
    If MatGruposPortPos(i, 4) = 1 Then
       mata = LeerEscHistRE(fecha, txtportCalc1, "Normal", MatGruposPortPos(i, 3), noesc, 1)
       contar = contar + 1
       ReDim Preserve matv(1 To 2, 1 To contar + 1) As Variant
       matv(1, contar + 1) = MatGruposPortPos(i, 3)
       If Not EsArrayVacio(mata) Then
          matb = RutinaOrden(mata, 2, SRutOrden)
          suma = suma + matb(1, 2)
          matv(2, contar + 1) = matb(1, 2)
       Else
          matv(2, contar + 1) = 0
       End If
    End If
Next i
If valposdiv > 0 Then
   mata = LeerEscHistRE(fecha, txtportCalc1, "Normal", "MC POS DOLARES", noesc, 1)
   contar = contar + 1
   ReDim Preserve matv(1 To 2, 1 To contar + 1) As Variant
   matv(1, contar + 1) = "MC POS DOLARES"
   If Not EsArrayVacio(mata) Then
      matb = RutinaOrden(mata, 2, SRutOrden)
      suma = suma + matb(1, 2)
      matv(2, contar + 1) = matb(1, 2)
   Else
      matv(2, contar + 1) = 0
   End If
   mata = LeerEscHistRE(fecha, txtportCalc1, "Normal", "MC POS EUROS", noesc, 1)
   contar = contar + 1
   ReDim Preserve matv(1 To 2, 1 To contar + 1) As Variant
   matv(1, contar + 1) = "MC POS EUROS"
   If Not EsArrayVacio(mata) Then
      matb = RutinaOrden(mata, 2, SRutOrden)
      suma = suma + matb(1, 2)
      matv(2, contar + 1) = matb(1, 2)
   Else
      matv(2, contar + 1) = 0
   End If
   mata = LeerEscHistRE(fecha, txtportCalc1, "Normal", "MC POS YENES", noesc, 1)
   contar = contar + 1
   ReDim Preserve matv(1 To 2, 1 To contar + 1) As Variant
   matv(1, contar + 1) = "MC POS YENES"
   If Not EsArrayVacio(mata) Then
      matb = RutinaOrden(mata, 2, SRutOrden)
      suma = suma + matb(1, 2)
      matv(2, contar + 1) = matb(1, 2)
   Else
      matv(2, contar + 1) = 0
   End If
Else
   contar = contar + 1
   ReDim Preserve matv(1 To 2, 1 To contar + 1) As Variant
   contar = contar + 1
   ReDim Preserve matv(1 To 2, 1 To contar + 1) As Variant
   contar = contar + 1
   ReDim Preserve matv(1 To 2, 1 To contar + 1) As Variant

End If
matv(1, 1) = "VaR 100%"
matv(2, 1) = suma

mata = LeerEscHistRE(fecha, txtportCalc1, "Normal", txtportBanobras, noesc, 1)
contar = contar + 1
ReDim Preserve matv(1 To 2, 1 To contar + 2) As Variant
If Not EsArrayVacio(mata) Then
   matb = RutinaOrden(mata, 2, SRutOrden)
   matv(1, contar + 1) = txtportBanobras
   matv(2, contar + 1) = matb(1, 2)
   matv(1, contar + 2) = "FECHA ESC"
   matv(2, contar + 2) = CLng(matb(1, 1))
Else
   matv(1, contar + 1) = txtportBanobras
   matv(2, contar + 1) = 0
   matv(1, contar + 2) = "FECHA ESC"
   matv(2, contar + 2) = 0
End If
LeerVaR100MDResumen = MTranV(matv)
End Function

Sub GuardaVaR100MDResumen(ByVal fecha As Date, ByRef mata() As Variant, ByVal txttabla As String, conex, rbase)
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Integer
Dim noreg1 As Integer
Dim j As Integer

noreg = UBound(mata, 1)
If noreg > 0 Then
   txtfiltro = "SELECT COUNT(*) FROM [" & txttabla & "] WHERE FECHA = " & CLng(fecha)
   rbase.Open txtfiltro, conex
   noreg1 = rbase.Fields(0)
   rbase.Close
   If noreg1 <> 0 Then
  
   Else
      txtcadena = "INSERT INTO [" & txttabla & "] VALUES("
      txtcadena = txtcadena & CLng(fecha) & ","
      For j = 1 To noreg - 1
          txtcadena = txtcadena & Val(mata(j, 2)) & ","
      Next j
      txtcadena = txtcadena & Val(mata(noreg, 2)) & ")"
      conex.Execute txtcadena
   End If
End If
End Sub

Function GeneraCal(ByVal fecha As Date, ByVal tcalendar As String, ByVal dia As Integer, ByVal salto As Integer, comporta, inhabil, modificador, pais)
Dim noreg As Integer
Dim matf() As Date

'esta rutina debe de generar un calendario en funcion de los parametros introducidos
noreg = 100
If tcalendar = "DIAS" Then
   matf = GeneraCalDias(fecha, comporta, noreg, inhabil, pais)
ElseIf tcalendar = "FECHAS PERIODICAS DE CORTE" Then  'a partir de fecha de inicio
   If comporta = "UDM1" Then
      matf = GeneraCalUDM(fecha, salto, noreg, inhabil, pais)
   Else
      matf = GeneraCalPer(fecha, salto, comporta, noreg, inhabil, pais)
  End If
End If
GeneraCal = matf
End Function

Function GeneraCalDias(ByVal fecha As Date, ByVal comporta As Integer, ByVal noreg As Long, ByVal inhabil As String, ByVal pais As String) As Date()
Dim i As Long

     ReDim matf(1 To noreg, 1 To 2) As Date
      matf(i, 1) = FechaCalend(fecha, inhabil, "MX")
      For i = 1 To noreg
        matf(i, 2) = FechaCalend(matf(i, 1) + comporta, inhabil, "MX")
        If i <> noreg Then matf(i + 1, 1) = matf(i, 2)
      Next i
GeneraCalDias = matf
End Function

Function GeneraCalPer(ByVal fecha As Date, ByVal salto As Integer, ByVal comporta As Integer, ByVal noreg As Long, ByVal inhabil As String, ByVal pais As String) As Date()
Dim año As Integer
Dim mes As Integer
Dim i As Long

   ReDim matf(1 To noreg, 1 To 2) As Date
   matf(1, 1) = FechaCalend(fecha, inhabil, "MX")
   año = Year(matf(i, 1))
   mes = Month(matf(i, 1))
   For i = 1 To noreg
       mes = mes + salto
       matf(i, 2) = FechaCalend(DateSerial(año, mes, comporta), inhabil, "MX")
       If i <> noreg Then matf(i + 1, 1) = matf(i, 2)
   Next i
GeneraCalPer = matf
End Function

Function GeneraCalUDM(ByVal fecha As Date, ByVal salto As Integer, ByVal noreg As Long, ByVal inhabil As String, ByVal pais As String) As Date()
Dim año As Integer
Dim mes As Integer
Dim i As Long


'ultimo dia del mes, siguiente dia habil
   ReDim matf(1 To noreg, 1 To 2) As Date
   matf(1, 1) = FechaCalend(fecha, inhabil, pais)
   año = Year(matf(1, 1))
   mes = Month(matf(1, 1))
   For i = 1 To noreg
       matf(i, 2) = FechaCalend(UDM(año, mes), inhabil, pais)
       If i <> noreg Then matf(i + 1, 1) = matf(i, 2)
       mes = mes + salto
   Next i
GeneraCalUDM = matf
End Function

Function UDM(ByVal año As Integer, ByVal mes As Integer) As Date
'se obtiene el ultimo dia del mes
  UDM = DateSerial(año, mes + 1, 1) - 1
End Function

Function FechaCalend(ByVal fecha As Date, ByVal inhabil As String, ByVal pais As String) As Date
If inhabil = "INHABIL" Then
  FechaCalend = fecha
ElseIf inhabil = "SIGUIENTE HABIL" Then
  FechaCalend = FBD(fecha, pais)
ElseIf inhabil = "ANTERIOR HABIL" Then
  FechaCalend = PBD(fecha, pais)
End If
End Function

Sub CambiarCuadro(objeto, ByVal inicio1x As Long, ByVal inicio1y As Long, ByVal inicio2x As Long, ByVal inicio2y As Long, objeto1, objeto2)
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
objeto1.Left = inicio1x
objeto1.top = inicio1y
objeto2.Left = inicio2x
objeto2.top = inicio2y
objeto1.Width = Maximo(objeto.Width - 600, 0)
objeto1.Height = Maximo(objeto.Height - 1600, 0)
objeto2.Width = Maximo(objeto1.Width - 300, 0)
objeto2.Height = Maximo(objeto1.Height - 1200, 0)
On Error GoTo 0
Exit Sub
ControlErrores:
MsgBox error(Err())
On Error GoTo 0
End Sub

Function ObtenerProcesosPendientes(ByVal opcion As Integer)
Dim txtcadena As String
Dim txtfiltro As String
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim txttabla As String
Dim rmesa As New ADODB.recordset

If opcion = 1 Then
   txttabla = TablaProcesos1
ElseIf opcion = 2 Then
   txttabla = TablaProcesos2
End If

txtcadena = "SELECT * FROM " & txttabla & " WHERE FINALIZADO = 'N' AND BLOQUEADO = 'N' ORDER BY FECHAP, ID_TAREA"
txtfiltro = "SELECT COUNT(*) FROM (" & txtcadena & ")"
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtcadena, ConAdo
   nocampos = rmesa.Fields.Count
   rmesa.MoveFirst
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   For i = 1 To noreg
       For j = 1 To nocampos
           mata(i, j) = rmesa.Fields(j - 1)
       Next j
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
ObtenerProcesosPendientes = mata
End Function

Function ObtenerProcesosPendFecha(ByVal fecha As Date, ByVal opcion As Integer) As Variant()
Dim txtfecha As String
Dim txtcadena As String
Dim txtfiltro As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtcadena = " from " & DetermTablaProc(opcion) & " WHERE FECHAP = " & txtfecha & " AND REALIZO = 'N' ORDER BY FECHA, IDTAREA"
txtfiltro = "SELECT COUNT(*) " & txtcadena
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtfiltro = "SELECT * " & txtcadena
   rmesa.Open txtfiltro, ConAdo
   rmesa.MoveFirst
   ReDim mata(1 To noreg, 1 To 8) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields(0) 'FECHA
       mata(i, 2) = rmesa.Fields(1) 'CLAVE DE LA TAREA
       mata(i, 3) = rmesa.Fields(2) 'CONCEPTO
       mata(i, 4) = rmesa.Fields(3) 'FECHA
       mata(i, 5) = rmesa.Fields(4) 'HORA
       mata(i, 6) = rmesa.Fields(5) 'NO DE REGISTROS
       mata(i, 7) = rmesa.Fields(6) 'BLOQUEADA
       mata(i, 8) = rmesa.Fields(7) 'REALIZADA
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
ObtenerProcesosPendFecha = mata
End Function

Function EnBlackList(ByVal clave As String, ByRef mata() As Variant) As Boolean
Dim i As Long

'esta rutina revisa si una operacion esta en la lista negra
EnBlackList = False
For i = 1 To UBound(mata, 1)
  If clave = mata(i, 1) Then
     EnBlackList = True
     Exit Function
  End If
Next i
End Function

Function LeerValContraparte(ByVal fecha As Date, ByVal idval As Integer) As Variant()
    Dim contar As Integer
    Dim i As Integer
    Dim noreg As Integer
    Dim mata() As Integer
    Dim txtfecha As String
    Dim txtfiltro1 As String
    Dim txtfiltro2 As String
    Dim cmdSp As New ADODB.Command
    Dim recordset As New ADODB.recordset
    Dim rmesa As New ADODB.recordset
    
    With cmdSp
        .ActiveConnection = ConAdo
        .CommandType = adCmdStoredProc
        .CommandText = "OBTENERCONTRAPFECHA"
    End With
    cmdSp.Parameters.Append cmdSp.CreateParameter("FECHAX", adDBDate, adParamInput, , Format(fecha, "dd/mm/yyyy"))
    Set recordset = cmdSp.Execute
    contar = 0
    ReDim mata(1 To 1) As Integer
    While (Not recordset.EOF)
        contar = contar + 1
        ReDim Preserve mata(1 To contar) As Integer
        mata(contar) = recordset.Fields(0)
        recordset.MoveNext
    Wend
    recordset.Close
Set cmdSp = Nothing

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
ReDim matb(1 To contar, 1 To 6) As Variant
 
For i = 1 To contar
    matb(i, 1) = CLng(fecha) & "|" & mata(i)
    txtfiltro2 = "SELECT * FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND SUBPORT = 'CCS Contrap " & mata(i) & "'"
    txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = " & idval
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       matb(i, 2) = rmesa.Fields("MTM_SUCIO")
       rmesa.Close
    Else
       matb(i, 2) = 0
    End If
    txtfiltro2 = "SELECT * FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND SUBPORT = 'Fwds Contrap " & mata(i) & "'"
    txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = " & idval
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       matb(i, 3) = rmesa.Fields("MTM_SUCIO")
       rmesa.Close
    Else
       matb(i, 3) = 0
    End If
    txtfiltro2 = "SELECT * FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND SUBPORT = 'IRS Contrap " & mata(i) & "'"
    txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = " & idval
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       matb(i, 4) = rmesa.Fields("MTM_SUCIO")
       rmesa.Close
    Else
       matb(i, 4) = 0
    End If
    matb(i, 5) = matb(i, 2) + matb(i, 3) + matb(i, 4)
    txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = 'Deriv Contrap " & mata(i) & "'"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    matb(i, 6) = noreg
Next i
 LeerValContraparte = matb
End Function

Sub CopiaBaseOracle(ByVal tabla1 As String, ByVal tabla2 As String)
Dim nocampos As Integer
Dim txtfiltro As String
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim txtfecha As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

nocampos = 14
ReDim MATTIPO(1 To nocampos) As Variant
'1 texto
'2 numero
'3 fecha
MATTIPO(1) = 1
MATTIPO(2) = 3
MATTIPO(3) = 1
MATTIPO(4) = 1
MATTIPO(5) = 1
MATTIPO(6) = 1
MATTIPO(7) = 3
MATTIPO(8) = 3
MATTIPO(9) = 3
MATTIPO(10) = 1
MATTIPO(11) = 1
MATTIPO(12) = 2
MATTIPO(13) = 2
MATTIPO(14) = 2

txtfiltro = "SELECT COUNT(*) FROM " & tabla1
rmesa.Open txtfiltro, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtfiltro = "SELECT * FROM " & tabla1
   rmesa.Open txtfiltro, ConAdo
   rmesa.MoveFirst
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
   For i = 1 To noreg
       For j = 1 To nocampos
           mata(i, j) = rmesa.Fields(j - 1)
       Next j
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo la tabla " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   rmesa.Close
   conAdo2.Execute "DELETE FROM " & tabla2
   For i = 1 To noreg
       txtcadena = "INSERT INTO " & tabla2 & " VALUES("
       For j = 1 To nocampos
           If j <> nocampos Then
              If MATTIPO(j) = 1 Then
                 txtcadena = txtcadena & "'" & mata(i, j) & "',"
              ElseIf MATTIPO(j) = 2 Then
                 txtcadena = txtcadena & mata(i, j) & ","
              ElseIf MATTIPO(j) = 3 Then
                 txtfecha = "to_date('" & Format(mata(i, j), "dd/mm/yyyy") & "','dd/mm/yyyy')"
                 txtcadena = txtcadena & txtfecha & ","
              End If
           Else
              If MATTIPO(j) = 1 Then
                 txtcadena = txtcadena & "'" & mata(i, j) & "')"
              ElseIf MATTIPO(j) = 2 Then
                 txtcadena = txtcadena & mata(i, j) & ")"
              ElseIf MATTIPO(j) = 3 Then
                 txtfecha = "to_date('" & Format(mata(i, j), "dd/mm/yyyy") & "','dd/mm/yyyy')"
                 txtcadena = txtcadena & txtfecha & ")"
              End If
           End If
       Next j
       conAdo2.Execute txtcadena
       AvanceProc = i / noreg
       MensajeProc = "Exportando la tabla " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
End If
End Sub

Function StOR(ByVal Y As Boolean, ByVal x As Boolean) As Boolean
If x = True Or Y Then
   StOR = True
Else
   StOR = False
End If

End Function

Function StAND(ByVal Y As Boolean, ByVal x As Boolean) As Boolean
If x = True And Y Then
   StAND = True
Else
   StAND = False
End If
End Function

Function DetFechasCalculo(ByVal fecha As Date, ByVal fecha2 As Date, ByVal nomes As Integer, ByRef matpos() As propPosRiesgo, ByRef matposswaps() As propPosSwaps, matposfwds() As propPosFwd, ByVal opcion As Integer) As Date()
Dim i As Integer
Dim fechax As Date
Dim contar As Integer
Dim indice As Integer
Dim ffinal As Date
If opcion = 1 Then
   For i = 1 To UBound(matpos, 1)
       indice = matpos(i).IndPosicion
       If i = 1 Then
          If matpos(i).No_tabla = 3 Then
             ffinal = matposswaps(indice).FvencSwap
          ElseIf matpos(i).No_tabla = 4 Then
             ffinal = matposfwds(indice).FVencFwd
          End If
       Else
          If matpos(i).No_tabla = 3 Then
             ffinal = Maximo(ffinal, matposswaps(indice).FvencSwap)
          ElseIf matpos(i).No_tabla = 4 Then
             ffinal = Maximo(ffinal, matposfwds(indice).FVencFwd)
          End If
       End If
   Next i
Else
   ffinal = fecha2
End If
ReDim matf(1 To 1) As Date
fechax = fecha
contar = 0
Do While fechax <= ffinal
   contar = contar + 1
   ReDim Preserve matf(1 To contar) As Date
   fechax = DateSerial(Year(fechax), Month(fechax) + nomes, Day(fechax))
   matf(contar) = fechax
Loop
DetFechasCalculo = matf
End Function

Function DetFechasCalcM(ByVal fecha As Date, ByVal nomes As Integer, ByRef matpos() As propPosRiesgo, ByRef matposswaps() As propPosSwaps, ByRef matposfwds() As propPosFwd, ByRef matposdeuda() As propPosDeuda)
Dim i As Long
Dim contar As Long
Dim indice As Long
Dim fvenc As Date
Dim fechax As Date
Dim f_pos As Date

fvenc = 0
For i = 1 To UBound(matpos, 1)
    indice = matpos(i).IndPosicion
    If matpos(i).No_tabla = 3 Then 'swap
       fvenc = Maximo(matposswaps(indice).FvencSwap, fvenc)
    ElseIf matpos(i).No_tabla = 4 Then 'fwd
       fvenc = Maximo(matposfwds(indice).FVencFwd, fvenc)
    ElseIf matpos(i).No_tabla = 5 Then 'DEUDA
       fvenc = Maximo(matposdeuda(indice).FVencDeuda, fvenc)
    End If
Next i
'primero se determina el primer dia del mes siguiente
fechax = DateSerial(Year(fecha), Month(fecha) + 1, 1)
f_pos = fecha
ReDim matf(1 To 1) As Date
contar = 0
Do While f_pos <= fvenc
   fechax = DateSerial(Year(fechax), Month(fechax) + nomes, 1)
   f_pos = fechax - 1
   contar = contar + 1
   ReDim Preserve matf(1 To contar) As Date
   matf(contar) = f_pos
Loop
DetFechasCalcM = matf
End Function

Function DetFechasCalcM2(ByVal fecha As Date, ByVal nodias As Integer, ByRef matpos() As propPosRiesgo, ByRef matposswaps() As propPosSwaps, ByRef matposfwds() As propPosFwd, ByRef matposdeuda() As propPosDeuda)
Dim i As Long
Dim contar As Long
Dim indice As Long
Dim fvenc As Date
Dim fechax As Date

fvenc = 0
For i = 1 To UBound(matpos, 1)
    indice = matpos(i).IndPosicion
    If matpos(i).No_tabla = 3 Then 'swap
       fvenc = Maximo(matposswaps(indice).FvencSwap, fvenc)
    ElseIf matpos(i).No_tabla = 4 Then 'fwd
       fvenc = Maximo(matposfwds(indice).FVencFwd, fvenc)
    ElseIf matpos(i).No_tabla = 5 Then 'DEUDA
       fvenc = Maximo(matposdeuda(indice).FVencDeuda, fvenc)
    End If
Next i
'primero se determina el primer dia del mes siguiente
fechax = fecha
ReDim matf(1 To 1) As Date
contar = 0
Do While fechax <= fvenc
   fechax = fechax + nodias
   contar = contar + 1
   ReDim Preserve matf(1 To contar) As Date
   matf(contar) = fechax
Loop
DetFechasCalcM2 = matf
End Function

Function CargarMatTrans(ByVal fecha As Date, ByVal escala As String) As Double()
Dim mata() As Variant
Dim noreg As Long
Dim noreg1 As Long
Dim i As Long

mata = LeerMatTrans(fecha, escala)
noreg = UBound(mata, 1)
If noreg <> 0 Then
noreg1 = noreg ^ 0.5
ReDim matb(1 To noreg1, 1 To noreg1) As Double
For i = 1 To noreg
   matb(mata(i, 1), mata(i, 2)) = mata(i, 3)
Next i
Else
ReDim matb(0 To 0, 0 To 0) As Double
End If
CargarMatTrans = matb
End Function

Function LeerMatTrans(ByVal fecha As Date, ByVal escala As String) As Variant()
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & PrefijoBD & TablaMTrans & " WHERE"
txtfiltro2 = txtfiltro2 & " REGION = '" & escala & "'"
txtfiltro2 = txtfiltro2 & " AND FECHA IN(SELECT MAX(FECHA) AS FECHA FROM " & PrefijoBD & TablaMTrans
txtfiltro2 = txtfiltro2 & " WHERE REGION = '" & escala & "'"
txtfiltro2 = txtfiltro2 & " AND FECHA <= " & txtfecha & ")"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 3) As Variant
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields(2)
       mata(i, 2) = rmesa.Fields(3)
       mata(i, 3) = rmesa.Fields(4)
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerMatTrans = mata
End Function

Function LeerMatTrans2(ByVal fecha As Date, ByVal escala As String) As Variant()
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfiltro0 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT * FROM " & PrefijoBD & TablaMTrans & " WHERE FECHA =  " & txtfecha & " AND REGION = '" & escala & "'"
txtfiltro0 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro0, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   ReDim mata(1 To noreg, 1 To 3) As Variant
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields(2)
       mata(i, 2) = rmesa.Fields(3)
       mata(i, 3) = rmesa.Fields(4)
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerMatTrans2 = mata
End Function

Function CalcProbDefault(ByVal calif As Integer, ByRef mattran1() As Double, ByVal noper As Integer) As Double()
Dim mata() As Double
Dim matb() As Double
Dim mattran11() As Double
Dim noreg As Long
Dim noaños As Long
Dim i As Long


mattran11 = MMult(mattran1, mattran1)    'elevo la matriz al cuadrado
noreg = UBound(mattran11, 1)             'no de calificaciones
noaños = Int((noper - 1) / 4) + 1        'no de veces que se eleva la matriz
mata = MIdentidad(noreg)
matb = MIdentidad(noreg)
ReDim matper(1 To noper) As Double
For i = 1 To noper
    If ((i - 1) Mod 4) = 0 Then          'si el periodo de calculo es multiplo de el año se calcule
       mata = matb
       If i = 1 Then
          matb = mattran11
       Else
          matb = MMult(mata, mattran1)
       End If
    End If
    If Int((i - 1) / 4) + 1 = 1 Then
       matper(i) = matb(calif, noreg) / 4
    Else
       matper(i) = ((matb(calif, noreg) - mata(calif, noreg)) / 4)
    End If
Next i
CalcProbDefault = matper
End Function

Function CalcProbDefault1(ByVal calif As Integer, ByRef mattran1() As Double, ByVal noper As Integer) As Double()
Dim mata() As Double
Dim matb() As Double
Dim noreg As Long
Dim noaños As Long
Dim i As Long
Dim mattran11() As Double

mattran11 = MMult(mattran1, mattran1)
noreg = UBound(mattran11, 1)
noaños = Int((noper - 1) / 4) + 1
mata = MIdentidad(noreg)
matb = MIdentidad(noreg)
ReDim matper(1 To noper) As Double
For i = 1 To noper
    If ((i - 1) Mod 4) = 0 Then
       mata = matb
       If i = 1 Then
          matb = mattran11
       Else
          matb = MMult(mata, mattran1)
       End If
    End If
    matper(i) = matb(calif, noreg)
Next i
CalcProbDefault1 = matper
End Function

Sub RutinaCargaFR(fecha As Date, ByRef exito As Boolean)
Dim indice As Long
Dim exito1 As Boolean
 'se busca la fecha en el indice de fechas con factores de riesgo para var
 indice = BuscarValorArray(fecha, MatFechasFR, 1)
 'si se encuentra se procede a realizar la valuacion
 exito = True
 If indice <> 0 Then
   'Se cargan los factores del dia actual en funcion de un catalogo de factores
    If fecha <> fechaFactoresR Or EsArrayVacio(MatFactR1) Then
       MatFactR1 = CargaFR1Dia(fecha, exito1)
       fechaFactoresR = fecha
    Else
       exito1 = True
    End If
    exito = exito And exito1
 Else
  MsgBox "no hay tasas para este día"
  exito = False
 End If
End Sub

Sub NuevaCFriesgo(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtcadena As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim noreg1 As Long
Dim i As Long
Dim j As Long
Dim r As Long
Dim matp() As Variant
Dim matx() As Variant
Dim exito As Boolean
Dim contar As Long
Dim valind As Double
Dim indice As Long
Dim rmesa As New ADODB.recordset

txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtcadena = TablaFechasVaR & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2
txtfiltro1 = "select count(*) from " & txtcadena
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
     ReDim MatFactRiesgo(1 To noreg, 1 To NoFactores + 1) As Variant
     txtfiltro2 = "select FECHA from " & txtcadena & " ORDER BY FECHA"
     rmesa.Open txtfiltro2, ConAdo
     rmesa.MoveFirst
     For i = 1 To noreg
         MatFactRiesgo(i, 1) = rmesa.Fields("FECHA")
         rmesa.MoveNext
         MensajeProc = "Leyendo los factores de riesgo " & Format(AvanceProc, "#,##0.00 %")
         Call MostrarMensajeSistema(MensajeProc, frmProgreso.Label2, 0, Date, Time, NomUsuario)
     Next i
     rmesa.Close
     NoFactores = UBound(MatCaracFRiesgo, 1)
     For i = 1 To UBound(MatResFRiesgo, 1)
         matp = DetPlazosCurva(MatResFRiesgo(i, 1), MatCaracFRiesgo)
         matx = LeerNodosCurvaO(fecha1, fecha2, MatResFRiesgo(i, 1), matp, exito)
         If UBound(matx, 1) > 0 Then
         For j = 1 To MatResFRiesgo(i, 2)
             If i <> 1 Then
                indice = MatResFRiesgo(i - 1, 3) + j + 1
             Else
                indice = j + 1
             End If
             For r = 1 To noreg
                MatFactRiesgo(r, indice) = matx(r, j + 1)
             Next r
         Next j
         End If
     Next i
'el array de indexado de datos
ReDim matind(1 To NoFactores * noreg, 1 To 3) As Variant
contar = 1
For i = 1 To noreg
   For j = 1 To NoFactores
       matind(contar, 1) = CLng(MatFactRiesgo(i, 1)) & MatCaracFRiesgo(j).nomFactor & Format(MatCaracFRiesgo(j).plazo, "0000000")
       matind(contar, 2) = i
       matind(contar, 3) = j
       contar = contar + 1
   Next j
Next i
matind = RutinaOrden(matind, 1, 1)
'por ultimo solo se leen los factores tipo indice o tipo de cambio
For i = 1 To NoFactores
    If MatCaracFRiesgo(i).plazo = 0 Then
    txtfiltro1 = "select COUNT(*) from " & TablaFRiesgoO & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & " AND CONCEPTO = '" & MatCaracFRiesgo(i).nomFactor & "' AND PLAZO = " & MatCaracFRiesgo(i).plazo & " ORDER BY FECHA"
    rmesa.Open txtfiltro1, ConAdo
    noreg1 = rmesa.Fields(0)
    rmesa.Close
    If noreg1 <> noreg Then
      MensajeProc = "Faltan registros para el factor " & MatCaracFRiesgo(i).nomFactor & " " & MatCaracFRiesgo(i).plazo
      'MsgBox MensajeProc
      'Call MostrarMensajeSistema(MensajeProc, frmprogreso.label2, 2, Date, Time, NomUsuario)
    End If
    If noreg1 <> 0 Then
       txtfiltro1 = "select FECHA, VALOR,INDICE from " & TablaFRiesgoO & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & " AND CONCEPTO = '" & MatCaracFRiesgo(i).nomFactor & "' AND PLAZO = " & MatCaracFRiesgo(i).plazo & " ORDER BY FECHA"
       rmesa.Open txtfiltro1, ConAdo
       rmesa.MoveFirst
       For j = 1 To noreg1
           valind = rmesa.Fields(2)
           indice = BuscarValorArray(valind, matind, 1)
           If indice <> 0 Then
             MatFactRiesgo(matind(indice, 2), matind(indice, 3) + 1) = rmesa.Fields(1)
           End If
           rmesa.MoveNext
       Next j
       rmesa.Close
    Else
       MensajeProc = "No hay datos para el factor " & MatCaracFRiesgo(i).nomFactor & " " & MatCaracFRiesgo(i).plazo
       MsgBox MensajeProc
    End If
    End If
     AvanceProc = i / NoFactores
    MensajeProc = "Leyendo los factores de riesgo " & MatCaracFRiesgo(i).nomFactor & " " & MatCaracFRiesgo(i).plazo & " " & Format(i / NoFactores, "#,##0.00 %")
    Call MostrarMensajeSistema(MensajeProc, frmProgreso.Label2, 0, Date, Time, NomUsuario)
    DoEvents
Next i
End If
End Sub

Function DetPlazosCurva(ByVal txtcurva As String, ByRef mata() As propNodosFRiesgo) As Variant()
Dim noreg As Long
Dim i As Long
Dim contar As Long

noreg = UBound(mata, 1)
ReDim matc(1 To 1) As Variant
contar = 0
For i = 1 To noreg
    If mata(i, 3) = txtcurva Then
       contar = contar + 1
       ReDim Preserve matc(1 To contar) As Variant
       matc(contar) = mata(i).plazo
    End If
Next i
DetPlazosCurva = matc
End Function

Sub ExpCaracFlujosSwaps(ByVal fecha As Date)
Dim matpos() As propPosSwaps
Dim flujos() As estFlujosDeuda
Dim salida As String
Dim txtsalida As String
Dim i As Long
Dim txtfiltro As String
Dim txtcontrap As String
Dim txtfecha As String
Dim txtport As String
Dim exito As Boolean
Dim exitoarch As Boolean

txtport = "SWAPS"
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro = "SELECT * FROM " & TablaPosSwaps
txtfiltro = txtfiltro & " WHERE (TIPOPOS,CPOSICION,FECHAREG,COPERACION) IN ("
txtfiltro = txtfiltro & "SELECT TIPOPOS,CPOSICION,FECHAREG,COPERACION FROM " & TablaPortPosicion & " "
txtfiltro = txtfiltro & " WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtport & "')"
     Call CrearPosSwaps(txtfiltro, matpos, flujos, exito)
     If UBound(matpos, 1) <> 0 Then
     salida = DirResVaR & "\Caract y flujos swaps " & Format(fecha, "yyyy-mm-dd") & ".txt"
     frmCalVar.CommonDialog1.FileName = salida
     frmCalVar.CommonDialog1.ShowSave
     salida = frmCalVar.CommonDialog1.FileName
     Call VerificarSalidaArchivo(salida, 1, exitoarch)
     If exitoarch Then
     txtsalida = "Clave de operación" & Chr(9) & "Intención" & Chr(9) & "Estructural" & Chr(9) & "Fecha inicio" & Chr(9) & "Fecha vencimiento" & Chr(9)
     txtsalida = txtsalida & "Intercambio Inicial flujos" & Chr(9) & "Intercambio Final flujos" & Chr(9) & "Reinvierte intereses activa" & Chr(9)
     txtsalida = txtsalida & "Reinvierte intereses pasiva" & Chr(9) & "Tasa cupon activa" & Chr(9) & "Tasa cupon pasiva" & Chr(9) & "Sobretasa activa" & Chr(9)
     txtsalida = txtsalida & "Sobretasa pasiva" & Chr(9) & "Concencion int activa" & Chr(9) & "Convencion int pasiva" & Chr(9) & "Clave tipo swap" & Chr(9) & "Tipo de swap" & Chr(9)
     txtsalida = txtsalida & "Clave de la Contraparte"
     Print #1, txtsalida
     For i = 1 To UBound(matpos, 1)
         txtsalida = ""
         txtsalida = txtsalida & matpos(i).c_operacion & Chr(9)          'clave de la operacion
         txtsalida = txtsalida & matpos(i).intencion & Chr(9)            'intencion
         txtsalida = txtsalida & matpos(i).EstructuralSwap & Chr(9)      'estructural
         txtsalida = txtsalida & matpos(i).FCompraSwap & Chr(9)          'fecha de inicio del swap
         txtsalida = txtsalida & matpos(i).FvencSwap & Chr(9)            'fecha de vencimiento del swap
         txtsalida = txtsalida & matpos(i).IntercIFSwap & Chr(9)         'intercambio inicial de nocionales
         txtsalida = txtsalida & matpos(i).IntercFFSwap & Chr(9)
         txtsalida = txtsalida & matpos(i).RIntAct & Chr(9)
         txtsalida = txtsalida & matpos(i).RIntPas & Chr(9)
         txtsalida = txtsalida & matpos(i).TCActivaSwap & Chr(9)         'tasa cupon activa
         txtsalida = txtsalida & matpos(i).TCPasivaSwap & Chr(9)         'tasa cupon pasiva
         txtsalida = txtsalida & matpos(i).STActiva & Chr(9)
         txtsalida = txtsalida & matpos(i).STPasiva & Chr(9)
         txtsalida = txtsalida & matpos(i).ConvIntAct & Chr(9)
         txtsalida = txtsalida & matpos(i).ConvIntPas & Chr(9)
         txtsalida = txtsalida & matpos(i).ClaveProdSwap & Chr(9)         'clave de tipo de swap
         txtsalida = txtsalida & matpos(i).cProdSwapGen & Chr(9)          'clave de tipo de swap
         txtsalida = txtsalida & matpos(i).c_em_pidv & Chr(9)          'clave de tipo de swap
         txtcontrap = DeterminaContraparte(matpos(i).ID_ContrapSwap)
         txtsalida = txtsalida & txtcontrap                            'clave de contraparte
         Print #1, txtsalida
     Next i
     Print #1, ""
     txtsalida = "Clave de la operacion" & Chr(9) & "Activa o pasiva" & Chr(9) & "Fecha inicio flujo" & Chr(9) & "Fecha fin flujo" & Chr(9) & "Fecha descuento flujo" & Chr(9)
     txtsalida = txtsalida & "Paga intereses en el periodo" & Chr(9) & "Calc Intereses sobre total saldo" & Chr(9) & "Saldo en el periodo" & Chr(9)
     txtsalida = txtsalida & "Amortizacion" & Chr(9) & "Tasa cupon"
     Print #1, txtsalida
     For i = 1 To UBound(flujos, 1)
         txtsalida = ""
         txtsalida = txtsalida & flujos(i).coperacion & Chr(9)   'clave de la operacion
         txtsalida = txtsalida & flujos(i).tpata & Chr(9)   'posicion activa o pasiva
         txtsalida = txtsalida & flujos(i).finicio & Chr(9)   'fecha de inicio del flujo
         txtsalida = txtsalida & flujos(i).ffin & Chr(9)   'fecha final del flujo
         txtsalida = txtsalida & flujos(i).fpago & Chr(9)   'fecha de descuento del flujo
         txtsalida = txtsalida & flujos(i).pago_int & Chr(9)   'paga intereses en el periodo
         txtsalida = txtsalida & flujos(i).int_t_saldo & Chr(9)   'intereses sobre saldo
         txtsalida = txtsalida & flujos(i).saldo & Chr(9)   'saldo el el periodo
         txtsalida = txtsalida & flujos(i).amort & Chr(9)   'amortizacion
         txtsalida = txtsalida & flujos(i).t_cupon & Chr(9)  'tasa de interes
         Print #1, txtsalida
     Next i
     Close #1
     MsgBox "Se genero el archivo " & salida
     End If
  End If
 
End Sub

Sub ExpCaracFlujosDeuda(ByVal fecha As Date)
Dim matpos() As propPosDeuda
Dim flujos() As estFlujosDeuda
Dim txtfiltro As String
Dim salida As String
Dim txtsalida As String
Dim i As Long
Dim exito As Boolean
Dim txtfecha As String
Dim exitoarch As Boolean

     txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfiltro = "SELECT * FROM " & TablaPosDeuda & " WHERE TIPOPOS = 1"
     txtfiltro = txtfiltro & " AND (FECHAREG,COPERACION) IN "
     txtfiltro = txtfiltro & "(select MAX(FECHAREG) AS FECHAREG,COPERACION FROM " & TablaPosDeuda
     txtfiltro = txtfiltro & " WHERE FECHAREG <=" & txtfecha
     txtfiltro = txtfiltro & " AND TIPOPOS = 1 GROUP BY COPERACION)"
     txtfiltro = txtfiltro & " AND FVENCIMIENTO > " & txtfecha
     Call LeerPosDeuda(txtfiltro, matpos, flujos, exito)
     salida = DirResVaR & "\Caract y flujos deuda " & Format(fecha, "yyyy-mm-dd") & ".txt"
     frmCalVar.CommonDialog1.FileName = salida
     frmCalVar.CommonDialog1.ShowSave
     salida = frmCalVar.CommonDialog1.FileName
     Call VerificarSalidaArchivo(salida, 1, exitoarch)
     If exitoarch Then
     txtsalida = "Intencion" & Chr(9) & "Clave Operacion" & Chr(9) & "Fecha inicio" & Chr(9) & "Fecha vencimiento" & Chr(9)
     txtsalida = txtsalida & "Intercambio Inicial flujos" & Chr(9) & "Intercambio Final flujos" & Chr(9) & "Reinvierte intereses" & Chr(9)
     txtsalida = txtsalida & "Tasa cupon" & Chr(9) & "Sobretasa" & Chr(9)
     txtsalida = txtsalida & "Convencion int" & Chr(9) & "Clave tipo producto"
     Print #1, txtsalida
     For i = 1 To UBound(matpos, 1)
         txtsalida = ""
         txtsalida = txtsalida & matpos(i).c_operacion & Chr(9)         'clave de la operacion
         txtsalida = txtsalida & matpos(i).intencion & Chr(9)           'intencion
         txtsalida = txtsalida & matpos(i).FinicioDeuda & Chr(9)        'fecha de inicio del swap
         txtsalida = txtsalida & matpos(i).FVencDeuda & Chr(9)          'fecha de vencimiento del swap
         txtsalida = txtsalida & matpos(i).InteriDeuda & Chr(9)         'intercambio inicial de intereses
         txtsalida = txtsalida & matpos(i).InterfDeuda & Chr(9)         'intercambio intermedio y final de intereses
         txtsalida = txtsalida & matpos(i).RintDeuda & Chr(9)           'reinvierte intereses
         txtsalida = txtsalida & matpos(i).TRefDeuda & Chr(9)           'tasa cupon
         txtsalida = txtsalida & matpos(i).SpreadDeuda & Chr(9)         'sobretasa cupon
         txtsalida = txtsalida & matpos(i).ConvIntDeuda & Chr(9)        'sobretasa cupon
         txtsalida = txtsalida & matpos(i).ProductoDeuda & Chr(9)      'clave de producto
         Print #1, txtsalida
     Next i
     Print #1, ""
     txtsalida = "Clave de la operacion" & Chr(9) & "Fecha inicio flujo" & Chr(9) & "Fecha fin flujo" & Chr(9) & "Fecha descuento flujo" & Chr(9)
     txtsalida = txtsalida & "Paga intereses en el periodo" & Chr(9) & "Calc Intereses sobre t saldo" & Chr(9) & "Saldo en el periodo" & Chr(9)
     txtsalida = txtsalida & "Amortizacion" & Chr(9) & "Tasa cupon"
     Print #1, txtsalida
     For i = 1 To UBound(flujos, 1)
         txtsalida = ""
         txtsalida = txtsalida & flujos(i, 1) & Chr(9)   'clave de la operacion
         txtsalida = txtsalida & flujos(i, 2) & Chr(9)   'fecha de inicio del flujo
         txtsalida = txtsalida & flujos(i, 3) & Chr(9)   'fecha final del flujo
         txtsalida = txtsalida & flujos(i, 4) & Chr(9)   'fecha de descuento del flujo
         txtsalida = txtsalida & flujos(i, 5) & Chr(9)   'paga intereses en el periodo
         txtsalida = txtsalida & flujos(i, 6) & Chr(9)   'intereses sobre saldo
         txtsalida = txtsalida & flujos(i, 7) & Chr(9)   'saldo el el periodo
         txtsalida = txtsalida & flujos(i, 8) & Chr(9)   'amortizacion
         txtsalida = txtsalida & flujos(i, 9) & Chr(9)   'tasa de interes
         Print #1, txtsalida
     Next i
     Close #1
     MsgBox "Se genero el archivo " & salida
     MsgBox "Fin de proceso"
     End If
End Sub

Sub GenerarParamUsuario(ByVal usuario As String)
Dim noreg As Long
Dim i As Long
Dim txtinserta As String

ConAdo.Execute "DELETE FROM " & TablaParamUsuario & " WHERE USUARIO = '" & usuario & "'"
noreg = UBound(MatParamSistema, 1)
For i = 1 To noreg
    txtinserta = "INSERT INTO " & TablaParamUsuario & " VALUES("
    txtinserta = txtinserta & "'" & usuario & "',"
    txtinserta = txtinserta & "'" & MatParamSistema(i, 2) & "',"
    txtinserta = txtinserta & "'" & MatParamSistema(i, 3) & "')"
    ConAdo.Execute txtinserta
Next i
End Sub

Sub BloquearSubProc(ByVal idfolio As Long, ByVal opcion As Integer, ByRef exito As Boolean)
On Error GoTo hayerror
    Dim txtfiltro As String
    Dim txtcadena As String
    Dim txtipdir As String
    Dim finicio  As String
    Dim hinicio  As String
    Dim noreg    As Long
    
    finicio = "TO_DATE('" & Format$(Date, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    hinicio = "TO_DATE('" & Format$(Time, "HH:MM:SS") & "','HH24:MI:SS')"
    txtipdir = RecuperarIP
    txtcadena = "UPDATE " & DetermTablaSubproc(opcion) & " SET BLOQUEADO = 'S', USUARIO = '" & NomUsuario & "', FINICIO = " & finicio & ", HINICIO = " & hinicio & ", IP_DIRECCION = '" & txtipdir & "' WHERE FOLIO = " & idfolio & " AND BLOQUEADO = 'N' AND FINALIZADO = 'N'"
    ConAdo.Execute txtcadena, noreg
    If noreg <> 0 Then
       exito = True
    Else
       exito = False
    End If
On Error GoTo 0
Exit Sub
hayerror:
   exito = False
End Sub

Sub DesbloquearSubProc(ByVal idfolio As Long, ByVal opcion As Integer, ByVal txtmsg As String, ByVal final As Boolean, ByVal exito As Boolean)
    Dim ffinal As String
    Dim hfinal As String
    Dim txtcadena As String
    Dim txttabla As String
    Dim noreg As Long
    ffinal = "TO_DATE('" & Format$(Date, "DD/MM/YYYY") & "','DD/MM/YYYY')"
    hfinal = "TO_DATE('" & Format$(Time, "HH:MM:SS") & "','HH24:MI:SS')"
    txtcadena = "UPDATE " & DetermTablaSubproc(opcion) & " SET COMENTARIO = '" & Trim(Left(txtmsg, 200)) & "', BLOQUEADO = 'N', FINALIZADO = '" & ConvBolStr(final) & "', EXITO = '" & ConvBolStr(exito) & "', HFINAL = " & hfinal & ", FFINAL =  " & ffinal & "  WHERE FOLIO = " & idfolio
    ConAdo.Execute txtcadena, noreg
End Sub

Function ObSubProcPend(ByRef bl_exito As Boolean, ByVal opcion As Integer) As Variant()
Dim txtfecha1 As String, txtfiltro As String, txtfiltro1 As String
Dim i As Long
Dim noreg As Long
Dim nocampos As Integer
Dim rmesa As New ADODB.recordset

On Error GoTo hayerror

    txtfiltro = "SELECT * FROM (SELECT * FROM " & DetermTablaSubproc(opcion) & " WHERE BLOQUEADO = 'N' AND FINALIZADO = 'N' ORDER BY FECHAP,ID_SUBPROCESO,FOLIO) WHERE ROWNUM <= 1"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg = rmesa.Fields(0)
    rmesa.Close
    If noreg <> 0 Then
        rmesa.Open txtfiltro, ConAdo
        nocampos = rmesa.Fields.Count
        ReDim mata(1 To nocampos) As Variant
        For i = 1 To nocampos
            If Not EsVariableVacia(rmesa.Fields(i - 1)) Then
                mata(i) = rmesa.Fields(i - 1)
            Else
                mata(i) = ""
            End If
            AvanceProc = i / 1
            MensajeProc = "Obteniendo subprocesos pendientes " & Format$(AvanceProc, "##0.00 %")
            DoEvents
        Next i
        rmesa.Close
        bl_exito = True
    Else
        ReDim mata(0 To 0) As Variant
        bl_exito = False
    End If
    ObSubProcPend = mata
Exit Function
hayerror:
MsgBox error(Err())
End Function

Function LeyendoEstadoUsuario(ByVal txtnomusuario As String) As Double
Dim txtfiltro2 As String
Dim txtfiltro1 As String
Dim txtfiltro As String
Dim noreg As Integer
Dim ufecha As Double
Dim uhora As Double
Dim pflota As Double
Dim utiempo As Double

txtfiltro2 = "SELECT * FROM " & TablaUsuarios & " WHERE USUARIO = '" & txtnomusuario & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
REstadoUsuario.Open txtfiltro1, ConAdo
noreg = REstadoUsuario.Fields(0)
REstadoUsuario.Close
If noreg <> 0 Then
   REstadoUsuario.Open txtfiltro2, ConAdo
   ufecha = CDbl(REstadoUsuario.Fields("FUREPORTE"))
   uhora = REstadoUsuario.Fields("HUREPORTE")
   REstadoUsuario.Close
   pflota = CDbl(CDbl(uhora) - Int(uhora))
   utiempo = ufecha + pflota
End If
LeyendoEstadoUsuario = utiempo
End Function

Sub GuardarValOper(ByVal f_pos As Date, _
                   ByVal f_factor As Date, _
                   ByVal f_val As Date, _
                   ByVal txtport As String, _
                   ByVal txtportfr As String, _
                   ByRef matpos() As propPosRiesgo, _
                   ByRef matposmd() As propPosMD, _
                   ByRef matpr() As resValIns, _
                   ByVal tipopos As Integer, _
                   ByVal fechareg As Date, _
                   ByVal txtnompos As String, _
                   ByVal cposicion As Integer, _
                   ByVal coperacion As String, _
                   ByVal tval As Integer, _
                   ByRef exito As Boolean)

Dim txtcadena As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfechap As String
Dim txtfechafr As String
Dim txtfechav As String
Dim i As Long
Dim indice As Long
Dim txtborra As String
Dim txtfechar As String
Dim matvalik() As Variant
Dim opcion As Boolean
opcion = False

txtfechap = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechafr = "to_date('" & Format(f_factor, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechav = "to_date('" & Format(f_val, "dd/mm/yyyy") & "','dd/mm/yyyy')"
'If opcion Then
'   txtborra = "DELETE FROM " & TablaValPos & " WHERE FECHAP = " & txtfechap
'   txtborra = txtborra & " AND FECHAFR = " & txtfechafr
'   txtborra = txtborra & " AND FECHAV = " & txtfechav
'   txtborra = txtborra & " AND ID_VALUACION = " & tval & " AND PORTAFOLIO = '" & txtport & "'"
'   conAdo.Execute txtborra
'End If
If UBound(matpos, 1) <> 0 Then
For i = 1 To UBound(matpos, 1)
    txtfechar = "to_date('" & Format(matpos(i).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "INSERT INTO " & TablaValPos & " VALUES("
    txtcadena = txtcadena & txtfechap & ","                             'fecha de la posicion
    txtcadena = txtcadena & txtfechafr & ","                            'fecha de los factores de riesgo
    txtcadena = txtcadena & txtfechav & ","                             'fecha de valuacion
    txtcadena = txtcadena & tval & ","                                  'tipo de valuacion
    txtcadena = txtcadena & "'" & txtport & "',"                        'PORTAFOLIO
    txtcadena = txtcadena & "'" & txtportfr & "',"                      'escenario de factores de riesgo
    txtcadena = txtcadena & matpos(i).C_Posicion & ","                  'CLAVE DE LA POSICION
    txtcadena = txtcadena & txtfechar & ","                             'fecha de registro
    txtcadena = txtcadena & "'" & matpos(i).c_operacion & "',"          'clave de operacion
    txtcadena = txtcadena & "'" & matpos(i).intencion & "',"            'intencion
    txtcadena = txtcadena & matpos(i).Signo_Op & ","                    'tipo de operacion
    txtcadena = txtcadena & matpr(i).pu_sucio & ","                      'precio sucio sistema
    txtcadena = txtcadena & matpr(i).mtm_sucio & ","                      'mtm sucio
    txtcadena = txtcadena & matpr(i).ps_activa & ","                      'val activa sucio
    txtcadena = txtcadena & matpr(i).ps_pasiva & ","                      'val pasiva sucio
    txtcadena = txtcadena & matpr(i).pu_limpio & ","                      'p limpio
    txtcadena = txtcadena & matpr(i).mtm_limpio & ","                                 'mtm limpio
    txtcadena = txtcadena & matpr(i).pl_activa & ","                                  'val activa limpia
    txtcadena = txtcadena & matpr(i).pl_pasiva & ","                                  'val pasiva limpia
    txtcadena = txtcadena & ReemplazaVacioValor(matpos(i).MtmIKOS, 0) & ","           'mtm ikos
    txtcadena = txtcadena & ReemplazaVacioValor(matpos(i).ValActivaIKOS, 0) & ","     'val activa ikos
    txtcadena = txtcadena & ReemplazaVacioValor(matpos(i).ValPasivaIKOS, 0) & ","     'val activa ikos
 
    If (matpos(i).No_tabla = 1) Then
       indice = matpos(i).IndPosicion
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).valSucioPIP, 0) & ","    'val pip sucio
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).valLimpioPIP, 0) & ","   'val pip limpio
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).noTitulosMD, 0) & ","    'no de titulos
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).dVencMD, 0) & ","        'dias de vencimiento
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).dVCuponMD, 0) & ","        'dias de vencimiento
    Else
      txtcadena = txtcadena & "0,"
      txtcadena = txtcadena & "0,"
      txtcadena = txtcadena & "0,"
      txtcadena = txtcadena & "0,"
      txtcadena = txtcadena & "0,"
    End If
    txtcadena = txtcadena & ReemplazaVacioValor(matpr(i).duractiva, 0) & ","                   'duracion activa
    txtcadena = txtcadena & ReemplazaVacioValor(matpr(i).durpasiva, 0) & ","                   'duracion pasiva
    txtcadena = txtcadena & ReemplazaVacioValor(matpr(i).dv01activa, 0) & ","                  'dv01 activa
    txtcadena = txtcadena & ReemplazaVacioValor(matpr(i).dv01pasiva, 0) & ","                  'dv01 pasiva
    txtcadena = txtcadena & "" & tipopos & ","                                                 'tipo de posicion
    txtcadena = txtcadena & "'" & txtnompos & "',"                                             'nombre de la posicion
    txtcadena = txtcadena & ReemplazaVacioValor(matpr(i).p_esperada, 0) & ")"                  'perdida esperada
    ConAdo.Execute txtcadena
    AvanceProc = i / UBound(matpos, 1)
    MensajeProc = "Guardando las valuaciones de la posicion " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
   exito = True
Else
   exito = False
End If

End Sub


Sub GuardarResValPort(ByVal f_pos As Date, _
                      ByVal f_factor As Date, _
                      ByVal f_val As Date, _
                      ByVal txtport As String, _
                      ByVal txtescfr As String, _
                      ByRef matpos() As propPosRiesgo, _
                      ByRef matposmd() As propPosMD, _
                      ByVal tval As Integer, _
                      ByRef exito As Boolean)
Dim txtcadena As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfechap As String
Dim txtfechafr As String
Dim txtfechav As String
Dim i As Long
Dim indice As Long
Dim txtborra As String
Dim txtfechar As String
Dim matvalik() As Variant

txtfechap = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechafr = "to_date('" & Format(f_factor, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechav = "to_date('" & Format(f_val, "dd/mm/yyyy") & "','dd/mm/yyyy')"
If f_val = f_factor Then
   matvalik = LeerValDerivIKOS(f_val)
   For i = 1 To UBound(matpos, 1)
       If matpos(i).C_Posicion = ClavePosDeriv Then
          indice = BuscarValorArray(matpos(i).c_operacion, matvalik, 2)
          If indice <> 0 Then
             matpos(i, CValActivaIKOS) = matvalik(indice, 3)
             matpos(i, CValPasivaIKOS) = matvalik(indice, 4)
             matpos(i, CMtmIKOS) = matvalik(indice, 5)
          Else
            MsgBox "No se encontro la valuacion de la operacion en los resultados de ikos"
          End If
       End If
   Next i
End If

txtborra = "DELETE FROM " & TablaValPos & " WHERE FECHAP = " & txtfechap & " AND FECHAFR = " & txtfechafr & " AND FECHAV = " & txtfechav & " AND ID_VALUACION = " & tval & " AND PORTAFOLIO = '" & txtport & "'"
ConAdo.Execute txtborra
If UBound(matpos, 1) <> 0 Then
For i = 1 To UBound(matpos, 1)
    txtfechar = "to_date('" & Format(matpos(i).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "INSERT INTO " & TablaValPos & " VALUES("
    txtcadena = txtcadena & txtfechap & ","                             'fecha de la posicion
    txtcadena = txtcadena & txtfechafr & ","                            'fecha de los factores de riesgo
    txtcadena = txtcadena & txtfechav & ","                             'fecha de valuacion
    txtcadena = txtcadena & tval & ","                                  'tipo de valuacion
    txtcadena = txtcadena & "'" & txtport & "',"                        'PORTAFOLIO
    txtcadena = txtcadena & "'" & txtescfr & "',"                       'escenario de factores de riesgo
    txtcadena = txtcadena & matpos(i).C_Posicion & ","           'CLAVE DE LA POSICION
    txtcadena = txtcadena & txtfechar & ","                             'fecha de registro
    txtcadena = txtcadena & "'" & matpos(i).c_operacion & "',"     'clave de operacion
    txtcadena = txtcadena & "'" & matpos(i).intencion & "',"      'intencion
    txtcadena = txtcadena & matpos(i).Signo_Op & ","               'tipo de operacion
    txtcadena = txtcadena & MatPrecios(i).pu_sucio & ","                      'precio sucio sistema
    txtcadena = txtcadena & MatPrecios(i).mtm_sucio & ","                      'mtm sucio
    txtcadena = txtcadena & MatPrecios(i).ps_activa & ","                      'val activa sucio
    txtcadena = txtcadena & MatPrecios(i).ps_pasiva & ","                      'val pasiva sucio
    txtcadena = txtcadena & MatPrecios(i).pu_limpio & ","                      'p limpio
    txtcadena = txtcadena & MatPrecios(i).mtm_limpio & ","                      'mtm limpio
    txtcadena = txtcadena & MatPrecios(i).pl_activa & ","                      'val activa limpia
    txtcadena = txtcadena & MatPrecios(i).pl_pasiva & ","                      'val pasiva limpia
    txtcadena = txtcadena & ReemplazaVacioValor(matpos(i).MtmIKOS, 0) & ","          'mtm ikos
    txtcadena = txtcadena & ReemplazaVacioValor(matpos(i).ValActivaIKOS, 0) & ","    'val activa ikos
    txtcadena = txtcadena & ReemplazaVacioValor(matpos(i).ValPasivaIKOS, 0) & ","    'val pasiva ikos
    If (matpos(i).No_tabla = 1) Then
       indice = matpos(i).IndPosicion
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).valSucioPIP, 0) & ","   'val pip sucio
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).valLimpioPIP, 0) & ","  'val pip limpio
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).noTitulosMD, 0) & ","   'no de titulos
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).dVencMD, 0) & ","       'dxv
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).dVCuponMD, 0) & ","     'dxv cupon
    Else
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & "0,"
    End If
    txtcadena = txtcadena & ReemplazaVacioValor(MatPrecios(i).duractiva, 0) & ","    'duracion activa
    txtcadena = txtcadena & ReemplazaVacioValor(MatPrecios(i).durpasiva, 0) & ","   'duracion pasiva
    txtcadena = txtcadena & ReemplazaVacioValor(MatPrecios(i).dv01activa, 0) & ","   'dv01 activa
    txtcadena = txtcadena & ReemplazaVacioValor(MatPrecios(i).dv01pasiva, 0) & ","   'dv01 pasiva
    txtcadena = txtcadena & "1,"                                              'tipo de posicion
    txtcadena = txtcadena & "'Real')"                                         'nombre de la posicion

    ConAdo.Execute txtcadena
    AvanceProc = i / UBound(matpos, 1)
    MensajeProc = "Guardando las valuaciones de la posicion " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
   exito = True
Else
   exito = False
End If

End Sub

Sub GuardarResValPos(ByVal f_pos As Date, _
                     ByVal f_factor As Date, _
                     ByVal f_val As Date, _
                     ByVal txtnompos As String, _
                     ByVal txtescfr As String, _
                     ByRef matpos() As propPosRiesgo, _
                     ByRef matposmd() As propPosMD, _
                     ByRef matpr() As resValIns, _
                     ByVal tval As Integer, _
                     ByRef exito As Boolean)

Dim txtcadena As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfechap As String
Dim txtfechafr As String
Dim txtfechav As String
Dim i As Long
Dim indice As Long
Dim txtborra As String
Dim txtfechar As String
Dim matvalik() As Variant

txtfechap = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechafr = "to_date('" & Format(f_factor, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfechav = "to_date('" & Format(f_val, "dd/mm/yyyy") & "','dd/mm/yyyy')"

txtborra = "DELETE FROM " & TablaValPos & " WHERE FECHAP = " & txtfechap & " AND FECHAFR = " & txtfechafr & " AND FECHAV = " & txtfechav & " AND ID_VALUACION = " & tval & " AND PORTAFOLIO = '" & txtnompos & "'"
ConAdo.Execute txtborra
If UBound(matpos, 1) <> 0 Then
For i = 1 To UBound(matpos, 1)
    txtfechar = "to_date('" & Format(matpos(i).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "INSERT INTO " & TablaValPos & " VALUES("
    txtcadena = txtcadena & txtfechap & ","                             'fecha de la posicion
    txtcadena = txtcadena & txtfechafr & ","                            'fecha de los factores de riesgo
    txtcadena = txtcadena & txtfechav & ","                             'fecha de valuacion
    txtcadena = txtcadena & tval & ","                                  'tipo de valuacion
    txtcadena = txtcadena & "'" & txtnompos & "',"                      'PORTAFOLIO
    txtcadena = txtcadena & "'" & txtescfr & "',"                       'escenario de factores de riesgo
    txtcadena = txtcadena & matpos(i).C_Posicion & ","           'CLAVE DE LA POSICION
    txtcadena = txtcadena & txtfechar & ","                             'fecha de registro
    txtcadena = txtcadena & "'" & matpos(i).c_operacion & "',"   'clave de operacion
    txtcadena = txtcadena & "'" & matpos(i).intencion & "',"    'intencion
    txtcadena = txtcadena & matpos(i).Signo_Op & ","               'tipo de operacion
    txtcadena = txtcadena & MatPrecios(i, 1) & ","                      'precio sucio sistema
    txtcadena = txtcadena & MatPrecios(i, 2) & ","                      'mtm sucio
    txtcadena = txtcadena & MatPrecios(i, 3) & ","                      'val activa sucio
    txtcadena = txtcadena & MatPrecios(i, 4) & ","                      'val pasiva sucio
    txtcadena = txtcadena & MatPrecios(i, 5) & ","                      'p limpio
    txtcadena = txtcadena & MatPrecios(i, 6) & ","                      'mtm limpio
    txtcadena = txtcadena & MatPrecios(i, 7) & ","                      'val activa limpia
    txtcadena = txtcadena & MatPrecios(i, 8) & ","                      'val pasiva limpia
    txtcadena = txtcadena & ReemplazaVacioValor(matpos(i, CMtmIKOS), 0) & ","          'mtm ikos
    txtcadena = txtcadena & ReemplazaVacioValor(matpos(i, CValActivaIKOS), 0) & ","    'val activa ikos
    txtcadena = txtcadena & ReemplazaVacioValor(matpos(i, CValPasivaIKOS), 0) & ","    'val pasiva ikos
    If (matpos(i).No_tabla = 1) Then
       indice = matpos(i).IndPosicion
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).valSucioPIP, 0) & ","   'val pip sucio
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).valLimpioPIP, 0) & ","  'val pip limpio
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).noTitulosMD, 0) & ","   'no de titulos
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).dVencMD, 0) & ","       'dxv
       txtcadena = txtcadena & ReemplazaVacioValor(matposmd(indice).dVCuponMD, 0) & ","     'dxv cupon
    Else
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & "0,"
    End If
    txtcadena = txtcadena & ReemplazaVacioValor(MatPrecios(i, 9), 0) & ","    'duracion activa
    txtcadena = txtcadena & ReemplazaVacioValor(MatPrecios(i, 10), 0) & ","   'duracion pasiva
    txtcadena = txtcadena & ReemplazaVacioValor(MatPrecios(i, 11), 0) & ","   'dv01 activa
    txtcadena = txtcadena & ReemplazaVacioValor(MatPrecios(i, 12), 0) & ","   'dv01 pasiva
    txtcadena = txtcadena & "" & matpos(i).tipopos & ","                                         'tipo de posicion
    txtcadena = txtcadena & "'" & matpos(i).nompos & "')"                                      'nombre de la posicion
    ConAdo.Execute txtcadena
    AvanceProc = i / UBound(matpos, 1)
    MensajeProc = "Guardando las valuaciones de la posicion " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
   exito = True
Else
   exito = False
End If

End Sub

Sub LeerPyGPostxt(ByVal f_pos As Date, ByVal txtport As String, ByVal txtescfr, ByVal txtgrupo As String, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim valor As Variant
Dim valt01 As Double
Dim suma As Double
Dim txtborra As String
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim siesfv As Boolean
Dim i As Long
Dim matc() As String
Dim matf() As Date
Dim j As Long
Dim nomarch As String
Dim fecha0 As Date
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPLHistOper & " WHERE F_POSICION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport
txtfiltro2 = txtfiltro2 & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo & " AND (CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION, COPERACION FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtgrupo & "') ORDER BY CPOSICION,COPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   siesfv = EsFechaVaR(f_pos)
   If siesfv Then
   fecha0 = DetFechaFNoEsc(f_pos, noesc)
   matf = LeerFechasVaR(fecha0, f_pos)
   ReDim mattxt(1 To noesc + 3) As String
   mattxt(1) = "Clave de posicion" & Chr(9)
   mattxt(2) = "Clave de operacion" & Chr(9)
   mattxt(3) = "MtM en t0" & Chr(9)
   For i = 1 To noesc
       mattxt(i + 3) = matf(i, 1) & Chr(9)
   Next i
   
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   suma = 0
   For i = 1 To noreg
       mattxt(1) = mattxt(1) & rmesa.Fields("CPOSICION") & Chr(9)
       mattxt(2) = mattxt(2) & rmesa.Fields("COPERACION") & Chr(9)
       mattxt(3) = mattxt(3) & rmesa.Fields("VALT0") & Chr(9)
       valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
       matc = EncontrarSubCadenas(valor, ",")
       suma = suma + valt01
       For j = 1 To noesc
           mattxt(j + 3) = mattxt(j + 3) & CDbl(matc(j)) & Chr(9)
       Next j
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo las las p y g del portafolio " & txtgrupo & " " & Format(AvanceProc, "##0.0 %")
   Next i
   rmesa.Close
   nomarch = DirResVaR & "\Escenarios portafolio " & txtgrupo & " no esc " & noesc & "  " & Format(f_pos, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName

   Open nomarch For Output As #1
   For i = 1 To noesc + 3
       Print #1, mattxt(i)
   Next i
   Close #1
   End If
Else
MsgBox "No hay datos para esta fecha"
End If

End Sub


Sub LeerPyGPosSim(ByVal f_pos As Date, ByVal txtnompos As String, ByVal txtescfr, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim valor As Variant
Dim valt01 As Double
Dim suma As Double
Dim txtborra As String
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim siesfv As Boolean
Dim i As Long
Dim matc() As String
Dim matf() As Date
Dim j As Long
Dim nomarch As String
Dim fecha0 As Date
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPLHistOper & " WHERE F_POSICION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND NOMPOS = '" & txtnompos
txtfiltro2 = txtfiltro2 & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   siesfv = EsFechaVaR(f_pos)
   If siesfv Then
   fecha0 = DetFechaFNoEsc(f_pos, noesc)
   matf = LeerFechasVaR(fecha0, f_pos)
   ReDim mattxt(1 To noesc + 3) As String
   mattxt(1) = "Clave de posicion" & Chr(9)
   mattxt(2) = "Clave de operacion" & Chr(9)
   mattxt(3) = "MtM en t0" & Chr(9)
   For i = 1 To noesc
       mattxt(i + 3) = matf(i, 1) & Chr(9)
   Next i
   
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   suma = 0
   For i = 1 To noreg
       mattxt(1) = mattxt(1) & rmesa.Fields(5) & Chr(9)
       mattxt(2) = mattxt(2) & rmesa.Fields(7) & Chr(9)
       mattxt(3) = mattxt(3) & rmesa.Fields(8) & Chr(9)
       valor = rmesa.Fields(9).GetChunk(rmesa.Fields(9).ActualSize)
       matc = EncontrarSubCadenas(valor, ",")
       suma = suma + valt01
       For j = 1 To noesc
           mattxt(j + 3) = mattxt(j + 3) & CDbl(matc(j)) & Chr(9)
       Next j
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo las las p y g del portafolio " & txtnompos & " " & Format(AvanceProc, "##0.0 %")
   Next i
   rmesa.Close
   nomarch = DirResVaR & "\Escenarios portafolio " & txtnompos & " no esc " & noesc & "  " & Format(f_pos, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName

   Open nomarch For Output As #1
   For i = 1 To noesc + 3
       Print #1, mattxt(i)
   Next i
   Close #1
   End If
Else
MsgBox "No hay datos para esta f_pos"
End If

End Sub


Sub LeerPyGHistPortPos(ByVal f_pos As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim siesfv As Boolean
Dim fecha0 As Date
Dim matf() As Date
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim indice As Long
Dim i As Long
Dim j As Long
Dim l As Long
Dim noreg As Long
Dim nomarch As String
Dim valor As String
Dim matc() As String
Dim rmesa As New ADODB.recordset

MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
If UBound(MatGruposPortPos, 1) <> 0 Then
   ReDim mattxt(1 To noesc + 2) As String
   siesfv = EsFechaVaR(f_pos)
   fecha0 = DetFechaFNoEsc(f_pos, noesc)
   matf = LeerFechasVaR(fecha0, f_pos)
   mattxt(1) = "Subportafolio" & Chr(9)
   mattxt(2) = "MtM t0" & Chr(9)
   For i = 1 To noesc
       mattxt(i + 2) = matf(i, 1) & Chr(9)
   Next i
   txtfecha = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   For i = 1 To UBound(MatGruposPortPos, 1)
       txtfiltro2 = "SELECT * FROM " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = "
       txtfiltro2 = txtfiltro2 & "'" & txtport & "' AND SUBPORT = '" & MatGruposPortPos(i, 3) & "'"
       txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       mattxt(1) = mattxt(1) & MatGruposPortPos(i, 3) & Chr(9)
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          rmesa.MoveFirst
          mattxt(2) = mattxt(2) & rmesa.Fields("VALT0") & Chr(9)
          valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
          matc = EncontrarSubCadenas(valor, ",")
          For l = 1 To UBound(matc, 1)
              mattxt(l + 2) = mattxt(l + 2) & CDbl(matc(l)) & Chr(9)
          Next l
          rmesa.Close
       Else
          For j = 1 To noesc + 1
              mattxt(j + 1) = mattxt(j + 1) & 0 & Chr(9)
          Next j
       End If
   Next i
   nomarch = DirResVaR & "\Escenarios p y g  " & txtport & " subport " & txtgrupoport & " esc fr " & txtportfr & " no esc " & noesc & " " & Format(f_pos, "YYYY-MM-DD") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName
   Open nomarch For Output As #1
   For i = 1 To noesc + 2
       Print #1, mattxt(i)
   Next i
   Close #1
End If
End Sub

Sub LeerPyGHistPortPos2(ByVal f_pos As Date, ByVal txtport As String, ByVal txtportfr As String, ByRef matport() As String, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim siesfv As Boolean
Dim fecha0 As Date
Dim i As Long
Dim j As Long
Dim l As Long
Dim noreg As Long
Dim nomarch As String
Dim valor As String
Dim matc() As String
Dim matf() As Date
Dim rmesa As New ADODB.recordset

If UBound(matport, 1) <> 0 Then
   ReDim mattxt(1 To noesc + 2) As String
   siesfv = EsFechaVaR(f_pos)
   fecha0 = DetFechaFNoEsc(f_pos, noesc)
   matf = LeerFechasVaR(fecha0, f_pos)
   mattxt(1) = "Subportafolio" & Chr(9)
   mattxt(2) = "MtM t0" & Chr(9)
   For i = 1 To noesc
       mattxt(i + 2) = matf(i, 1) & Chr(9)
   Next i
   txtfecha = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   For i = 1 To UBound(matport, 1)
       txtfiltro2 = "SELECT * FROM " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = "
       txtfiltro2 = txtfiltro2 & "'" & txtport & "' AND SUBPORT = '" & matport(i, 1) & "'"
       txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       mattxt(1) = mattxt(1) & matport(i, 1) & Chr(9)
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          rmesa.MoveFirst
          mattxt(2) = mattxt(2) & rmesa.Fields(6) & Chr(9)
          valor = rmesa.Fields("DATOS").GetChunk(rmesa.Fields("DATOS").ActualSize)
          matc = EncontrarSubCadenas(valor, ",")
          For l = 1 To UBound(matc, 1)
              mattxt(l + 2) = mattxt(l + 2) & CDbl(matc(l)) & Chr(9)
          Next l
          rmesa.Close
       Else
          For j = 1 To noesc + 1
              mattxt(j + 1) = mattxt(j + 1) & 0 & Chr(9)
          Next j
       End If
   Next i
   nomarch = DirResVaR & "\Escenarios p y g  " & txtport & " esc fr " & txtportfr & " no esc " & noesc & " " & Format(f_pos, "YYYY-MM-DD") & ".txt"
   nomarch = "d:\Escenarios p y g  " & txtport & " esc fr " & txtportfr & " no esc " & noesc & " " & Format(f_pos, "YYYY-MM-DD") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName
   Open nomarch For Output As #1
   For i = 1 To noesc + 2
       Print #1, mattxt(i)
   Next i
   Close #1
End If
End Sub

Sub LeerPyGHistPortPos3(ByVal f_pos As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim siesfv As Boolean
Dim fecha0 As Date
Dim matf() As Date
Dim i As Long
Dim j As Long
Dim l As Long
Dim noreg As Long
Dim nomarch As String
Dim valor As String
Dim matc() As String
Dim rmesa As New ADODB.recordset

MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
If UBound(MatGruposPortPos, 1) <> 0 Then
   ReDim mattxt(1 To noesc + 2) As String
   siesfv = EsFechaVaR(f_pos)
   fecha0 = DetFechaFNoEsc(f_pos, noesc)
   matf = LeerFechasVaR(fecha0, f_pos)
   mattxt(1) = f_pos & Chr(9) & "Subportafolio" & Chr(9)
   mattxt(2) = f_pos & Chr(9) & "MtM t0" & Chr(9)
   For i = 1 To noesc
       mattxt(i + 2) = f_pos & Chr(9) & matf(i, 1) & Chr(9)
   Next i
   txtfecha = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   For i = 1 To UBound(MatGruposPortPos, 1)
       txtfiltro2 = "SELECT * FROM " & TablaPLEscHistPort & " WHERE F_POSICION = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = "
       txtfiltro2 = txtfiltro2 & "'" & txtport & "' AND SUBPORT = '" & MatGruposPortPos(i, 3) & "'"
       txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = '" & txtportfr & "' AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg = rmesa.Fields(0)
       rmesa.Close
       mattxt(1) = mattxt(1) & MatGruposPortPos(i, 3) & Chr(9)
       If noreg <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          rmesa.MoveFirst
          mattxt(2) = mattxt(2) & rmesa.Fields(6) & Chr(9)
          valor = rmesa.Fields(7).GetChunk(rmesa.Fields(7).ActualSize)
          matc = EncontrarSubCadenas(valor, ",")
          For l = 1 To UBound(matc, 1)
              mattxt(l + 2) = mattxt(l + 2) & CDbl(matc(l)) & Chr(9)
          Next l
          rmesa.Close
       Else
          For j = 1 To noesc + 1
              mattxt(j + 1) = mattxt(j + 1) & 0 & Chr(9)
          Next j
       End If
   Next i
   For i = 1 To noesc + 2
       Print #1, mattxt(i)
   Next i
End If
End Sub

Sub LeerPyGMontOper(ByVal fecha As Date, ByVal txtport As String, ByVal txtescfr, ByVal txtport2 As String, ByVal noesc As Integer, ByVal htiempo As Integer, ByVal nosim As Integer)
Dim valt01 As Double
Dim suma As Double
Dim txtborra As String
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg As Long
Dim j As Long
Dim nomarch As String
Dim valor As String
Dim matc() As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPyGMontOper & " WHERE FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport
txtfiltro2 = txtfiltro2 & "' AND ESC_FACTORES = '" & txtescfr & "' AND NOESC = " & noesc
txtfiltro2 = txtfiltro2 & " AND HTIEMPO = " & htiempo
txtfiltro2 = txtfiltro2 & " AND NOSIM = " & nosim
txtfiltro2 = txtfiltro2 & " AND (CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION, COPERACION FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha & " AND PORTAFOLIO = '" & txtport2 & "') ORDER BY CPOSICION,COPERACION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mattxt(1 To nosim + 3) As String
   mattxt(1) = "Clave de posicion" & Chr(9)
   mattxt(2) = "Clave de operacion" & Chr(9)
   mattxt(3) = "MtM en t0" & Chr(9)
   For i = 1 To nosim
       mattxt(i + 3) = i & Chr(9)
   Next i
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       mattxt(1) = mattxt(1) & rmesa.Fields(6) & Chr(9)
       mattxt(2) = mattxt(2) & rmesa.Fields(8) & Chr(9)
       mattxt(3) = mattxt(3) & rmesa.Fields(9) & Chr(9)
       valor = rmesa.Fields(10).GetChunk(rmesa.Fields(10).ActualSize)
       matc = EncontrarSubCadenas(valor, ",")
       For j = 1 To nosim
           mattxt(j + 3) = mattxt(j + 3) & CDbl(matc(j)) & Chr(9)
       Next j
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo las las p y g del portafolio " & txtport2 & " " & Format(AvanceProc, "##0.0 %")
   Next i
   rmesa.Close
   nomarch = DirResVaR & "\Escenarios Montecarlo portafolio " & txtport2 & " no esc " & noesc & "  " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName

   Open nomarch For Output As #1
   For i = 1 To nosim + 3
       Print #1, mattxt(i)
   Next i
   Close #1
End If

End Sub

Sub GuardarEscEstres(ByVal fecha As Date, _
                     ByVal txtport As String, _
                     ByVal txtescestres As String, _
                     ByRef matpos() As propPosRiesgo, _
                     ByRef mata() As Double)

Dim txtcadena As String
Dim txtfecha As String
Dim txtborra As String
Dim txtfechareg As String
Dim i As Long

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
For i = 1 To UBound(matpos, 1)
    txtfechareg = "to_date('" & Format(matpos(i).fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "INSERT INTO " & TablaResEscEstres & " VALUES("
    txtcadena = txtcadena & txtfecha & ","                            'FECHA DE LA POSICION
    txtcadena = txtcadena & "'" & txtescestres & "',"                 'escenario de estres
    txtcadena = txtcadena & "'" & txtport & "',"                      'PORTAFOLIO
    txtcadena = txtcadena & matpos(i).C_Posicion & ","         'CLAVE DE LA POSICION
    txtcadena = txtcadena & txtfechareg & ","                         'fecha de registro
    txtcadena = txtcadena & "'" & matpos(i).c_operacion & "'," 'clave de operacion
    txtcadena = txtcadena & mata(i, 1) & ")"                          'VALOR OBSERVADO DE ESTRES
    ConAdo.Execute txtcadena
    AvanceProc = i / UBound(matpos, 1)
    MensajeProc = "Guardando el escenario de estres " & Format(AvanceProc, "##0.00 %")
Next i
End Sub

Sub GenerarEscEstresPort(ByVal fecha As Date, ByVal txtestres As String, ByVal txtport As String, ByVal txtsubport As String, ByRef matres() As Variant, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtborra As String
Dim txtinserta As String
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim valor As Double
Dim suma As Double
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtsubport & "'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 5) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("CPOSICION")
       mata(i, 2) = rmesa.Fields("FECHAREG")
       mata(i, 3) = rmesa.Fields("COPERACION")
       rmesa.MoveNext
   Next i
   rmesa.Close
   suma = 0
   For i = 1 To UBound(matres, 1)
       For j = 1 To noreg
           If mata(j, 1) = matres(i, 3) And mata(j, 2) = matres(i, 4) And mata(j, 3) = matres(i, 5) And matres(i, 2) = txtestres Then
              suma = suma + matres(i, 6)
           End If
       Next j
   Next i
   
   txtinserta = "INSERT INTO " & TablaResEscEstresPort & " VALUES("
   txtinserta = txtinserta & txtfecha & ","
   txtinserta = txtinserta & "'" & txtestres & "',"
   txtinserta = txtinserta & "'" & txtport & "',"
   txtinserta = txtinserta & "'" & txtsubport & "',"
   txtinserta = txtinserta & suma & ")"
   ConAdo.Execute txtinserta
   MensajeProc = "Guardando el escenario " & txtestres
   txtmsg = "El proceso finalizo correctamente"
   exito = True
Else
   If txtsubport = "PI DISPONIBLE PARA LA VENTA" Or txtsubport = "PI DERIVADOS" Or txtsubport = "CONSOLIDADO Y RELACIONADOS" Or txtsubport = "ESTRUCTURALES Y RELACIONADOS" Or txtsubport = "DERIVADOS DE COBERTURA" Or txtsubport = "DERIVADOS NEGOCIACION RECLASIFICACION" Then
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      txtmsg = "No hay datos para este portafolio"
      exito = False
   End If
End If
End Sub

Sub ObtenerYieldsIS(ByVal fecha As Date, ByRef txtmsg As String, ByRef exito As Boolean)
Dim matf() As New estFlujosMD
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim i As Integer
Dim noreg As Integer
Dim noreg1 As Integer
Dim valtasa As Double
Dim valor As Double
Dim yield As Double
Dim txtconcepto1 As String
Dim txtindice1 As String
Dim txtborra1 As String
Dim txtinserta1 As String
Dim rmesa As New ADODB.recordset

exito = True
txtmsg = "El proceso finalizo correctamente"


txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaVecPrecios & " WHERE FECHA = " & txtfecha & " AND TV = 'IS'"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 4) As Variant
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("SERIE")                 'serie
       mata(i, 2) = rmesa.Fields("PSUCIO")                'precio sucio
       mata(i, 3) = rmesa.Fields("TCUPON") / 100          'tasa cupon vigente
       mata(i, 4) = Val(rmesa.Fields("YIELD")) / 100      'yield
       rmesa.MoveNext
   Next i
   rmesa.Close
   txtfiltro2 = "SELECT * FROM " & TablaVecPrecios & " WHERE FECHA = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TV ='TR' AND SERIE = '182'"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg1 = rmesa.Fields(0)
   rmesa.Close
   If noreg1 <> 0 Then
      rmesa.Open txtfiltro2, ConAdo
      valtasa = rmesa.Fields("PSUCIO") / 100
      rmesa.Close
      exito = True
      txtmsg = ""
      For i = 1 To noreg
          matf = CFlujosEmisionesMD(fecha, "IS" & mata(i, 1), False)
          If UBound(matf, 1) <> 0 Then
             valor = 0
             valor = YieldPIPAB(fecha, matf, mata(i, 2), mata(i, 3), valtasa, 182)
             If mata(i, 4) <> 0 Then
                yield = mata(i, 4)
                If Abs(mata(i, 4) - valor) > 0.01 Then
                   txtmsg = txtmsg & "Hay diferencias para el calculo de la yield de la emisión IS" & mata(i, 1) & ","
                   exito = False
                End If
             Else
                yield = valor
             End If
             txtconcepto1 = "Y IS" & mata(i, 1)
             txtindice1 = CLng(fecha) & txtconcepto1 & "0000000"
             txtborra1 = "DELETE FROM " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha & " AND CONCEPTO = '" & txtconcepto1 & "'"
             ConAdo.Execute txtborra1
             txtinserta1 = "INSERT INTO " & TablaFRiesgoO & " VALUES("
             txtinserta1 = txtinserta1 & txtfecha & ","
             txtinserta1 = txtinserta1 & "'" & txtconcepto1 & "',"
             txtinserta1 = txtinserta1 & "0,"
             txtinserta1 = txtinserta1 & valor & ","
             txtinserta1 = txtinserta1 & "'" & txtindice1 & "')"
             ConAdo.Execute txtinserta1
          Else
             txtmsg = "No hay flujos para la emision IS" & mata(i, 1)
             exito = False
          End If
      Next i
      If Len(txtmsg) = 0 Then txtmsg = "El proceso finalizo correctamente"
   Else
      txtmsg = "No hay tasa de referencia CT182 para esta fecha"
      exito = False
   End If
Else
   txtmsg = "No hay emisiones IS para esta fecha"
   exito = False
End If
End Sub

Sub GenValPosGrupoPort(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtportfr As String, ByVal txtgrupoport As String, ByVal tval As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
Dim noport As Long
Dim i As Long
Dim j As Long
Dim exito1 As Boolean
Dim exito2 As Boolean

exito1 = True
MatGruposPortPos = CargaGruposPortPos(txtgrupoport)
If UBound(MatGruposPortPos, 1) <> 0 Then
   For i = 1 To UBound(MatGruposPortPos, 1)
       Call CalcValSubportPos(f_pos, f_factor, f_val, txtport, MatGruposPortPos(i, 3), txtportfr, tval, exito2)
       exito1 = exito1 And exito2
   Next i
   If exito1 Then
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   Else
      txtmsg = "Algo salio mal"
      exito = False
   End If
Else
   exito = False
End If
End Sub

Sub GenValPortContrap(ByVal fecha As Date, ByVal txtportfr As String, ByVal txtport As String, ByVal txtgrupoport As String, ByVal id_proc As Integer, ByVal id_tabla As Integer, ByRef exito As Boolean)
Dim noreg As Long
Dim i As Long
Dim contar As Long
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtcadena As String
Dim mata() As String
Dim rmesa As New ADODB.recordset

exito = False
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT PORTAFOLIO FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO LIKE '" & txtgrupoport & "%' GROUP BY PORTAFOLIO ORDER BY PORTAFOLIO"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg) As String
   For i = 1 To noreg
       mata(i) = rmesa.Fields("PORTAFOLIO")
       rmesa.MoveNext
   Next i
   rmesa.Close
   contar = DeterminaMaxRegSubproc(id_tabla)
   For i = 1 To noreg
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Valuación de contraparte", txtport, mata(i), txtportfr, 1, "", "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
       contar = contar + 1
       txtcadena = CrearCadInsSub(fecha, id_proc, contar, "Valuación de contraparte", txtport, mata(i), txtportfr, 2, "", "", "", "", "", "", "", "", id_tabla)
       ConAdo.Execute txtcadena
   Next i
   exito = True
End If
End Sub

Sub CalcValSubportPos(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtsubport As String, ByRef txtportfr As String, ByVal tval As Integer, ByRef exito As Boolean)

Dim noreg As Long
Dim i As Long
Dim valor As Variant
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim valor4 As Double
Dim valor5 As Double
Dim valor6 As Double
Dim valor7 As Double
Dim valor8 As Double
Dim valor9 As Double
Dim valor10 As Double
Dim valor11 As Double
Dim valor12 As Double
Dim valor13 As Double
Dim val_PE As Double

Dim mtms As Double
Dim mtml As Double
Dim valacts As Double
Dim valpass As Double
Dim valactl As Double
Dim valpasl As Double
Dim vdv01act As Double
Dim vdv01pas As Double
Dim ntitulosa As Double
Dim ntitulosp As Double
Dim DVCuponAct As Double
Dim DVCuponPas As Double
Dim dvactiva As Double
Dim dvpasiva As Double
Dim duract As Double
Dim durpas As Double

Dim valt01 As Double
Dim sumant As Integer
Dim sumaval As Double
Dim sumadv01 As Double
Dim suma_pe As Double
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtborra As String
Dim txtcadena As String
Dim signo As Double
Dim rmesa As New ADODB.recordset

On Error GoTo hayerror

txtfecha1 = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(f_factor, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha3 = "to_date('" & Format(f_val, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtborra = "DELETE FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha1 & " AND FECHAFR = " & txtfecha2 & " AND FECHAV = " & txtfecha3 & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT = '" & txtsubport & "' AND ID_VALUACION = " & tval
ConAdo.Execute txtborra
txtfiltro2 = "SELECT * FROM " & TablaValPos & " WHERE FECHAP = " & txtfecha1
txtfiltro2 = txtfiltro2 & " AND FECHAFR = " & txtfecha2
txtfiltro2 = txtfiltro2 & " AND FECHAV = " & txtfecha3
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "' AND ID_VALUACION = " & tval
txtfiltro2 = txtfiltro2 & " AND (CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION, COPERACION FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha1 & " AND PORTAFOLIO = '" & txtsubport & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
mtms = 0
valacts = 0
valpass = 0
mtml = 0
valactl = 0
valpasl = 0
    If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       signo = rmesa.Fields(10)       'signo de la operacion
       valor1 = rmesa.Fields(12)      'mtms
       valor2 = rmesa.Fields(13)      'val activas
       valor3 = rmesa.Fields(14)      'val pasivas
       valor4 = rmesa.Fields(16)      'mtml
       valor5 = rmesa.Fields(17)      'val actival
       valor6 = rmesa.Fields(18)      'val pasival
       valor7 = Val(rmesa.Fields(27)) 'dur activa
       valor8 = rmesa.Fields(28)      'dur pasiva
       valor9 = rmesa.Fields(29)      'dv01 activa
       valor10 = rmesa.Fields(30)     'dv01 pasiva
       valor11 = rmesa.Fields(24)     'n titulos
       valor12 = rmesa.Fields(25)             'dias de vencimiento
       valor13 = rmesa.Fields(26)             'dias de vencimiento del cupon
       val_PE = rmesa.Fields("P_ESPERADA")    'perdida esperada
       mtms = mtms + valor1           'suma mtms
       suma_pe = suma_pe + val_PE       'suma perdida esperada
       valacts = valacts + valor2     'suma pos activa
       valpass = valpass + valor3     'suma pos pasiva
       mtml = mtml + valor4           'suma mtml
       valactl = valactl + valor5
       valpasl = valpasl + valor6
       duract = duract + valor2 * valor7
       durpas = durpas + valor3 * valor8
       If signo = 1 Then
          vdv01act = vdv01act + valor11 * valor9
          ntitulosa = ntitulosa + valor11
          dvactiva = dvactiva + valor12 * valor2
          DVCuponAct = DVCuponAct + valor13 * valor2
       Else
          vdv01pas = vdv01pas + valor11 * valor10
          ntitulosp = ntitulosp + valor11
          dvpasiva = dvpasiva + valor12 * valor3
          DVCuponPas = DVCuponPas + valor13 * valor3
       End If
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo las valuaciones del portafolio " & txtsubport & " " & Format(AvanceProc, "##0.0 %")
       DoEvents
   Next i
   If valacts <> 0 Then
      duract = duract / valacts
      dvactiva = dvactiva / valacts
      DVCuponAct = DVCuponAct / valacts
   End If
   If valpass <> 0 Then
      durpas = durpas / valpass
      dvpasiva = dvpasiva / valpass
      DVCuponPas = DVCuponPas / valpass
   End If
   rmesa.Close
   txtcadena = "INSERT INTO " & TablaValPosPort & " VALUES("
   txtcadena = txtcadena & txtfecha1 & ","
   txtcadena = txtcadena & txtfecha2 & ","
   txtcadena = txtcadena & txtfecha3 & ","
   txtcadena = txtcadena & tval & ","
   txtcadena = txtcadena & "'" & txtportfr & "',"
   txtcadena = txtcadena & "'" & txtport & "',"
   txtcadena = txtcadena & "'" & txtsubport & "',"
   txtcadena = txtcadena & mtms & ","
   txtcadena = txtcadena & valacts & ","
   txtcadena = txtcadena & valpass & ","
   txtcadena = txtcadena & mtml & ","           'mtm limpio
   txtcadena = txtcadena & valactl & ","        'val activa limpia
   txtcadena = txtcadena & valpasl & ","
   If Not EsVariableVacia(ntitulosa) Then
   txtcadena = txtcadena & ntitulosa & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(ntitulosp) Then
      txtcadena = txtcadena & ntitulosp & ","
   Else
      txtcadena = txtcadena & "0,"
   End If

   
   If Not EsVariableVacia(duract) Then
   txtcadena = txtcadena & duract & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(durpas) Then
      txtcadena = txtcadena & durpas & ","
   Else
      txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(vdv01act) Then
      txtcadena = txtcadena & vdv01act & ","
   Else
      txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(vdv01pas) Then
   txtcadena = txtcadena & vdv01pas & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(DVCuponAct) Then
   txtcadena = txtcadena & DVCuponAct & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(DVCuponPas) Then
   txtcadena = txtcadena & DVCuponPas & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(dvactiva) Then
   txtcadena = txtcadena & dvactiva & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(dvpasiva) Then
   txtcadena = txtcadena & dvpasiva & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   txtcadena = txtcadena & ReemplazaVacioValor(suma_pe, 0) & ")"
   ConAdo.Execute txtcadena
   MensajeProc = "Guardando las valuaciones del portafolio " & txtsubport
   DoEvents
End If
exito = True
On Error GoTo 0
Exit Sub
hayerror:
MsgBox error(Err())
If Err() = "03113" Then
   Call ReiniciarConexOracleP(ConAdo)
   exito = False
End If
On Error GoTo 0
End Sub


Sub GenValPortPosPension(ByVal f_pos As Date, ByVal f_factor As Date, ByVal f_val As Date, ByVal txtport As String, ByVal txtsubport As String, ByRef txtportfr As String, ByVal tval As Integer, ByRef exito As Boolean)

Dim noreg As Long
Dim i As Long
Dim valor As Variant
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim valor4 As Double
Dim valor5 As Double
Dim valor6 As Double
Dim valor7 As Double
Dim valor8 As Double
Dim valor9 As Double
Dim valor10 As Double
Dim valor11 As Double
Dim valor12 As Double
Dim valor13 As Double

Dim mtms As Double
Dim mtml As Double
Dim valacts As Double
Dim valpass As Double
Dim valactl As Double
Dim valpasl As Double
Dim vdv01act As Double
Dim vdv01pas As Double
Dim ntitulosa As Double
Dim ntitulosp As Double
Dim DVCuponAct As Double
Dim DVCuponPas As Double
Dim dvactiva As Double
Dim dvpasiva As Double
Dim duract As Double
Dim durpas As Double

Dim valt01 As Double
Dim sumant As Integer
Dim sumaval As Double
Dim sumadv01 As Double
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String

Dim txtborra As String
Dim txtcadena As String
Dim signo As Double
Dim valpip As Double
Dim rmesa As New ADODB.recordset

On Error GoTo hayerror

txtfecha1 = "to_date('" & Format(f_pos, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(f_factor, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha3 = "to_date('" & Format(f_val, "dd/mm/yyyy") & "','dd/mm/yyyy')"

txtborra = "DELETE FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha1 & " AND FECHAFR = " & txtfecha2 & " AND FECHAV = " & txtfecha3 & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT = '" & txtsubport & "' AND ID_VALUACION = " & tval
ConAdo.Execute txtborra
txtfiltro2 = "SELECT * FROM " & TablaValPos & " WHERE FECHAP = " & txtfecha1
txtfiltro2 = txtfiltro2 & " AND FECHAFR = " & txtfecha2
txtfiltro2 = txtfiltro2 & " AND FECHAV = " & txtfecha3
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "' AND ID_VALUACION = " & tval
txtfiltro2 = txtfiltro2 & " AND (CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION, COPERACION FROM " & TablaPortPosicion & "  WHERE FECHA_PORT = " & txtfecha1 & " AND PORTAFOLIO = '" & txtsubport & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
mtms = 0
valacts = 0
valpass = 0
mtml = 0
valactl = 0
valpasl = 0
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   For i = 1 To noreg
       signo = rmesa.Fields("T_OPERACION")       'signo de la operacion
       valor11 = rmesa.Fields("NO_TITULOS_")    'n titulos
       valpip = rmesa.Fields("VAL_PIP_S")       'VALUACION DE PIP
       If valpip <> 0 Then
          valor1 = signo * valpip * valor11     'mtms
       Else
          valor1 = rmesa.Fields("MTM_S")        'mtms
       End If
       valor2 = rmesa.Fields("VAL_ACT_S")       'val activas
       valor3 = rmesa.Fields(14)                'val pasivas
       valor4 = rmesa.Fields(16)                'mtml
       valor5 = rmesa.Fields(17)                'val actival
       valor6 = rmesa.Fields(18)                'val pasival
       valor7 = Val(rmesa.Fields(27))           'dur activa
       valor8 = rmesa.Fields(28)                'dur pasiva
       valor9 = rmesa.Fields(29)                'dv01 activa
       valor10 = rmesa.Fields(30)               'dv01 pasiva
       valor12 = rmesa.Fields(25)     'dias de vencimiento
       valor13 = rmesa.Fields(26)     'dias de vencimiento del cupon
       mtms = mtms + valor1           'suma mtms
       valacts = valacts + valor2     'suma pos activa
       valpass = valpass + valor3     'suma pos pasiva
       mtml = mtml + valor4           'suma mtml
       valactl = valactl + valor5
       valpasl = valpasl + valor6
       duract = duract + valor2 * valor7
       durpas = durpas + valor3 * valor8
       If signo = 1 Then
          vdv01act = vdv01act + valor11 * valor9
          ntitulosa = ntitulosa + valor11
          dvactiva = dvactiva + valor12 * valor2
          DVCuponAct = DVCuponAct + valor13 * valor2
       Else
          vdv01pas = vdv01pas + valor11 * valor10
          ntitulosp = ntitulosp + valor11
          dvpasiva = dvpasiva + valor12 * valor3
          DVCuponPas = DVCuponPas + valor13 * valor3
       End If
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo las valuaciones del portafolio " & txtsubport & " " & Format(AvanceProc, "##0.0 %")
       DoEvents
   Next i
   If valacts <> 0 Then
      duract = duract / valacts
      dvactiva = dvactiva / valacts
      DVCuponAct = DVCuponAct / valacts
   End If
   If valpass <> 0 Then
      durpas = durpas / valpass
      dvpasiva = dvpasiva / valpass
      DVCuponPas = DVCuponPas / valpass
   End If
   rmesa.Close
   txtcadena = "INSERT INTO " & TablaValPosPort & " VALUES("
   txtcadena = txtcadena & txtfecha1 & ","
   txtcadena = txtcadena & txtfecha2 & ","
   txtcadena = txtcadena & txtfecha3 & ","
   txtcadena = txtcadena & tval & ","
   txtcadena = txtcadena & "'" & txtportfr & "',"
   txtcadena = txtcadena & "'" & txtport & "',"
   txtcadena = txtcadena & "'" & txtsubport & "',"
   txtcadena = txtcadena & mtms & ","
   txtcadena = txtcadena & valacts & ","
   txtcadena = txtcadena & valpass & ","
   txtcadena = txtcadena & mtml & ","           'mtm limpio
   txtcadena = txtcadena & valactl & ","        'val activa limpia
   txtcadena = txtcadena & valpasl & ","
   If Not EsVariableVacia(ntitulosa) Then
   txtcadena = txtcadena & ntitulosa & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(ntitulosp) Then
      txtcadena = txtcadena & ntitulosp & ","
   Else
      txtcadena = txtcadena & "0,"
   End If

   
   If Not EsVariableVacia(duract) Then
   txtcadena = txtcadena & duract & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(durpas) Then
      txtcadena = txtcadena & durpas & ","
   Else
      txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(vdv01act) Then
      txtcadena = txtcadena & vdv01act & ","
   Else
      txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(vdv01pas) Then
   txtcadena = txtcadena & vdv01pas & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(DVCuponAct) Then
   txtcadena = txtcadena & DVCuponAct & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(DVCuponPas) Then
   txtcadena = txtcadena & DVCuponPas & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(dvactiva) Then
   txtcadena = txtcadena & dvactiva & ","
   Else
   txtcadena = txtcadena & "0,"
   End If
   If Not EsVariableVacia(dvpasiva) Then
   txtcadena = txtcadena & dvpasiva & ")"
   Else
   txtcadena = txtcadena & "0)"
   End If
   ConAdo.Execute txtcadena
   MensajeProc = "Guardando las valuaciones del portafolio " & txtsubport
   DoEvents
End If
exito = True
On Error GoTo 0
Exit Sub
hayerror:
MsgBox error(Err())
If Err() = "03113" Then
   Call ReiniciarConexOracleP(ConAdo)
   exito = False
End If
On Error GoTo 0
End Sub

Function GenRepVaRMD(ByVal fecha As Date, ByVal txtrep As String, ByVal capneto As Double, ByVal valcvar As Double, ByVal vallim As Double, ByVal valcons As Double)
Dim suma As Double
Dim contar As Integer
Dim i As Integer
Dim j As Integer
Dim matb() As Double
Dim matport() As Variant
Dim nocol As Integer

contar = 1
nocol = 7
matport = CargaPortReporteCVAR(txtrep)
If UBound(matport, 1) <> 0 Then
ReDim mata(1 To nocol, 1 To contar) As Variant
   mata(1, 1) = "Portafolio"
   mata(2, 1) = "Posición activa"
   mata(3, 1) = "Posición pasiva"
   mata(4, 1) = "Marca a mercado"
   mata(5, 1) = "VaR Markowitz"
   mata(6, 1) = "VaR Montecarlo"
   mata(7, 1) = "CVaR Historico"
   For i = 1 To UBound(matport, 1)
       matb = GenRengRepVaR(fecha, txtportCalc1, matport(i, 3), suma)
       If suma <> 0 Then
          contar = contar + 1
       ReDim Preserve mata(1 To nocol, 1 To contar) As Variant
          mata(1, contar) = matport(i, 4)
          For j = 1 To nocol - 1
              mata(j + 1, contar) = Format(matb(j), "###,###,###,###,###,###,##0")
          Next j
       End If
   Next i
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar)
mata(1, contar) = "CAPITAL NETO"
mata(nocol, contar) = Format(capneto, "###,###,###,###,###,##0.00")
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar)
mata(1, contar) = "LIMITE COMO % CAPITAL NETO"
mata(nocol, contar) = Format(vallim, "##0.0000 %")
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar)
mata(1, contar) = "LIMITE EN PESOS"
mata(nocol, contar) = Format(capneto * vallim, "###,###,###,###,###,##0.00")
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar)
mata(1, contar) = "CVaR MERCADO DE DINERO"
mata(nocol, contar) = Format(-valcvar, "###,###,###,###,###,##0.00")
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar)
mata(1, contar) = "CONSUMO DE LIMITE"
mata(nocol, contar) = Format(valcons, "##0.0000 %")
GenRepVaRMD = MTranV(mata)
Else
End If
End Function

Function GenRepVaRDeriv(ByVal fecha As Date, ByVal vallim1 As Double, ByVal valcons1 As Double, ByVal vallim2 As Double, ByVal valcons2 As Double, ByVal vallim3 As Double, ByVal valcons3 As Double) As Variant()
Dim suma As Double
Dim contar As Long
Dim i As Long
Dim j As Long
Dim matb() As Double
Dim matport() As Variant
Dim nocol As Integer

contar = 1
nocol = 7
matport = CargaPortReporteCVAR("REPORTE DERIV")
ReDim mata(1 To nocol, 1 To contar) As Variant
mata(1, 1) = "Portafolio"
mata(2, 1) = "Posición activa"
mata(3, 1) = "Posición pasiva"
mata(4, 1) = "Marca a mercado"
mata(5, 1) = "VaR Markowitz"
mata(6, 1) = "VaR Montecarlo"
mata(7, 1) = "CVaR Historico"

For i = 1 To UBound(matport, 1)
    matb = GenRengRepVaR(fecha, txtportCalc1, matport(i, 3), suma)
    If suma <> 0 Then
       contar = contar + 1
       ReDim Preserve mata(1 To nocol, 1 To contar) As Variant
       mata(1, contar) = matport(i, 4)
       For j = 1 To nocol - 1
           mata(j + 1, contar) = Format(matb(j), "###,###,###,###,###,###,##0")
       Next j
     End If
Next i
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar) As Variant
mata(1, contar) = "Límite Derivados"
mata(nocol, contar) = Format(vallim1, "##0.00 %")
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar) As Variant
mata(1, contar) = "Consumo de límite"
mata(nocol, contar) = Format(valcons1, "##0.00 %")
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar) As Variant
mata(1, contar) = "Límite Derivados Estructurales"
mata(nocol, contar) = Format(vallim2, "##0.00 %")
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar) As Variant
mata(1, contar) = "Consumo de límite"
mata(nocol, contar) = Format(valcons2, "##0.00 %")
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar) As Variant
mata(1, contar) = "Límite Reclasificación"
mata(nocol, contar) = Format(vallim3, "##0.00 %")
contar = contar + 1
ReDim Preserve mata(1 To nocol, 1 To contar) As Variant
mata(1, contar) = "Consumo de límite"
mata(nocol, contar) = Format(valcons3, "##0.00 %")
GenRepVaRDeriv = MTranV(mata)

End Function

Function GenRengRepVaR(ByVal fecha As Date, ByVal txtport As String, ByVal txtsubport As String, ByRef suma As Double)
Dim txtport2 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Long
Dim i As Long
Dim noesc As Integer
Dim rmesa As New ADODB.recordset

noesc = 500
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
ReDim mata(1 To 6) As Double
txtfiltro2 = "SELECT * FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FECHAFR = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND FECHAV = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & txtsubport & "'"
txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = 1"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   mata(1) = rmesa.Fields(8)
   mata(2) = rmesa.Fields(9)
   mata(3) = rmesa.Fields(7)
   rmesa.Close
Else
   mata(1) = 0
   mata(2) = 0
   mata(3) = 0
End If

mata(4) = -LeerResVaRTabla(fecha, txtportCalc2, txtsubport, "VARMark", noesc, 1, 0, 0.01, 0)
mata(5) = -LeerResVaRTabla(fecha, txtportCalc2, txtsubport, "VARMont", noesc, 1, 10000, 0.01, 0)
mata(6) = -LeerResVaRTabla(fecha, txtport, txtsubport, "CVARH", noesc, 1, 0, 0.03, 0)

suma = 0
For i = 1 To 6
    suma = suma + Abs(mata(i))
Next i
GenRengRepVaR = mata

End Function

Sub GuardaFlujosEmEsp(ByRef matem() As Variant, ByRef mata() As Variant)
Dim i As Long
Dim txtfecha0 As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtcadena As String


For i = 1 To UBound(matem, 1)
    txtfecha0 = "to_date('" & Format(matem(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    ConAdo.Execute "DELETE FROM " & TablaFlujosMD & " WHERE EMISION = '" & matem(i, 1) & "' AND FREGISTRO = " & txtfecha0
Next i
For i = 1 To UBound(mata, 1)
    txtfecha1 = "to_date('" & Format(mata(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(mata(i, 3), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha3 = "to_date('" & Format(mata(i, 4), "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "INSERT INTO " & TablaFlujosMD & " VALUES("
    txtcadena = txtcadena & "'" & mata(i, 1) & "',"     'EMISION
    txtcadena = txtcadena & txtfecha1 & ","             'fecha de registro
    txtcadena = txtcadena & txtfecha2 & ","             'inicio del flujo
    txtcadena = txtcadena & txtfecha3 & ","             'FIN DEL FLUJO
    txtcadena = txtcadena & mata(i, 5) & ","            'NOCIONAL
    txtcadena = txtcadena & mata(i, 6) & ","            'AMORTIZACION
    txtcadena = txtcadena & mata(i, 7) & ","            'tasa
    txtcadena = txtcadena & mata(i, 8) & ")"            'pcupon
    ConAdo.Execute txtcadena
    AvanceProc = i / UBound(mata, 1)
    MensajeProc = "Guardando flujos de MD " & Format(AvanceProc, "#,##0.00 %")
    DoEvents
Next i

End Sub

Sub ImpFRExcel(ByVal nomarch As String, ByVal nomtabla As String)
Dim i As Long
Dim j As Long
Dim noreg As Long
Dim nocampos As Long
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim txtindice As String
Dim txtborra As String
Dim txtinserta As String
Dim txtfecha As String
Dim fecha1 As Date
Dim fecha2 As Date
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim matd() As Variant

Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
Set registros1 = base1.OpenRecordset(nomtabla, dbOpenDynaset, dbReadOnly)
registros1.MoveLast
noreg = registros1.RecordCount
nocampos = registros1.Fields.Count - 1
ReDim matb(1 To noreg, 1 To nocampos) As Variant
ReDim matc(1 To nocampos) As String
ReDim matff(1 To noreg, 1 To 1) As Variant
For i = 1 To nocampos
    matc(i) = registros1.Fields(i).Name
Next i
registros1.MoveFirst
For i = 1 To noreg
    matff(i, 1) = LeerTAccess(registros1, 0, i)        'fecha
    For j = 1 To nocampos
        If Not EsVariableVacia(LeerTAccess(registros1, j, i)) Then
           matb(i, j) = Val(LeerTAccess(registros1, j, i)) 'valor
        Else
           matb(i, j) = 0
        End If
    Next j
 registros1.MoveNext
Next i
registros1.Close
base1.Close
If noreg <> 0 Then
matd = RutinaOrden(matff, 1, SRutOrden)
fecha1 = matd(1, 1)
fecha2 = matd(noreg, 1)
For i = 1 To nocampos
    txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtborra = "DELETE FROM " & TablaFRiesgoO & " WHERE "
    txtborra = txtborra & "FECHA >=  " & txtfecha1 & " AND "
    txtborra = txtborra & "FECHA <=  " & txtfecha2 & " AND "
    txtborra = txtborra & "CONCEPTO=  '" & matc(i) & "'"
    ConAdo.Execute txtborra
    For j = 1 To noreg
        txtindice = CLng(matff(j, 1)) & matc(i) & Format(0, "0000000")
        txtinserta = "INSERT INTO " & TablaFRiesgoO & " VALUES("
        txtfecha = "to_date('" & Format(matff(j, 1), "dd/mm/yyyy") & "','dd/mm/yyyy')"
        txtinserta = txtinserta & txtfecha & ","                'fecha
        txtinserta = txtinserta & "'" & matc(i) & "',"          'descripcion
        txtinserta = txtinserta & 0 & ","                       'plazo
        txtinserta = txtinserta & matb(j, i) & ","              'valor
        txtinserta = txtinserta & "'" & txtindice & "')"        'indice
        ConAdo.Execute txtinserta
    Next j
       If j Mod 50 = 0 Then
         DoEvents
         AvanceProc = j / noreg
         MensajeProc = "Insertando registro " & matff(i) & " " & Format(AvanceProc, "##0.00 %")
       End If
Next i
End If
End Sub

Function LeerCVaRHist(ByVal fecha As Date, ByVal txtport As String, ByVal txtsubport As String, ByVal nconf As Double, ByVal noesc As Integer, ByVal htiempo As Integer)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "' AND SUBPORT = '" & txtsubport & "'"
txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES ='Normal' AND TVAR = 'CVARH' AND NCONF = " & nconf & " AND NOESC = " & noesc & " AND HTIEMPO = " & htiempo
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   rmesa.MoveFirst
   LeerCVaRHist = rmesa.Fields("VALOR")
   rmesa.Close
Else
   LeerCVaRHist = 0
End If
End Function

Sub ImportarPosID(ByVal fecha As Date, ByVal txtnompos As String, ByRef nrn As Integer, ByRef nrc As Integer)
Dim nrn1 As Integer
Dim nrc1 As Integer
Dim nrn2 As Integer
Dim nrc2 As Integer
    Call ImpPosSwapsID(fecha, False, txtnompos, nrn1, nrc1)
    Call ImpFwdID(fecha, False, txtnompos, nrn2, nrc2)
    nrn = nrn1 + nrn2
    nrc = nrc1 + nrc2
End Sub

Sub ValidarOperaciones(ByVal fecha As Date, ByVal coperacion As String, ByVal horareg As Long, ByVal hinicio As Date, ByVal hfinal As Date)
Dim i As Integer
Dim txtcadena As String
Dim txtfecha As String
Dim txthinicio As String
Dim txthfinal As String

    txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txthinicio = "TO_DATE('" & Format(hinicio, "HH:MM:SS") & "','HH24:MI:SS')"
    txthfinal = "TO_DATE('" & Format(hfinal, "HH:MM:SS") & "','HH24:MI:SS')"
    If horareg <> 0 Then
       txtcadena = "INSERT INTO " & TablaOperValidada & " VALUES("
       txtcadena = txtcadena & "'" & coperacion & "',"
       txtcadena = txtcadena & horareg & ","
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & txthinicio & ","
       txtcadena = txtcadena & txthfinal & ")"
       ConAdo.Execute txtcadena
    End If
    
End Sub

Sub ValidarOperaciones2(ByRef mata() As Variant, ByVal finicio As Date, ByVal hinicio As Date, ByVal ffinal As Date, ByVal hfinal As Date)
Dim i As Integer
Dim txtcadena As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txthinicio As String
Dim txthfinal As String

    txtfecha1 = "to_date('" & Format(finicio, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(ffinal, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txthinicio = "TO_DATE('" & Format(hinicio, "HH:MM:SS") & "','HH24:MI:SS')"
    txthfinal = "TO_DATE('" & Format(hfinal, "HH:MM:SS") & "','HH24:MI:SS')"
    If hfinal <> 0 Then
       For i = 1 To UBound(mata, 1)
       txtcadena = "INSERT INTO " & TablaOperValidada & " VALUES("
       txtcadena = txtcadena & "'" & mata(i, 1) & "',"    'clave de operacion
       txtcadena = txtcadena & mata(i, 2) & ","           'hora de registro en ikos
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & txthinicio & ","
       txtcadena = txtcadena & txtfecha2 & ","
       txtcadena = txtcadena & txthfinal & ")"
       ConAdo.Execute txtcadena
       Next i
    End If
    
End Sub

Sub ValidarOperacion3(ByVal coperacion As String, ByVal horareg As String, ByVal finicio As Date, ByVal hinicio As Date, ByVal ffinal As Date, ByVal hfinal As Date)
Dim txtcadena As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txthinicio As String
Dim txthfinal As String
    txtfecha1 = "to_date('" & Format(finicio, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(ffinal, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txthinicio = "TO_DATE('" & Format(hinicio, "HH:MM:SS") & "','HH24:MI:SS')"
    txthfinal = "TO_DATE('" & Format(hfinal, "HH:MM:SS") & "','HH24:MI:SS')"
    If hfinal <> 0 Then
       txtcadena = "INSERT INTO " & TablaOperValidada & " VALUES("
       txtcadena = txtcadena & "'" & coperacion & "',"    'clave de operacion
       txtcadena = txtcadena & horareg & ","           'hora de registro en ikos
       txtcadena = txtcadena & txtfecha1 & ","
       txtcadena = txtcadena & txthinicio & ","
       txtcadena = txtcadena & txtfecha2 & ","
       txtcadena = txtcadena & txthfinal & ")"
       ConAdo.Execute txtcadena
    End If
    
End Sub

Sub ImpReporteEfProsSwap(ByVal fecha As Date, ByRef mata() As Variant, ByVal coperacion As String, ByVal eficpros As Double)
Dim nofval As Integer
Dim nomarch1 As String
Dim contar As Long
Dim txtcadena As String
Dim i As Long
Dim renglon As Long
Dim exitoarch As Boolean

   nofval = UBound(mata, 1)
   nomarch1 = DirResVaR & "\Resumen eficiencia prospectiva swap " & coperacion & " " & Format(fecha, "yyyy-mm-dd") & ".txt"
   frmCalVar.CommonDialog1.FileName = nomarch1
   frmCalVar.CommonDialog1.ShowSave
   nomarch1 = frmCalVar.CommonDialog1.FileName
   Call VerificarSalidaArchivo(nomarch1, 5, exitoarch)
   If exitoarch Then
   contar = 0
   txtcadena = "Fecha" & Chr(9)
   txtcadena = txtcadena & "No. de simulaciones" & Chr(9)
   txtcadena = txtcadena & "No. de aciertos"
      Print #5, txtcadena
      For i = 1 To nofval
          txtcadena = mata(i, 1) & Chr(9) & mata(i, 2) & Chr(9) & mata(i, 3)
          Print #5, txtcadena
      Next i
   Print #5, "Porcentaje de aciertos" & Chr(9) & Format(eficpros, "##0.00 %")
   Close #5
   MsgBox "Se debe de tomar el archivo " & nomarch1 & ", pegarse en un documento de word y enviarse al area de Derivados."
   End If
End Sub

Sub ExtrapolTTIIE(ByVal fecha1 As Date, ByVal fecha2 As Date, ByRef mata() As Variant)
Dim mattiie28() As Variant
Dim mattiie91() As Variant
Dim noreg As Long
Dim i As Long
Dim txtfecha1 As String
Dim finicioem As Date
Dim pc As Integer
Dim cemision As String
Dim concepto As String
'1  TIPO VALOR
'2  EMISION
'3  SERIE
'4  FECHA DE INICIO DE EMISION

'se carga toda la historia disponible de tiie a 28 dias
mattiie28 = Leer1FactorRC("TIIE28 PIP", 0)
mattiie91 = Leer1FactorRC("TIIE91 PIP", 0)
noreg = UBound(mata, 1)
For i = 1 To noreg
    cemision = mata(i, 1) & mata(i, 2) & mata(i, 3)
    concepto = "Y " & mata(i, 2) & mata(i, 3)                   'nombre de la yield
    finicioem = mata(i, 4)
    pc = mata(i, 5)
    st = mata(i, 6)                                             'sobretasa cupon
    If pc = 28 Or pc = 30 Then
       Call ExtrapolaYieldEmTIIE(cemision, concepto, st, fecha1, fecha2, finicioem, mattiie28)
    Else
       Call ExtrapolaYieldEmTIIE(cemision, concepto, st, fecha1, fecha2, finicioem, mattiie91)
    End If
    AvanceProc = i / noreg
    MensajeProc = "Extrapolando " & concepto & " " & Format(AvanceProc, "##0.00 %")
    DoEvents
Next i
End Sub

Sub ExtrapolaYieldEmTIIE(ByVal cemision As String, ByVal concepto As String, ByVal st As Double, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal finicioem As Date, matfr() As Variant)
Dim fecha As Date
Dim fechax As Date
Dim indice As Long
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtborra As String
Dim txtinserta As String
Dim txtindice As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset
Dim noreg As Long
Dim i As Long

If fecha1 < finicioem Then
   fechax = Minimo(fecha2, finicioem)
   txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha2 = "to_date('" & Format(fechax, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtborra = "DELETE FROM " & TablaFRiesgoO & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <  " & txtfecha2 & "  AND CONCEPTO = '" & concepto & "' AND PLAZO = 0"
   ConAdo.Execute txtborra
   fecha = fecha1
   Do While fecha <= fechax
      indice = BuscarValorArray(fecha, matfr, 1)
      If indice <> 0 Then
         txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
         txtindice = CLng(fecha) & concepto & "0000000"
         txtinserta = "INSERT INTO " & TablaFRiesgoO & " VALUES("
         txtinserta = txtinserta & txtfecha & ","                             'FECHA
         txtinserta = txtinserta & "'" & concepto & "',"                      'CONCEPTO
         txtinserta = txtinserta & "0,"                                       'PLAZO
         txtinserta = txtinserta & matfr(indice, 2) + st & ","                'VALOR
         txtinserta = txtinserta & "'" & txtindice & "')"                     'INDICE
         ConAdo.Execute txtinserta
      End If
      fecha = fecha + 1
   Loop
   txtfiltro2 = "SELECT FECHA,YIELD FROM " & TablaVecPrecios & " WHERE CLAVE_EMISION = '" & cemision & "' ORDER BY FECHA"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      txtfecha1 = "to_date('" & Format(fechax, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtborra = "DELETE FROM " & TablaFRiesgoO & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & "  AND CONCEPTO = '" & concepto & "' AND PLAZO = 0"
      ConAdo.Execute txtborra
      ReDim mata(1 To noreg, 1 To 2) As Variant
      rmesa.Open txtfiltro2, ConAdo
      For i = 1 To noreg
          mata(i, 1) = rmesa.Fields(0)
          mata(i, 2) = CDbl(rmesa.Fields(1)) / 100
          rmesa.MoveNext
      Next i
      rmesa.Close
      fecha = fechax
      Do While fecha <= fecha2
         indice = BuscarValorArray(fecha, mata, 1)
         If indice <> 0 Then
            txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
            txtindice = CLng(fecha) & concepto & "0000000"
            txtinserta = "INSERT INTO " & TablaFRiesgoO & " VALUES("
            txtinserta = txtinserta & txtfecha & ","                             'FECHA
            txtinserta = txtinserta & "'" & concepto & "',"                      'CONCEPTO
            txtinserta = txtinserta & "0,"                                       'PLAZO
            txtinserta = txtinserta & mata(indice, 2) & ","                      'VALOR
            txtinserta = txtinserta & "'" & txtindice & "')"                     'INDICE
            ConAdo.Execute txtinserta
         End If
         fecha = fecha + 1
      Loop
   End If
End If
End Sub

Sub ExtrapolTTIIE2(ByVal fecha As Date, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txttv As String, ByVal txtemision As String, ByVal txtserie As String)
Dim mattiie() As Variant
Dim noreg1 As Long
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtborra1 As String
Dim txtindice As String
Dim txtinserta1 As String
Dim txtfiltro As String
Dim fechax As Date
Dim indice As Long
Dim i As Long
Dim vst As Double
Dim txtconcepto As String
Dim finicio As Date

Dim esnulo As Boolean

'se carga toda la historia disponible de tiie a 28 dias
mattiie = Leer1FRiesgoxVaR(fecha1, fecha2, "TIIE28 PIP", 0, esnulo)
noreg1 = UBound(MatTValSTCupon, 1)
For i = 1 To noreg1
    If txttv = MatTValSTCupon(i, 1) And txtemision = MatTValSTCupon(i, 2) And txtserie = MatTValSTCupon(i, 3) Then
       finicio = MatTValSTCupon(i, 5)
       vst = MatTValSTCupon(i, 8)
       Exit For
    End If
Next i
txtconcepto = "Y " & txtemision & txtserie                         'nombre de la yield
fechax = fecha1
Do While fechax < finicio
   indice = BuscarValorArray(fechax, mattiie, 1)
   If indice <> 0 Then
      txtfecha1 = "to_date('" & Format(fechax, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtborra1 = "DELETE FROM " & TablaFRiesgoO & " WHERE FECHA = " & txtfecha1
      txtborra1 = txtborra1 & " AND CONCEPTO = '" & txtconcepto & "' AND PLAZO = 0"
      txtindice = CLng(fechax) & txtconcepto & "0000000"
      txtinserta1 = "INSERT INTO " & TablaFRiesgoO & " VALUES("
      txtinserta1 = txtinserta1 & txtfecha1 & ","                         'FECHA
      txtinserta1 = txtinserta1 & "'" & txtconcepto & "',"                'CONCEPTO
      txtinserta1 = txtinserta1 & "0,"                                    'PLAZO
      txtinserta1 = txtinserta1 & mattiie(indice, 2) + vst & ","          'VALOR
      txtinserta1 = txtinserta1 & "'" & txtindice & "')"                  'INDICE
      ConAdo.Execute txtborra1
      ConAdo.Execute txtinserta1
      AvanceProc = (fechax - fecha1 + 1) / (fecha2 - fecha1 + 1)
      MensajeProc = "Extrapolando " & txtconcepto & " " & Format(fechax, "dd/mm/yyyy") & " " & Format(AvanceProc, "##0.00 %")
      DoEvents
   End If
   fechax = fechax + 1
Loop
End Sub


Sub ActListaBonosRefTIIE(ByVal fecha As Date, ByVal txttref As String, ByRef noreg As Long, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txttv As String
Dim txtemision As String
Dim txtserie As String
Dim vst As Double
Dim txtydesc As String
Dim txtcadena As String
Dim i As Long
Dim valor1 As Integer
Dim valor2 As String
Dim noreg1 As Long
Dim rmesa As New ADODB.recordset

exito = False
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro1 = "SELECT COUNT(*) FROM " & TablaVecPrecios & " WHERE "
txtfiltro1 = txtfiltro1 & " FECHA = " & txtfecha
txtfiltro1 = txtfiltro1 & " AND ST_COLOCACION <> 0"
rmesa.Open txtfiltro1, ConAdo
noreg1 = rmesa.Fields(0)
rmesa.Close
If noreg1 = 0 Then
   txtmsg = "No hay elemento con sobretasa no nula para esta fecha"
   Exit Sub
End If
txtfiltro2 = "SELECT * FROM " & TablaVecPrecios & " WHERE "
txtfiltro2 = txtfiltro2 & " FECHA = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND REGLA_CUPON = '" & txttref & "'"
txtfiltro2 = txtfiltro2 & " AND (TV,EMISION,SERIE) NOT IN("
txtfiltro2 = txtfiltro2 & "SELECT TIPO_VALOR,EMISORA,SERIE FROM " & PrefijoBD & TablaValBSC & ")"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       If txttref = "TIIE 28 dias" Then
          valor1 = 28
          valor2 = "TIIE28 PIP"
       Else
         valor1 = 91
         valor2 = "TIIE91 PIP"
       End If
       txttv = rmesa.Fields("TV")
       txtemision = rmesa.Fields("EMISION")
       txtserie = rmesa.Fields("SERIE")
       vst = ConvValor(rmesa.Fields("ST_COLOCACION")) / 100
       txtydesc = "Y " & txtemision & txtserie
       txtcadena = "INSERT INTO " & PrefijoBD & TablaValBSC & " VALUES("
       txtcadena = txtcadena & "'" & txttv & "',"             'tipo valor
       txtcadena = txtcadena & "'" & txtemision & "',"        'emision
       txtcadena = txtcadena & "'" & txtserie & "',"          'serie
       txtcadena = txtcadena & "'BONO STC Y',"                'funcion de valuacion
       txtcadena = txtcadena & valor1 & ","                   'periodo cupon
       txtcadena = txtcadena & vst & ","                      'sobretasa
       txtcadena = txtcadena & "'" & txtydesc & "',"          'factor de descuento
       txtcadena = txtcadena & "1,"                           'interpolacion
       txtcadena = txtcadena & "'" & valor2 & "',"            'tasa de referencia
       txtcadena = txtcadena & "1)"                           'interpolacion
       ConAdo.Execute txtcadena
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Actualizando la tabla " & PrefijoBD & TablaValBSC & " " & Format(AvanceProc, "##0.00 %")
   Next i
   rmesa.Close
   txtmsg = "El proceso finalizo correctamente"
   exito = True
End If
End Sub

Function Leer1FRiesgoxVaR(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal txtconcepto As String, ByVal plazo As Long, ByRef esnulo As Boolean) As Variant()
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset

esnulo = False
txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT " & TablaFechasVaR & ".FECHA,A.VALOR"
txtfiltro2 = txtfiltro2 & " FROM " & TablaFechasVaR
txtfiltro2 = txtfiltro2 & " LEFT JOIN (SELECT * FROM " & TablaFRiesgoO
txtfiltro2 = txtfiltro2 & " WHERE CONCEPTO = '" & txtconcepto & "'"
txtfiltro2 = txtfiltro2 & " AND PLAZO = " & plazo & ") A"
txtfiltro2 = txtfiltro2 & " ON " & TablaFechasVaR & ".FECHA = A.FECHA"
txtfiltro2 = txtfiltro2 & " WHERE " & TablaFechasVaR & ".FECHA >= " & txtfecha1
txtfiltro2 = txtfiltro2 & " AND " & TablaFechasVaR & ".FECHA <= " & txtfecha2
txtfiltro2 = txtfiltro2 & " ORDER BY " & TablaFechasVaR & ".FECHA"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim mata(1 To noreg, 1 To 2) As Variant
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("FECHA")
       If Not IsNull(rmesa.Fields("VALOR")) Then
          mata(i, 2) = rmesa.Fields("VALOR")
       Else
          mata(i, 2) = 0
          esnulo = True
       End If
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Variant
End If
Leer1FRiesgoxVaR = mata
End Function

Sub DeterminaYieldsTIIE(ByVal fecha As Date, ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim txtfecha As String
Dim txtfiltro As String
Dim exito1 As Boolean
Dim matpos() As propPosMD
Dim i As Long
Dim esnulo As Boolean
Dim mata() As Variant
   mata = Leer1FRiesgoxVaR(fecha1, fecha2, "TIIE28 PIP", 0, esnulo)
   If esnulo Then MsgBox "hay valores nulos en la tiie 28"
   mata = Leer1FRiesgoxVaR(fecha1, fecha2, "TIIE91 PIP", 0, esnulo)
   If esnulo Then MsgBox "hay valores nulos en la tiie 91"
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro = "SELECT * FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
   txtfiltro = txtfiltro & " AND (TOPERACION =1 OR TOPERACION =4)"
   matpos = LeerBaseMD(txtfiltro)
   'Call ClasTProdMD(matpos, MatPosMD, exito1)
   If UBound(matpos, 1) <> 0 Then
   For i = 1 To UBound(matpos, 1)
       If matpos(i).fValuacion = "BONO STC Y" Then
         mata = Leer1FRiesgoxVaR(fecha1, fecha2, matpos(i).fRiesgo1MD, 0, esnulo)
         If esnulo Then
           MsgBox "HAY valores nulos para el factor" & matpos(i).fRiesgo1MD
           Call ExtrapolTTIIE2(fecha, fecha1, fecha2, matpos(i).tValorMD, matpos(i).emisionMD, matpos(i).serieMD)
         End If
       End If
   Next i
   End If
End Sub

Function FechasProcG(ByVal opcion As Integer)
Dim txtfiltro As String
Dim txtfiltro1 As String
Dim noreg As Integer
Dim i As Integer
Dim txttabla As String
Dim rmesa As New ADODB.recordset
If opcion = 1 Then
   txttabla = TablaProcesos1
ElseIf opcion = 2 Then
   txttabla = TablaProcesos2
End If
txtfiltro = "SELECT FECHAP FROM " & txttabla & " GROUP BY FECHAP ORDER BY FECHAP DESC"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro, ConAdo
   ReDim mata(1 To noreg, 1 To 1) As Date
   rmesa.MoveFirst
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields(0)
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
   ReDim mata(0 To 0, 0 To 0) As Date
End If
FechasProcG = mata
End Function


Function LeerFechasEsc(ByVal nomarch As String)
    Dim txtnomarch As String
    Dim base1      As DAO.Database
    Dim i          As Integer, noreg As Integer
    Dim registros1 As DAO.recordset

    Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
    Set registros1 = base1.OpenRecordset("Fechas$", dbOpenDynaset, dbReadOnly)
    'se cargan las dos curvas necesarias para este proceso
    If registros1.RecordCount <> 0 Then
        registros1.MoveLast
        noreg = registros1.RecordCount
        ReDim mata(1 To noreg, 1 To 1) As Date
        registros1.MoveFirst
        For i = 1 To noreg
            mata(i, 1) = LeerTAccess(registros1, 0, i)        'fecha
            registros1.MoveNext
        Next i

        registros1.Close
        base1.Close
        LeerFechasEsc = mata
    End If

End Function

Function LeerEscEstres(ByVal nomarch As String)
    Dim txtnomarch As String
    Dim base1      As DAO.Database
    Dim i          As Integer, noreg As Integer
    Dim registros1 As DAO.recordset
   
    Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
    Set registros1 = base1.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)

    'se cargan las dos curvas necesarias para este proceso
    If registros1.RecordCount <> 0 Then
        registros1.MoveLast
        noreg = registros1.RecordCount
        ReDim mata(1 To noreg, 1 To 7) As Variant
        registros1.MoveFirst

        For i = 1 To noreg
            mata(i, 1) = LeerTAccess(registros1, 0, i)        'nombre del escenario
            mata(i, 2) = Val(LeerTAccess(registros1, 1, i))   'curvas
            mata(i, 3) = Val(LeerTAccess(registros1, 2, i))   'tasas de referencia
            mata(i, 4) = Val(LeerTAccess(registros1, 3, i))   'yield is
            mata(i, 5) = Val(LeerTAccess(registros1, 4, i))   'yields
            mata(i, 6) = Val(LeerTAccess(registros1, 5, i))   'sobretasas
            mata(i, 7) = Val(LeerTAccess(registros1, 6, i))   'tipos de cambio
            registros1.MoveNext
        Next i

        registros1.Close
        base1.Close
        LeerEscEstres = mata
    End If

End Function

Function LeerPosMDExcel(ByVal nomarch As String) As Variant()
If ActivarControlErrores Then
On Error GoTo ControlErrores
End If
Dim sihayarch As Boolean
Dim i As Long
Dim j As Integer
Dim mata() As Variant
Dim base1 As DAO.Database
Dim registros1 As DAO.recordset
Dim noreg As Long
Dim nocampos As Integer

sihayarch = VerifAccesoArch(nomarch)
If sihayarch Then
   Set base1 = OpenDatabase(nomarch, dbDriverNoPrompt, False, VersExcel)
   Set registros1 = base1.OpenRecordset("Hoja1$", dbOpenDynaset, dbReadOnly)
   If registros1.RecordCount <> 0 Then
      registros1.MoveLast
      noreg = registros1.RecordCount
      nocampos = registros1.Fields.Count
      registros1.MoveFirst
      ReDim mata(1 To noreg, 1 To nocampos) As Variant
      For i = 1 To noreg
          For j = 1 To nocampos
              mata(i, j) = registros1.Fields(j - 1)
          Next j
          registros1.MoveNext
       Next i
   Else
     ReDim mata(0 To 0, 0 To 0) As Variant
   End If
   registros1.Close
   base1.Close
Else
    ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerPosMDExcel = mata
On Error GoTo 0
Exit Function
ControlErrores:
   ReDim mata(0 To 0, 0 To 0) As Variant
   LeerPosMDExcel = mata

End Function

Sub ExportarPosFwdDiaAnterior(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim siesfv As Boolean
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtfecha4 As String
Dim txtfecha5 As String
Dim txtfiltro As String
Dim txtcadena As String
Dim noreg As Long
Dim nocampos As Long
Dim i As Long
Dim j As Long
Dim txtborra As String
Dim rmesa As New ADODB.recordset


siesfv = EsFechaVaR(fecha1)
If siesfv Then
   txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro = TablaPosFwd & " WHERE FECHAREG = " & txtfecha1 & " AND FVENCIMIENTO > " & txtfecha2
   txtcadena = "SELECT COUNT(*) FROM " & txtfiltro
   rmesa.Open txtcadena, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      txtcadena = "SELECT * FROM " & txtfiltro
      rmesa.Open txtcadena, ConAdo
      nocampos = rmesa.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos) As Variant
      rmesa.MoveFirst
      For i = 1 To noreg
          For j = 1 To nocampos
              mata(i, j) = rmesa.Fields(j - 1)
          Next j
          rmesa.MoveNext
      Next i
      rmesa.Close
      txtborra = "DELETE FROM " & TablaPosFwd & " WHERE FECHAREG = " & txtfecha2
      ConAdo.Execute txtborra
      For i = 1 To noreg
          txtfecha3 = "to_date('" & Format(mata(i, 12), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtfecha4 = "to_date('" & Format(mata(i, 13), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtfecha5 = "to_date('" & Format(mata(i, 14), "dd/mm/yyyy") & "','dd/mm/yyyy')"
          txtcadena = "INSERT INTO " & TablaPosFwd & " VALUES("
          txtcadena = txtcadena & mata(i, 1) & ","
          txtcadena = txtcadena & txtfecha2 & ","
          txtcadena = txtcadena & "'Real',"
          txtcadena = txtcadena & "'" & mata(i, 4) & "',"
          txtcadena = txtcadena & "'" & mata(i, 5) & "',"
          txtcadena = txtcadena & "'" & mata(i, 6) & "',"
          txtcadena = txtcadena & "'" & mata(i, 7) & "',"
          txtcadena = txtcadena & "'" & mata(i, 8) & "',"
          txtcadena = txtcadena & mata(i, 9) & ","
          txtcadena = txtcadena & mata(i, 10) & ","
          txtcadena = txtcadena & mata(i, 11) & ","
          txtcadena = txtcadena & txtfecha3 & ","
          txtcadena = txtcadena & txtfecha4 & ","
          txtcadena = txtcadena & txtfecha5 & ","
          txtcadena = txtcadena & mata(i, 15) & ","
          txtcadena = txtcadena & "'" & mata(i, 16) & "',"
          txtcadena = txtcadena & "'" & mata(i, 17) & "')"
          ConAdo.Execute txtcadena
       Next i
   End If
End If
End Sub

Sub RepDetalleCLimiteC(ByVal fecha As Date, ByVal coperacion As String, ByVal tabla As String)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Long
Dim noreg1 As Long
Dim noreg2 As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim txtcadena As String
Dim matb() As String
Dim matc() As String
Dim mate() As String
Dim matf() As String
Dim matg() As String
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & tabla & " WHERE FECHA = " & txtfecha & " AND  COPERACION = '" & coperacion & "' ORDER BY ID_GRUPOC"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   If tabla = TablaLimContrap1 Then
      Open "d:\simulaciones " & coperacion & " " & Format(fecha, "yyyy-mm-dd") & " max exp 1.txt" For Output As #1
   Else
      Open "d:\simulaciones " & coperacion & " " & Format(fecha, "yyyy-mm-dd") & " max exp 2.txt" For Output As #1
   End If
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 6) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("ID_GRUPOC")
       mata(i, 2) = rmesa.Fields("H_FECHAS1")
       mata(i, 3) = rmesa.Fields("H_FECHAS2")
       mata(i, 4) = rmesa.Fields("H_VALMAX")
       mata(i, 5) = rmesa.Fields("H_VALMAXACT")
       mata(i, 6) = rmesa.Fields("H_VALMAXPAS")
       matb = EncontrarSubCadenas(mata(i, 2), ",")
       matc = EncontrarSubCadenas(mata(i, 3), ",")
       mate = EncontrarSubCadenas(mata(i, 4), ",")
       matf = EncontrarSubCadenas(mata(i, 5), ",")
       matg = EncontrarSubCadenas(mata(i, 6), ",")
       noreg1 = UBound(matb, 1)
       noreg2 = UBound(matc, 1)
       Print #1, "Grupo de calculo " & Chr(9) & mata(i, 1)
       Print #1, "MTM"
       txtcadena = "Fecha" & Chr(9)
       For j = 1 To noreg2
           txtcadena = txtcadena & matc(j) & Chr(9)
       Next j
       Print #1, txtcadena
       For j = 1 To noreg1
           txtcadena = matb(j) & Chr(9)
           For k = 1 To noreg2
               txtcadena = txtcadena & mate((j - 1) * noreg2 + k) & Chr(9)
           Next k
           Print #1, txtcadena
       Next j
       
       Print #1, "val act"
       txtcadena = "Fecha" & Chr(9)
       For j = 1 To noreg2
           txtcadena = txtcadena & matc(j) & Chr(9)
       Next j
       Print #1, txtcadena
       For j = 1 To noreg1
           txtcadena = matb(j) & Chr(9)
           For k = 1 To noreg2
               txtcadena = txtcadena & matf((j - 1) * noreg2 + k) & Chr(9)
           Next k
           Print #1, txtcadena
       Next j

       Print #1, "val pas"
       txtcadena = "Fecha" & Chr(9)
       For j = 1 To noreg2
           txtcadena = txtcadena & matc(j) & Chr(9)
       Next j
       Print #1, txtcadena
       For j = 1 To noreg1
           txtcadena = matb(j) & Chr(9)
           For k = 1 To noreg2
               txtcadena = txtcadena & matg((j - 1) * noreg2 + k) & Chr(9)
           Next k
           Print #1, txtcadena
       Next j
       
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Extrayendo la informacion " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   rmesa.Close
   Close #1
End If
End Sub

Sub RepDetalleCLimiteC2(ByVal fecha As Date, ByVal coperacion As String, ByVal tabla As String)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Long
Dim noreg1 As Long
Dim noreg2 As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim txtcadena As String
Dim matb() As String
Dim matc() As String
Dim mate() As String
Dim matf() As String
Dim matg() As String
Dim mata() As Variant
Dim matr() As Variant
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & tabla & " WHERE FECHA = " & txtfecha & " AND  COPERACION = '" & coperacion & "' ORDER BY ID_GRUPOC"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close

If noreg <> 0 Then
   If tabla = TablaLimContrap1 Then
      Open "d:\simulaciones " & coperacion & " " & Format(fecha, "yyyy-mm-dd") & " max exp 1.txt" For Output As #1
   Else
      Open "d:\simulaciones " & coperacion & " " & Format(fecha, "yyyy-mm-dd") & " max exp 2.txt" For Output As #1
   End If
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 10) As Variant
   ReDim matmtm(1 To noreg + 2, 1 To 100 + 1) As Variant
   ReDim matva(1 To noreg + 2, 1 To 100 + 1) As Variant
   ReDim matvp(1 To noreg + 2, 1 To 100 + 1) As Variant
   ReDim mats1(0 To 0, 0 To 0) As Variant
   ReDim mats2(0 To 0, 0 To 0) As Variant
   ReDim mats3(0 To 0, 0 To 0) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("ID_GRUPOC")     'grupo de calculos
       mata(i, 2) = rmesa.Fields("H_FECHAS1")     'fechas de escenarios
       mata(i, 3) = rmesa.Fields("H_FECHAS2")     'fechas de valuacion
       mata(i, 4) = rmesa.Fields("H_VALMAX")      'mtm futuro
       mata(i, 5) = rmesa.Fields("H_VALMAXACT")   'val act futura
       mata(i, 6) = rmesa.Fields("H_VALMAXPAS")   'val pas futura
       matb = EncontrarSubCadenas(mata(i, 2), ",")
       matc = EncontrarSubCadenas(mata(i, 3), ",")
       mate = EncontrarSubCadenas(mata(i, 4), ",")
       matf = EncontrarSubCadenas(mata(i, 5), ",")
       matg = EncontrarSubCadenas(mata(i, 6), ",")
       
       noreg1 = UBound(matb, 1)
       noreg2 = UBound(matc, 1)
       For j = 1 To noreg2
           matr(1, j) = matc(j)
       Next j
       For j = 1 To noreg1
           matmtm(j, 1) = matb(j)
           For k = 1 To noreg2
               matmtm(j, k) = mate((j - 1) * noreg2 + k)
           Next k
       Next j
       For j = 1 To noreg2
           matva(1, j) = matc(j)
       Next j
       For j = 1 To noreg1
           matva(j, 1) = matb(j)
           For k = 1 To noreg2
               matva(j, k) = matf((j - 1) * noreg2 + k)
           Next k
           Print #1, txtcadena
       Next j
       
       For j = 1 To noreg2
           matvp(1, j) = matc(j)
       Next j
       For j = 1 To noreg1
           matva(j, 1) = matb(j)
           For k = 1 To noreg2
               matva(j, k) = matg((j - 1) * noreg2 + k)
           Next k
           Print #1, txtcadena
       Next j
       If UBound(mats1, 1) <> 0 Then
          mats1 = UnirMatrices(mats1, matr, 1)
       Else
          mats1 = matr
       End If
       If UBound(mats2, 1) <> 0 Then
          mats2 = UnirMatrices(mats2, matva, 1)
       Else
          mats2 = matva
       End If
       If UBound(mats3, 1) <> 0 Then
          mats3 = UnirMatrices(mats3, matvp, 1)
       Else
          mats3 = matvp
       End If
      
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Extrayendo la informacion " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   rmesa.Close
   Close #1
End If
End Sub

Sub ProcEficRetro(ByVal fecha As Date, ByRef txtmsg As String, ByRef exito As Boolean)
Dim fecha1 As Date
   If esFinMes(fecha) Then
      Call GenSubprocEficRetro(fecha, 1, txtmsg, exito)
   Else
      fecha1 = PBD1(fecha, 1, "MX")
      Call CopiarEficRetro(fecha1, fecha, txtmsg, exito)
   End If
End Sub


Sub GenSubprocEficRetro(ByVal fecha As Date, ByVal id_tabla As Integer, ByRef txtmsg As String, ByRef exito As Boolean)
   Dim txtfiltro1 As String
   Dim txtfiltro2 As String
   Dim txtfecha As String
   Dim noreg As Long
   Dim i As Long
   Dim j As Long
   Dim tipo_efic As Integer
   Dim rmesa As New ADODB.recordset
   Dim fecha0 As Date
   Dim contar As Long
   Dim coperacion As String
   Dim pactiva As String
   Dim ppasiva As String
   Dim pswap As String
   Dim txtport As String
   Dim finicio As Date
   Dim ffinal As Date
   Dim txtcadena As String
   Dim txtborra As String
   Dim indice As Long
   exito = False
   txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
   txtfiltro2 = "SELECT * FROM " & TablaPosFwd
   txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND INTENCION ='C'"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg = 0 Then
      Call ExportarPosFwdDiaAnterior(fecha - 1, fecha)
   End If
   fecha0 = DateSerial(Year(fecha), Month(fecha), 1) - 1
   fecha0 = BuscarFechaNBFR(fecha0)
   txtborra = "DELETE FROM " & TablaEficRetro & " WHERE FECHA = " & txtfecha
   ConAdo.Execute txtborra
   txtfiltro2 = "SELECT * FROM " & TablaPosFwd
   txtfiltro2 = txtfiltro2 & " WHERE FECHAREG = " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND INTENCION ='C'"
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      contar = DeterminaMaxRegSubproc(id_tabla)
      rmesa.Open txtfiltro2, ConAdo
      For i = 1 To noreg
          contar = contar + 1
          coperacion = rmesa.Fields("COPERACION")        'clave de operacion
          pactiva = ""
          ppasiva = ""
          pswap = ""
          txtport = "Efect retro oper " & coperacion
          Call GenPortEfect(fecha, coperacion, pactiva, ppasiva, pswap, txtport)
          txtcadena = CrearCadInsSub(fecha, 56, contar, "efectividad retro forward", fecha0, fecha, txtport, "", "", "", "", "", "", "", "", "", id_tabla)
          ConAdo.Execute txtcadena
          rmesa.MoveNext
      Next i
      rmesa.Close
   End If
   txtfiltro2 = "SELECT * FROM " & TablaPosSwaps
   txtfiltro2 = txtfiltro2 & " WHERE (FECHAREG,COPERACION) IN("
   txtfiltro2 = txtfiltro2 & " SELECT MAX(FECHAREG) AS FECHAREG,COPERACION FROM "
   txtfiltro2 = txtfiltro2 & TablaPosSwaps & " WHERE FECHAREG <= " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND TIPOPOS =1 AND CPOSICION = " & ClavePosDeriv
   txtfiltro2 = txtfiltro2 & " GROUP BY COPERACION) AND FINICIO <= " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND FVENCIMIENTO > " & txtfecha
   txtfiltro2 = txtfiltro2 & " AND INTENCION ='C' AND TIPOPOS = 1 AND CPOSICION = " & ClavePosDeriv
   txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
   rmesa.Open txtfiltro1, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg <> 0 Then
      contar = DeterminaMaxRegSubproc(id_tabla)
      rmesa.Open txtfiltro2, ConAdo
      For i = 1 To noreg
          coperacion = rmesa.Fields("COPERACION")        'clave de operacion
          indice = 0
          For j = 1 To UBound(MatRelSwapsPrim, 1)
             If MatRelSwapsPrim(j).coperacion = coperacion Then
                indice = j
                Exit For
             End If
          Next j
          If indice <> 0 Then
             pactiva = ReemplazaVacioValor(MatRelSwapsPrim(indice).c_ppactiva, "")          'pos primaria activa
             ppasiva = ReemplazaVacioValor(MatRelSwapsPrim(indice).c_pppasiva, "")          'pos primaria pasiva
             pswap = ReemplazaVacioValor(MatRelSwapsPrim(indice).c_pswap, "")               'proxy swap
             tipo_efic = MatRelSwapsPrim(indice).t_efect                                    'tipo de calculo de efectividad
             finicio = MatRelSwapsPrim(indice).finicio
             ffinal = MatRelSwapsPrim(indice).ffin
             txtport = "Efect retro oper " & coperacion
             Call GenPortEfect(fecha, coperacion, pactiva, ppasiva, pswap, txtport)
             If Not EsVariableVacia(txtport) Then
                contar = contar + 1
                If tipo_efic = 1 Then
                   txtcadena = CrearCadInsSub(fecha, 53, contar, "efectividad retro pasivo", fecha0, fecha, txtport, "", "", "", "", "", "", "", "", "", id_tabla)
                ElseIf tipo_efic = 2 Then
                   txtcadena = CrearCadInsSub(fecha, 54, contar, "efectividad retro activo", fecha0, fecha, txtport, "", "", "", "", "", "", "", "", "", id_tabla)
                ElseIf tipo_efic = 3 Then
                   txtcadena = CrearCadInsSub(fecha, 55, contar, "efectividad retro activa-pasiva", fecha0, fecha, txtport, "", "", "", "", "", "", "", "", "", id_tabla)
                ElseIf tipo_efic = 5 Then
                   txtcadena = CrearCadInsSub(fecha, 57, contar, "efectividad retro proxy swap", fecha0, fecha, txtport, "", "", "", "", "", "", "", "", "", id_tabla)
                End If
                ConAdo.Execute txtcadena
             End If
          Else
            MsgBox "No se definio la eficiencia de la operacion " & coperacion
          End If
          rmesa.MoveNext
      Next i
      rmesa.Close
      txtmsg = "El proceso finalizo correctamente"
      exito = True
   End If
End Sub

Sub GenPortEfect(ByVal fecha As Date, ByVal c_operacion As String, pactiva, ByVal ppasiva As String, ByVal pswap As String, ByVal txtport As String)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim tipopos As Integer
Dim fechareg As Date
Dim txtnompos As String
Dim horareg As String
Dim cposicion As Integer
Dim coperacion As String
Dim txtnomport As String
Dim rmesa As New ADODB.recordset
Dim noreg As Long
Dim i As Long
Dim txtborra As String
Dim txtinserta As String

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosSwaps
txtfiltro2 = txtfiltro2 & " WHERE TIPOPOS = 1 AND COPERACION = '" & c_operacion & "' AND FECHAREG IN "
txtfiltro2 = txtfiltro2 & "(SELECT MAX(FECHAREG) FROM " & TablaPosSwaps & " WHERE COPERACION = '" & c_operacion & "'"
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND FECHAREG <= " & txtfecha & ")"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosFwd
txtfiltro2 = txtfiltro2 & " WHERE TIPOPOS = 1 AND COPERACION = '" & c_operacion & "'"
txtfiltro2 = txtfiltro2 & " AND FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDeuda
txtfiltro2 = txtfiltro2 & " WHERE TIPOPOS = 1 AND COPERACION = '" & pactiva & "' AND FECHAREG IN "
txtfiltro2 = txtfiltro2 & "(SELECT MAX(FECHAREG) FROM " & TablaPosDeuda & " WHERE COPERACION = '" & pactiva & "'"
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND FECHAREG <= " & txtfecha & ")"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosDeuda
txtfiltro2 = txtfiltro2 & " WHERE TIPOPOS = 1 AND COPERACION = '" & ppasiva & "' AND FECHAREG IN "
txtfiltro2 = txtfiltro2 & "(SELECT MAX(FECHAREG) FROM " & TablaPosDeuda & " WHERE COPERACION = '" & ppasiva & "'"
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND FECHAREG <= " & txtfecha & ")"
txtfiltro2 = txtfiltro2 & " UNION "
txtfiltro2 = txtfiltro2 & "SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPosSwaps
txtfiltro2 = txtfiltro2 & " WHERE TIPOPOS = 1 AND COPERACION = '" & pswap & "' AND FECHAREG IN "
txtfiltro2 = txtfiltro2 & "(SELECT MAX(FECHAREG) FROM " & TablaPosSwaps & " WHERE COPERACION = '" & pswap & "'"
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND FECHAREG <= " & txtfecha & ")"

txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   txtborra = "DELETE FROM " & TablaPortPosicion & " WHERE FECHA_PORT = " & txtfecha
   txtborra = txtborra & " AND PORTAFOLIO = '" & txtport & "'"
   ConAdo.Execute txtborra
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
      tipopos = rmesa.Fields("TIPOPOS")
      fechareg = rmesa.Fields("FECHAREG")
      txtnompos = rmesa.Fields("NOMPOS")
      horareg = rmesa.Fields("HORAREG")
      cposicion = rmesa.Fields("CPOSICION")
      coperacion = rmesa.Fields("COPERACION")
      txtinserta = "INSERT INTO " & TablaPortPosicion & " VALUES("
      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtinserta = txtinserta & txtfecha & ","
      txtinserta = txtinserta & "'" & txtport & "',"
      txtinserta = txtinserta & tipopos & ","
      txtfecha = "to_date('" & Format(fechareg, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtinserta = txtinserta & txtfecha & ","
      txtinserta = txtinserta & "'" & txtnompos & "',"
      txtinserta = txtinserta & "'" & horareg & "',"
      txtinserta = txtinserta & cposicion & ","
      txtinserta = txtinserta & "'" & coperacion & "')"
      ConAdo.Execute txtinserta
      rmesa.MoveNext
   Next i
   rmesa.Close
Else
txtport = ""
End If
End Sub

Sub CalcEficRetroAct(ByVal fecha As Date, ByRef txtmsg As String, ByRef exito As Boolean)

   Dim mattxt() As String
   Dim matem() As Variant
   Dim matrelp() As New propRelSwapPrim
   Dim bl_exito As Boolean
   Dim fecha0 As Date
   Dim fecha1 As Date
   Dim f_factor As Date
   Dim indice As Integer
   Dim txtport As String
   Dim matpos() As New propPosRiesgo
   Dim matposmd() As New propPosMD
   Dim matposdiv() As New propPosDiv
   Dim matposswaps() As New propPosSwaps
   Dim matposfwd() As New propPosFwd
   Dim matposdeuda() As New propPosDeuda
   Dim matflswap() As New estFlujosDeuda
   Dim matfldeuda() As New estFlujosDeuda
   Dim f_pos As Date
   Dim f_val As Date
   Dim exito0 As Boolean
   Dim exito1 As Boolean
   Dim exito2 As Boolean
   Dim txtmsg1 As String
   Dim txtmsg2 As String
   Dim txtfecha As String
   Dim txtfiltro As String
   Dim noreg As Integer
   Dim txtmsg0 As String
   Dim rmesa As New ADODB.recordset
   
   ValEficiencia = True
   ValExacta = True
   f_factor = BuscarFechaNBFR(fecha)
   f_pos = f_factor
   f_val = fecha
   fecha0 = DateSerial(Year(fecha), Month(fecha), 1) - 1
   fecha0 = BuscarFechaNBFR(fecha0)
   
   'se cargan los datos solo para las fechas existentes
   Call CrearMatFRiesgo2(fecha0, f_factor, MatFactRiesgo, "", exito)
   txtport = "SWAPS DE COBERTURA Y PRIMARIAS"
   Call DeterminaPosSwapsCobyPrim(f_pos, f_pos, "SWAPS DE COBERTURA Y PRIMARIAS", 1)
'SECUENCIA DE carga y valuacion de posicion
   mattxt = CrearFiltroPosPort(f_pos, txtport)
   Call LeerPosBDatos(mattxt, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, txtmsg0, bl_exito)
   If UBound(matpos, 1) > 0 Then
      Call DefinirParValPos(matpos, matposmd, matposdiv, matposswaps, matposfwd, matposdeuda, 1, txtmsg1, exito0)
      If exito0 Then
         matrelp = FiltrarRelPrimFecha(fecha, MatRelSwapsPrim)
         Call CalculoEfRetroSwap(fecha0, f_factor, matpos, matposmd, matposdiv, matposswaps, matposfwd, matflswap, matposdeuda, matfldeuda, MatResEficSwaps, matrelp)
         Call VerificarSwapIneficiente(fecha, matpos, MatResEficSwaps)
         'Call GuardarResEfRetroSwap2(fecha, MatResEficSwaps)
         Call InConexOracle("alm2", conAdo2)
         Call GuardaResEfRetroSwaps(fecha, fecha0, fecha, MatResEficSwaps, conAdo2)
         conAdo2.Close
         exito1 = True
      Else
         exito1 = False
       End If
      
   End If
   If exito Then
      txtmsg = "El proceso finalizo correctamente"
   Else
      txtmsg = txtmsg1 & " " & txtmsg2
   End If
   ValEficiencia = False
    
End Sub


Sub CopiarEficRetro(ByVal dtfecha1 As Date, ByVal dtfecha2 As Date, ByRef txtmsg As String, ByRef exito As Boolean)
Dim txtfecha As String
Dim txtfiltro As String
Dim txtfechaf As String
Dim txtfechac1 As String
Dim txtfechac2 As String
Dim txtfechac3 As String
Dim txtcadena As String
Dim noreg As Integer
Dim nocampos As Integer
Dim i As Integer
Dim j As Integer
Dim mata() As Variant
Dim matb() As Variant
Dim rmesa As New ADODB.recordset

 txtfecha = "to_date('" & Format(dtfecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 txtfiltro = "SELECT COUNT(*) from " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha
 rmesa.Open txtfiltro, ConAdo
 noreg = rmesa.Fields(0)
 rmesa.Close
If noreg <> 0 Then
   txtfiltro = "SELECT * from " & TablaEficienciaCob & " WHERE FECHA = " & txtfecha
   rmesa.Open txtfiltro, ConAdo
   nocampos = rmesa.Fields.Count
   ReDim mata(1 To noreg, 1 To nocampos + 1) As Variant
   For i = 1 To noreg
       For j = 1 To nocampos
           mata(i, j) = rmesa.Fields(j - 1)
       Next j
       mata(i, nocampos + 1) = CLng(mata(i, 1)) & " " & mata(i, 4)
       If mata(i, 17) < 80 Or mata(i, 17) > 125 Then
          MsgBox "Alerta. Una efectividad no es 80-125 Clave de operación " & mata(i, 4)
       End If
 rmesa.MoveNext
 Next i
 rmesa.Close
  matb = RutinaOrden(mata, nocampos + 1, 1)
 For i = 2 To UBound(matb, 1)
     If matb(i - 1, nocampos + 1) = matb(i, nocampos + 1) Then
        MsgBox "se repite la efectividad de la operación " & matb(i, 1) & "   " & matb(i, 4)
        Exit Sub
     End If
 Next i
 
 txtfechaf = "to_date('" & Format(dtfecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
 Call IniciarConexOracle(conAdo2, BDIKOS)
 For i = 1 To noreg
     conAdo2.Execute "DELETE FROM " & TablaEficienciaCob & " WHERE FECHA = " & txtfechaf & " and CLAVE_SWAP = '" & mata(i, 4) & "'"
     txtfechac1 = "to_date('" & Format(dtfecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfechac2 = "to_date('" & Format(mata(i, 2), "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtfechac3 = "to_date('" & Format(dtfecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
     txtcadena = "INSERT INTO " & TablaEficienciaCob & " VALUES("
     txtcadena = txtcadena & txtfechac1 & ","                'FECHA
     txtcadena = txtcadena & txtfechac2 & ","                'FECHA
     txtcadena = txtcadena & txtfechac3 & ","                'FECHA
     txtcadena = txtcadena & "'" & mata(i, 4) & "',"         'Clave de operación
     For j = 5 To 16
     If Not EsVariableVacia(mata(i, j)) Then
        txtcadena = txtcadena & mata(i, j) & ","
     Else
        txtcadena = txtcadena & "null,"
     End If
  Next j
  txtcadena = txtcadena & mata(i, 17) & ","
  txtcadena = txtcadena & mata(i, 18) & ")"               'FECHA
  conAdo2.Execute txtcadena
  AvanceProc = i / noreg
  MensajeProc = "Copiando efectividad del dia " & Format(dtfecha1, "dd/mm/yyyy") & " " & Format(AvanceProc, "##0.00 %")
  DoEvents
 Next i
 conAdo2.Close
 txtmsg = "Se copiaron " & noreg & " registros"
 exito = True
End If
End Sub

Sub GuardarResEfRetroSwap2(ByVal fecha As Date, ByRef matpos() As propPosRiesgo, _
                           ByRef matposswaps() As propPosSwaps, _
                           ByVal ind As Integer, ByVal t_efic As Integer, ByVal valefic As Double)
Dim txtfecha As String
Dim txtcadena As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim indice As Integer
Dim fvence As Date
Dim cproducto As String

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
indice = matpos(ind).IndPosicion
fvence = matposswaps(indice).FvencSwap
cproducto = matposswaps(indice).ClaveProdSwap

    txtcadena = "DELETE FROM " & TablaEficRetro & " WHERE FECHA = " & txtfecha & " AND COPERACION = '" & matpos(ind).c_operacion & "'"
    ConAdo.Execute txtcadena
    txtfecha1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(fvence, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "INSERT INTO " & TablaEficRetro & " VALUES("
    txtcadena = txtcadena & txtfecha1 & ","
    txtcadena = txtcadena & "null,"
    txtcadena = txtcadena & "null,"
    txtcadena = txtcadena & "'" & matpos(ind).c_operacion & "',"
    txtcadena = txtcadena & "'" & cproducto & "',"
    txtcadena = txtcadena & txtfecha2 & ","
    txtcadena = txtcadena & "'" & t_efic & "',"
    txtcadena = txtcadena & valefic * 100 & ")"
    ConAdo.Execute txtcadena


End Sub

Sub GuardarResEfRetroFwd(ByVal fecha As Date, ByRef matpos() As propPosRiesgo, _
                         ByRef matposfwds() As propPosFwd, ByVal ind As Integer, ByVal t_efic As Integer, ByVal valefic As Double)
Dim txtfecha As String
Dim indice As Integer
Dim txtcadena As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim cproducto As String
Dim fvence As Date

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
indice = matpos(ind).IndPosicion
cproducto = matposfwds(indice).ClaveProdFwd
fvence = matposfwds(indice).FVencFwd
    txtcadena = "DELETE FROM " & TablaEficRetro & " WHERE FECHA = " & txtfecha & " AND COPERACION = '" & matpos(ind).c_operacion & "'"
    ConAdo.Execute txtcadena
    txtfecha1 = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtfecha2 = "to_date('" & Format(fvence, "dd/mm/yyyy") & "','dd/mm/yyyy')"
    txtcadena = "INSERT INTO " & TablaEficRetro & " VALUES("
    txtcadena = txtcadena & txtfecha1 & ","
    txtcadena = txtcadena & "null,"
    txtcadena = txtcadena & "null,"
    txtcadena = txtcadena & "'" & matpos(ind).c_operacion & "',"
    txtcadena = txtcadena & "'" & cproducto & "',"
    txtcadena = txtcadena & txtfecha2 & ","                      'fecha de vencimiento
    txtcadena = txtcadena & t_efic & ","                         'tipo de efectividad
    txtcadena = txtcadena & valefic * 100 & ")"
    ConAdo.Execute txtcadena


End Sub

Function LeerCurvasRE(ByVal fecha As Date)
Dim indice As Long
Dim fecha1 As Date
Dim fecha7 As Date
Dim fecha30 As Date
Dim mcurvas0() As Variant
Dim mcurvas1() As Variant
Dim mcurvas7() As Variant
Dim mcurvas30() As Variant
Dim idescirs As Integer
Dim icetes As Integer
Dim ilibor As Integer
Dim ibondesd As Integer
Dim iipabis As Integer
Dim iccmid As Integer
Dim iccsudi As Integer
Dim iccsyen As Integer
Dim irealimp As Integer

Dim noreg As Long
Dim i As Long
Dim exito1 As Boolean
noreg = 12000
idescirs = 21
icetes = 20
ilibor = 30
ibondesd = 9
iipabis = 29
iccmid = 16
iccsudi = 45
iccsyen = 17
irealimp = 44

indice = BuscarValorArray(fecha, MatFechasVaR, 1)
If indice <> 0 Then
   fecha1 = MatFechasVaR(indice - 1, 1)
   fecha7 = MatFechasVaR(indice - 5, 1)
   fecha30 = MatFechasVaR(indice - 21, 1)
   mcurvas0 = LeerCurvaCompleta(fecha, exito1)
   mcurvas1 = LeerCurvaCompleta(fecha1, exito1)
   mcurvas7 = LeerCurvaCompleta(fecha7, exito1)
   mcurvas30 = LeerCurvaCompleta(fecha30, exito1)
   ReDim mcurvas(1 To noreg, 1 To 36) As Variant
   For i = 1 To noreg
       mcurvas(i, 1) = mcurvas0(i + 1, idescirs)
       mcurvas(i, 2) = mcurvas1(i + 1, idescirs)
       mcurvas(i, 3) = mcurvas7(i + 1, idescirs)
       mcurvas(i, 4) = mcurvas30(i + 1, idescirs)
       mcurvas(i, 5) = mcurvas0(i + 1, icetes)
       mcurvas(i, 6) = mcurvas1(i + 1, icetes)
       mcurvas(i, 7) = mcurvas7(i + 1, icetes)
       mcurvas(i, 8) = mcurvas30(i + 1, icetes)
       mcurvas(i, 9) = mcurvas0(i + 1, ilibor)
       mcurvas(i, 10) = mcurvas1(i + 1, ilibor)
       mcurvas(i, 11) = mcurvas7(i + 1, ilibor)
       mcurvas(i, 12) = mcurvas30(i + 1, ilibor)
       mcurvas(i, 13) = mcurvas0(i + 1, ibondesd)
       mcurvas(i, 14) = mcurvas1(i + 1, ibondesd)
       mcurvas(i, 15) = mcurvas7(i + 1, ibondesd)
       mcurvas(i, 16) = mcurvas30(i + 1, ibondesd)
       mcurvas(i, 17) = mcurvas0(i + 1, iipabis)
       mcurvas(i, 18) = mcurvas1(i + 1, iipabis)
       mcurvas(i, 19) = mcurvas7(i + 1, iipabis)
       mcurvas(i, 20) = mcurvas30(i + 1, iipabis)
       mcurvas(i, 21) = mcurvas0(i + 1, iccmid)
       mcurvas(i, 22) = mcurvas1(i + 1, iccmid)
       mcurvas(i, 23) = mcurvas7(i + 1, iccmid)
       mcurvas(i, 24) = mcurvas30(i + 1, iccmid)
       mcurvas(i, 25) = mcurvas0(i + 1, iccsudi)
       mcurvas(i, 26) = mcurvas1(i + 1, iccsudi)
       mcurvas(i, 27) = mcurvas7(i + 1, iccsudi)
       mcurvas(i, 28) = mcurvas30(i + 1, iccsudi)
       mcurvas(i, 29) = mcurvas0(i + 1, iccsyen)
       mcurvas(i, 30) = mcurvas1(i + 1, iccsyen)
       mcurvas(i, 31) = mcurvas7(i + 1, iccsyen)
       mcurvas(i, 32) = mcurvas30(i + 1, iccsyen)
       mcurvas(i, 33) = mcurvas0(i + 1, irealimp)
       mcurvas(i, 34) = mcurvas1(i + 1, irealimp)
       mcurvas(i, 35) = mcurvas7(i + 1, irealimp)
       mcurvas(i, 36) = mcurvas30(i + 1, irealimp)
   Next i
Else
  ReDim mcurvas(0 To 0, 0 To 0) As Variant
End If
LeerCurvasRE = mcurvas
End Function

Sub EncontrarInstTabValuacion(ByVal cemision As String, ByVal tv As String, ByVal emision As String, ByVal serie As String, t_ref As String, ByVal st As String, ByVal yield As Double)
Dim noreg As Long
Dim txtfiltro As String
Dim rmesa As New ADODB.recordset
        If tv = "90" Or tv = "91" Or tv = "92" Or tv = "93" Or tv = "94" Or tv = "95" Or tv = "D8" Or tv = "JI" Or tv = "2U" Or tv = "CD" Or tv = "JE" Or tv = "D2" Or tv = "F" Then
           If t_ref = "TIIE 28 dias" Then
                txtfiltro = "SELECT COUNT(*) FROM " & PrefijoBD & TablaValBSC & " WHERE TIPO_VALOR = '" & tv & "'"
                txtfiltro = txtfiltro & " AND EMISORA = '" & emision & "' AND SERIE = '" & serie & "'"
                rmesa.Open txtfiltro, ConAdo
                noreg = rmesa.Fields(0)
                rmesa.Close
                If noreg = 0 Then
                   Print #1, tv & Chr(9) & emision & Chr(9) & serie & Chr(9) & "BONO STC Y" & Chr(9) & t_ref & Chr(9) & st & Chr(9) & "28" & Chr(9) & Val(st) / 100 & Chr(9) & "Y " & emision & serie & Chr(9) & "1" & Chr(9) & "TIIE28 PIP" & Chr(9) & "1"
                End If
            ElseIf t_ref = "TIIE 91 dias" Then
                txtfiltro = "SELECT COUNT(*) FROM " & PrefijoBD & TablaValBSC & " WHERE TIPO_VALOR = '" & tv & "'"
                txtfiltro = txtfiltro & " AND EMISORA = '" & emision & "' AND SERIE = '" & serie & "'"
                rmesa.Open txtfiltro, ConAdo
                noreg = rmesa.Fields(0)
                rmesa.Close
                If noreg = 0 Then
                   Print #1, tv & Chr(9) & emision & Chr(9) & serie & Chr(9) & "BONO STC Y" & Chr(9) & t_ref & Chr(9) & st & Chr(9) & "28" & Chr(9) & Val(yield) / 100 & Chr(9) & "Y " & emision & serie & Chr(9) & "1" & Chr(9) & "TIIE91 PIP" & Chr(9) & "1"
                End If
            ElseIf t_ref = "Tasa Fija" Then
                If tv <> "M" And tv <> "S" And tv <> "2U" And tv <> "PI" Then
                   txtfiltro = "SELECT COUNT(*) FROM " & PrefijoBD & TablaValBonos & " WHERE TIPO_VALOR = '" & tv & "'"
                   txtfiltro = txtfiltro & " AND EMISORA = '" & emision & "' AND SERIE = '" & serie & "'"
                   rmesa.Open txtfiltro, ConAdo
                   noreg = rmesa.Fields(0)
                   rmesa.Close
                   If noreg = 0 Then
                      Print #2, tv & Chr(9) & emision & Chr(9) & serie & Chr(9) & "BONO TASA FIJA Y" & Chr(9) & "182" & Chr(9) & "Y " & emision & serie & Chr(9) & "1" & Chr(9) & "TC"
                   End If
                End If
            End If
        ElseIf tv = "1" Or tv = "1A" Or tv = "1I" Or tv = "41" Or tv = "51" Or tv = "52" Or tv = "CF" Or tv = "1B" Then
                txtfiltro = "SELECT COUNT(*) FROM " & PrefijoBD & TablaValInds & " WHERE TIPO_VALOR = '" & tv & "'"
                txtfiltro = txtfiltro & " AND EMISORA = '" & emision & "' AND SERIE = '" & serie & "'"
                rmesa.Open txtfiltro, ConAdo
                noreg = rmesa.Fields(0)
                rmesa.Close
                If noreg = 0 Then
                   Print #1, tv & Chr(9) & emision & Chr(9) & serie & Chr(9) & "TIPO CAMBIO" & Chr(9) & emision & serie & Chr(9) & "1"
                End If
         ElseIf tv = "I" Or tv = "LD" Or tv = "PI" Or tv = "IS" Or tv = "M" Or tv = "S" Or tv = "BI" Or tv = "IQ" Then
            Print #7, "No esta en la lista " & tv
         Else
            Print #7, "No esta en ninguna tabla de valuacion " & cemision
        End If
End Sub

Sub EncontrarInstTabParametros(ByVal cemision As String, ByVal tv As String, ByVal emision As String, ByVal serie As String, ByVal valor1 As String, ByVal valor2 As String)
Dim noreg As Long
Dim txtfiltro As String
Dim rmesa As New ADODB.recordset
If tv <> "M" And tv <> "S" And tv <> "2U" And tv <> "PI" And tv <> "BI" And tv <> "I" And tv <> "IQ" Then
   txtfiltro = "SELECT COUNT(*) FROM " & PrefijoBD & TablaIndVecPreciosO & " WHERE CEMISION = '" & cemision & "'"
   rmesa.Open txtfiltro, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg = 0 Then
      If tv = "90" Or tv = "91" Or tv = "92" Or tv = "93" Or tv = "94" Or tv = "95" Or tv = "CD" Or tv = "F" Or tv = "JI" Then
         Print #4, valor1 & Chr(9) & valor2 & Chr(9) & "Y " & emision & serie & Chr(9) & "TASA" & Chr(9) & "17" & Chr(9) & cemision
      ElseIf tv = "51" Or tv = "52" Or tv = "1" Or tv = "1A" Or tv = "1I" Or tv = "41" Or tv = "CF" Then
         Print #4, valor1 & Chr(9) & valor2 & Chr(9) & emision & serie & Chr(9) & "INDICE" & Chr(9) & "7" & Chr(9) & cemision
      ElseIf tv = "I" Or tv = "LD" Or tv = "M" Or tv = "S" Or tv = "PI" Or tv = "BI" Or tv = "IS" Or tv = "IQ" Then
      'NO SE TOMAN
      Else
         Print #7, "No esta en los catalogos de instrumentos " & cemision
      End If
   End If
End If
End Sub

Sub DetermSiHayFlujosEm(ByVal cemision As String, ByVal tv As String, regla_cp As String, ByVal f_em As Date, ByVal f_venc As Date)
Dim txtfiltro As String
Dim pcupon As Integer
Dim mnocional As Double
Dim noreg As Long
Dim rmesa As New ADODB.recordset
If tv <> "I" And tv <> "BI" And tv <> "51" And tv <> "52" And tv <> "1I" And tv <> "1A" And tv <> "1" And tv <> "41" And tv <> "CF" Then
   txtfiltro = "SELECT COUNT(*) FROM " & TablaFlujosMD & " WHERE EMISION = '" & cemision & "'"
   rmesa.Open txtfiltro, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg = 0 Then
      Print #7, "No hay flujos para la emision " & cemision
      mnocional = 100
      If regla_cp = "Tasa Fija" Then
         pcupon = 182
      ElseIf regla_cp = "TIIE 28 dias" Then
         pcupon = 28
      ElseIf regla_cp = "TIIE 91 dias" Then
         pcupon = 91
      End If
      If (f_venc - f_em) Mod pcupon = 0 Then
         Call ConstruirFlujosEmision(f_em, f_venc, cemision, mnocional, pcupon)
      Else
      End If
   End If
End If
End Sub

Sub DetermSiFRBD(ByVal tv As String, ByVal emision As String, ByVal serie As String)
Dim txtfiltro As String
Dim txtclave As String
Dim txtdescrip As String
Dim noreg As Long
Dim txtmsg As String
Dim rmesa As New ADODB.recordset

If tv = "41" Or tv = "1" Or tv = "51" Or tv = "52" Or tv = "1A" Then
   txtclave = emision & serie
   txtdescrip = "Indice " & tv & emision & serie
   txtfiltro = "SELECT COUNT(*) FROM " & PrefijoBD & TablaPortFR & " WHERE PORTAFOLIO = 'PRUEBA 2'"
   txtfiltro = txtfiltro & " AND CONCEPTO = '" & txtclave & "'"
   rmesa.Open txtfiltro, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg = 0 Then
      Print #6, #1/1/2003# & Chr(9) & "PRUEBA 2" & Chr(9) & "ME" & Chr(9) & "INDICE" & Chr(9) & "INDICE" & Chr(9) & txtclave & Chr(9) & "0" & Chr(9) & txtdescrip
   End If

ElseIf tv = "90" Or tv = "91" Or tv = "92" Or tv = "932" Or tv = "94" Or tv = "95" Or tv = "F" Or tv = "D8" Then
   txtclave = "Y " & emision & serie
   txtdescrip = "Yield " & emision & serie
   txtfiltro = "SELECT COUNT(*) FROM " & PrefijoBD & TablaPortFR & " WHERE PORTAFOLIO = 'PRUEBA 2'"
   txtfiltro = txtfiltro & " AND CONCEPTO = '" & txtclave & "'"
   rmesa.Open txtfiltro, ConAdo
   noreg = rmesa.Fields(0)
   rmesa.Close
   If noreg = 0 Then
      Print #6, #1/1/2003# & Chr(9) & "PRUEBA 2" & Chr(9) & "MD" & Chr(9) & "YIELD" & Chr(9) & "INDICE" & Chr(9) & txtclave & Chr(9) & "0" & Chr(9) & txtdescrip
   End If
Else
  txtmsg = "no esta en la lista" & tv & emision & serie
End If
End Sub


Sub DeterminaSiHayFR(ByVal tv As String, ByVal emision As String, ByVal serie As String, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal nop As Integer)
Dim txtfiltro As String
Dim txtclave As String
Dim txtdescrip As String
Dim noreg1 As Long
Dim noreg2 As Long
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtmsg As String
Dim rmesa As New ADODB.recordset

txtfecha1 = "TO_DATE('" & Format(fecha1, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfecha2 = "TO_DATE('" & Format(fecha2, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro = "SELECT COUNT(*) FROM " & TablaFechasVaR & " WHERE "
txtfiltro = txtfiltro & " FECHA >= " & txtfecha1
txtfiltro = txtfiltro & " AND FECHA <= " & txtfecha2
rmesa.Open txtfiltro, ConAdo
noreg1 = rmesa.Fields(0)
rmesa.Close

If tv = "41" Or tv = "1" Or tv = "51" Or tv = "52" Or tv = "1A" Or tv = "1I" Or tv = "FE" Or tv = "1B" Then
   txtclave = emision & serie
   txtdescrip = "Indice " & tv & emision & serie
   txtfiltro = "SELECT COUNT(*) FROM " & TablaFRiesgoO & " WHERE CONCEPTO = '" & txtclave & "'"
   txtfiltro = txtfiltro & " AND FECHA >= " & txtfecha1
   txtfiltro = txtfiltro & " AND FECHA <= " & txtfecha2
   rmesa.Open txtfiltro, ConAdo
   noreg2 = rmesa.Fields(0)
   rmesa.Close
   If noreg2 < noreg1 Then
      Print #nop, txtclave & Chr(9) & noreg2 & Chr(9) & noreg1
      Call ExtraeInfVPaFR(tv, emision, serie, txtclave, fecha1, fecha2, 0)
   End If
ElseIf tv = "90" Or tv = "91" Or tv = "92" Or tv = "93" Or tv = "94" Or tv = "95" Or tv = "F" Or tv = "D8" Or tv = "CF" Or tv = "CD" Or tv = "JI" Then
   txtclave = "Y " & emision & serie
   txtdescrip = "Yield " & emision & serie
   txtfiltro = "SELECT COUNT(*) FROM " & TablaFRiesgoO & " WHERE CONCEPTO = '" & txtclave & "'"
   txtfiltro = txtfiltro & " AND FECHA >= " & txtfecha1
   txtfiltro = txtfiltro & " AND FECHA <= " & txtfecha2
   rmesa.Open txtfiltro, ConAdo
   noreg2 = rmesa.Fields(0)
   rmesa.Close
   If noreg2 < noreg1 Then
      Print #nop, txtclave & Chr(9) & noreg2 & Chr(9) & noreg1
      Call ExtraeInfVPaFR(tv, emision, serie, txtclave, fecha1, fecha2, 1)
   End If
ElseIf tv = "LD" Or tv = "IM" Or tv = "IQ" Or tv = "IS" Or tv = "M" Or tv = "S" Or tv = "BI" Or tv = "I" Or tv = "2U" Or tv = "PI" Then

Else
  Print #nop, "no esta en la lista" & tv & emision & serie
End If

End Sub

Sub DeterminaSiHayFR2(ByVal tv As String, ByVal emision As String, ByVal serie As String, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal nop As Integer)
Dim txtfiltro As String
Dim txtclave As String
Dim txtdescrip As String
Dim noreg1 As Long
Dim noreg2 As Long
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtmsg As String
Dim rmesa As New ADODB.recordset

txtfecha1 = "TO_DATE('" & Format(fecha1, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfecha2 = "TO_DATE('" & Format(fecha2, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro = "SELECT COUNT(*) FROM " & TablaFechasVaR & " WHERE "
txtfiltro = txtfiltro & " FECHA >= " & txtfecha1
txtfiltro = txtfiltro & " AND FECHA <= " & txtfecha2
rmesa.Open txtfiltro, ConAdo
noreg1 = rmesa.Fields(0)
rmesa.Close

If tv = "41" Or tv = "1" Or tv = "51" Or tv = "52" Or tv = "1A" Or tv = "1I" Or tv = "FE" Or tv = "1B" Then
   txtclave = emision & serie
   txtdescrip = "Indice " & tv & emision & serie
   txtfiltro = "SELECT COUNT(*) FROM " & TablaFRiesgoO & " WHERE CONCEPTO = '" & txtclave & "'"
   txtfiltro = txtfiltro & " AND FECHA >= " & txtfecha1
   txtfiltro = txtfiltro & " AND FECHA <= " & txtfecha2
   rmesa.Open txtfiltro, ConAdo
   noreg2 = rmesa.Fields(0)
   rmesa.Close
   If noreg2 < noreg1 Then
      Print #nop, txtclave & Chr(9) & noreg2 & Chr(9) & noreg1
   End If
ElseIf tv = "90" Or tv = "91" Or tv = "92" Or tv = "93" Or tv = "94" Or tv = "95" Or tv = "F" Or tv = "D8" Or tv = "CF" Or tv = "CD" Or tv = "JI" Then
   txtclave = "Y " & emision & serie
   txtdescrip = "Yield " & emision & serie
   txtfiltro = "SELECT COUNT(*) FROM " & TablaFRiesgoO & " WHERE CONCEPTO = '" & txtclave & "'"
   txtfiltro = txtfiltro & " AND FECHA >= " & txtfecha1
   txtfiltro = txtfiltro & " AND FECHA <= " & txtfecha2
   rmesa.Open txtfiltro, ConAdo
   noreg2 = rmesa.Fields(0)
   rmesa.Close
   If noreg2 < noreg1 Then
      Print #nop, txtclave & Chr(9) & noreg2 & Chr(9) & noreg1
   End If
ElseIf tv = "LD" Or tv = "IM" Or tv = "IQ" Or tv = "IS" Or tv = "M" Or tv = "S" Or tv = "BI" Or tv = "I" Or tv = "2U" Or tv = "PI" Then

Else
  Print #nop, "no esta en la lista" & tv & emision & serie
End If

End Sub


Sub ConstruirFlujosEmision(ByVal finicio As Date, ByVal ffinal As Date, ByVal cemision As String, ByVal mnocional As Double, ByVal pcupon As Integer)
Dim noreg As Long
Dim i As Long
Dim j As Long
Dim txtcadena As String

noreg = (ffinal - finicio) / pcupon
ReDim mata(1 To noreg, 1 To 8) As Variant
For i = 1 To noreg
    mata(i, 1) = cemision
    mata(i, 2) = finicio
    If i = 1 Then
       mata(i, 3) = finicio
    Else
        mata(i, 3) = mata(i - 1, 4)
    End If
    mata(i, 4) = mata(i, 3) + pcupon
    mata(i, 5) = mnocional
    If i = noreg Then
       mata(i, 6) = mnocional
    Else
       mata(i, 6) = 0
    End If
    mata(i, 7) = 0
    mata(i, 8) = pcupon
Next i
For i = 1 To noreg
txtcadena = ""
For j = 1 To 8
txtcadena = txtcadena & mata(i, j) & Chr(9)
Next j
Print #5, txtcadena
Next i
End Sub

Sub ExtraeInfVPaFR(ByVal tv As String, ByVal emision As String, ByVal serie As String, ByVal concepto As String, ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal opcion As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfechaa As String
Dim txtfechab As String
Dim noreg As Long
Dim columna As Integer
Dim fechaa As Date
Dim fechab As Date
Dim txtborra As String
Dim i As Long
Dim fecha As Date
Dim valor As Double
Dim indice As String
Dim txtcadena As String
Dim rmesa As New ADODB.recordset

txtfecha1 = "TO_DATE('" & Format(fecha1, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfecha2 = "TO_DATE('" & Format(fecha2, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT MIN(FECHA) AS FECHA1,MAX(FECHA) AS FECHA2 FROM " & TablaVecPrecios & " WHERE FECHA >= " & txtfecha1
txtfiltro2 = txtfiltro2 & " AND FECHA <= " & txtfecha2
txtfiltro2 = txtfiltro2 & " AND TV = '" & tv & "'"
txtfiltro2 = txtfiltro2 & " AND EMISION = '" & emision & "'"
txtfiltro2 = txtfiltro2 & " AND SERIE = '" & serie & "'"
rmesa.Open txtfiltro2, ConAdo
fechaa = rmesa.Fields("FECHA1")
fechab = rmesa.Fields("FECHA2")
rmesa.Close
txtfiltro2 = "SELECT FECHA,PSUCIO,YIELD FROM " & TablaVecPrecios & " WHERE FECHA >= " & txtfecha1
txtfiltro2 = txtfiltro2 & " AND FECHA <= " & txtfecha2
txtfiltro2 = txtfiltro2 & " AND TV = '" & tv & "'"
txtfiltro2 = txtfiltro2 & " AND EMISION = '" & emision & "'"
txtfiltro2 = txtfiltro2 & " AND SERIE = '" & serie & "' ORDER BY FECHA"
txtfiltro1 = "SELECT COUNT(*) FROM  (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
  If opcion = 0 Then
      columna = 1
   Else
      columna = 2
   End If
   txtfechaa = "TO_DATE('" & Format(fechaa, "dd/mm/yyyy") & "','DD/MM/YYYY')"
   txtfechab = "TO_DATE('" & Format(fechab, "dd/mm/yyyy") & "','DD/MM/YYYY')"
   txtborra = "DELETE FROM " & TablaFRiesgoO & " WHERE CONCEPTO = '" & concepto & "'"
   txtborra = txtborra & " AND PLAZO = 0"
   txtborra = txtborra & " AND FECHA >= " & txtfechaa
   txtborra = txtborra & " AND FECHA <= " & txtfechab
   ConAdo.Execute txtborra
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       fecha = rmesa.Fields("FECHA")
       If opcion = 0 Then
          valor = Val(rmesa.Fields(columna))
       Else
          valor = Val(rmesa.Fields(columna)) / 100
       End If
       txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
       indice = CLng(fecha) & concepto & "0000000"
       txtcadena = "INSERT INTO " & TablaFRiesgoO & " VALUES("
       txtcadena = txtcadena & txtfecha & ","
       txtcadena = txtcadena & "'" & concepto & "',"
       txtcadena = txtcadena & "0,"
       txtcadena = txtcadena & valor & ","
       txtcadena = txtcadena & "'" & indice & "')"
       ConAdo.Execute txtcadena
       rmesa.MoveNext
   Next i
   rmesa.Close
End If
End Sub

Function CambFechaxSerie(ByVal txtserie As String)
If txtserie = "01-oct." Then
   CambFechaxSerie = "1-10"
ElseIf txtserie = "06-oct." Then
   CambFechaxSerie = "6-10"
ElseIf txtserie = "09-abr." Then
   CambFechaxSerie = "9-4"
ElseIf txtserie = "10-feb." Then
   CambFechaxSerie = "10-2"
ElseIf txtserie = "11-mar." Then
   CambFechaxSerie = "11-3"
ElseIf txtserie = "12-mar." Then
   CambFechaxSerie = "12-3"
ElseIf txtserie = "13-feb." Then
   CambFechaxSerie = "13-2"
ElseIf txtserie = "14-feb." Then
   CambFechaxSerie = "14-2"
ElseIf txtserie = "15-feb." Then
   CambFechaxSerie = "15-2"
ElseIf txtserie = "16-feb." Then
   CambFechaxSerie = "16-2"
ElseIf txtserie = "17-feb." Then
   CambFechaxSerie = "17-2"
ElseIf txtserie = "17-mar." Then
   CambFechaxSerie = "17-3"
ElseIf txtserie = "17-jun." Then
   CambFechaxSerie = "17-6"
ElseIf txtserie = "18-feb." Then
   CambFechaxSerie = "18-2"
ElseIf txtserie = "18-mar." Then
   CambFechaxSerie = "18-3"
Else
   CambFechaxSerie = txtserie
End If

End Function

Function depurartablafp1(ByRef mata() As Variant, ByVal fecha As Date, ByVal txttabla As String, ByVal txtnompos As String, ByRef contar As Long)
Dim txtsubportfp As String
Dim i As Integer
Dim contar1 As Integer

ReDim matb(1 To 18, 1 To 1)
txtsubportfp = ""
contar1 = 0
For i = 1 To UBound(mata, 1)
    If Left(mata(i, 4), 4) = "983 " Or Right(mata(i, 4), 28) = "(PENSIONES 983 / antes 2505)" Or Right(mata(i, 4), 15) = "(983 PENSIONES)" Then txtsubportfp = "983"
    If Left(mata(i, 4), 4) = "984 " Then txtsubportfp = "984"
    If Left(mata(i, 4), 4) = "985 " Or Right(mata(i, 4), 17) = "985 / antes 2507)" Then txtsubportfp = "985"
    If Left(mata(i, 4), 4) = "986 " Then txtsubportfp = "986"
    If Left(mata(i, 4), 4) = "987 " Then txtsubportfp = "987"
    If Val(mata(i, 7)) > 0 And Val(mata(i, 9)) <> 0 And Not EsVariableVacia(mata(i, 1)) And Not EsVariableVacia(mata(i, 2)) Then
       contar1 = contar1 + 1
       contar = contar + 1
       ReDim Preserve matb(1 To 18, 1 To contar1) As Variant
       matb(1, contar1) = fecha
       matb(2, contar1) = "N"
       matb(3, contar1) = ClavePosPension1
       matb(4, contar1) = contar                                 'clave de operacion
       If mata(i, 10) <> 0 Then
          matb(5, contar1) = "R"
       Else
          matb(5, contar1) = "D"
       End If
       matb(6, contar1) = mata(i, 1)                            'TIPO VALOR
       matb(7, contar1) = mata(i, 2)                            'EMISION
       matb(8, contar1) = CambFechaxSerie(Trim(mata(i, 3)))     'SERIE
       matb(9, contar1) = mata(i, 4)                            'CLAVE DE EMISION
       matb(10, contar1) = mata(i, 7)                           'NO DE TITULOS
       matb(11, contar1) = mata(i, 5)                           'FECHA DE COMPRA
       matb(12, contar1) = mata(i, 6)                           'FECHA DE VENCIMIENTO
       matb(13, contar1) = mata(i, 9)                           'P COMPRA
       matb(14, contar1) = mata(i, 10)                          'T PREMIO
       matb(15, contar1) = txtsubportfp                         'TXTSUBPORTFP
       matb(16, contar1) = txttabla                             'administradora
       matb(17, contar1) = ""                                   'calificacion
       matb(18, contar1) = "N"                                  'si flujos
    End If
Next i
If contar1 = 0 Then
ReDim matb(0 To 0, 0 To 0) As Variant
depurartablafp1 = matb
Else
depurartablafp1 = MTranV(matb)
End If
End Function

Function depurartablafp2(ByRef mata() As Variant, ByVal fecha As Date, ByVal txtposicion As String, ByRef contar As Long)
Dim txtadmin As String
Dim i As Integer
Dim contar1 As Integer
ReDim matb(1 To 18, 1 To 1)
txtadmin = ""
contar1 = 0
For i = 1 To UBound(mata, 1)
    If mata(i, 4) = "GBM" Then txtadmin = "GBM"
    If mata(i, 4) = "VECTOR" Then txtadmin = "VECTOR"
    If mata(i, 4) = "ACTINVER" Then txtadmin = "ACTINVER"
    If mata(i, 4) = "BANAMEX" Then txtadmin = "BANAMEX"
    If mata(i, 4) = "BANOBRAS" Then txtadmin = "BANOBRAS"
    If Val(mata(i, 7)) > 0 And Val(mata(i, 9)) <> 0 And Not EsVariableVacia(mata(i, 1)) And Not EsVariableVacia(mata(i, 2)) Then
       contar = contar + 1
       contar1 = contar1 + 1
       ReDim Preserve matb(1 To 18, 1 To contar1) As Variant
       matb(1, contar1) = fecha
       matb(2, contar1) = "N"
       matb(3, contar1) = ClavePosPension2
       matb(4, contar1) = contar
       If mata(i, 10) <> 0 Then
          matb(5, contar1) = "R"
       Else
          matb(5, contar1) = "D"
       End If
       matb(6, contar1) = mata(i, 1)                            'TIPO VALOR
       matb(7, contar1) = mata(i, 2)                            'EMISION
       matb(8, contar1) = CambFechaxSerie(Trim(mata(i, 3)))     'SERIE
       matb(9, contar1) = mata(i, 4)                            'CLAVE DE EMISION
       matb(10, contar1) = mata(i, 8)                           'NO DE TITULOS
       matb(11, contar1) = mata(i, 5)                           'FECHA DE COMPRA
       matb(12, contar1) = mata(i, 6)                           'FECHA DE VENCIMIENTO
       matb(13, contar1) = mata(i, 9)                           'P COMPRA
       matb(14, contar1) = mata(i, 10)                          'T PREMIO
       matb(15, contar1) = txtposicion                          'posicion
       matb(16, contar1) = txtadmin                             'txtadmin
       matb(17, contar1) = ""                                   'calificacion
       matb(18, contar1) = "N"                                  'si flujos
    End If
Next i
If contar1 = 0 Then
   ReDim matb(0 To 0, 0 To 0) As Variant
   depurartablafp2 = matb
Else
   depurartablafp2 = MTranV(matb)
End If
End Function

Function LeerHojaCalc1(ByVal txtnomarch As String, ByVal txthojacalc As String)

Dim SheetName As String
Dim RS As ADODB.recordset
Dim LI As ListItem
Dim i As Integer
Dim txtcadena As String
Dim l() As Variant
Dim strconexion As String
Dim noreg As Integer
Dim txtval As String
Dim val1 As Variant
Dim conadoex As New ADODB.Connection

strconexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & txtnomarch & "';"
strconexion = strconexion & "Extended Properties=" & Chr(34) & "Excel 12.0 Xml;HDR=YES;IMEX=0" & Chr(34)
conadoex.ConnectionString = strconexion
conadoex.Open
RegExcel.Open "SELECT count(*) FROM [" & txthojacalc & "$]", conadoex
noreg = RegExcel.Fields(0)
RegExcel.Close
RegExcel.Open "SELECT * FROM [" & txthojacalc & "$]", conadoex
    ReDim l(1 To noreg, 1 To 10) As Variant

    For i = 1 To noreg
        l(i, 1) = ReemplazaVacioValor(RegExcel.Fields(0).value, "")                  'TIPO VALOR
        l(i, 2) = ReemplazaVacioValor(RegExcel.Fields(1).value, "")                  'EMISION
        l(i, 3) = ReemplazaVacioValor(RegExcel.Fields(2).value, "")                  'serie
        l(i, 4) = ReemplazaVacioValor(RegExcel.Fields(7).value, "")                  'SUBPORTAFOLIO
        val1 = ReemplazaVacioValor(RegExcel.Fields(8).value, 0)
        If IsDate(val1) Then
        l(i, 5) = CDate(val1)                                                        'FECHA DE COMPRA
        Else
        l(i, 5) = 0
        End If
        val1 = ReemplazaVacioValor(RegExcel.Fields(12).value, 0)
        If IsDate(val1) Then
        l(i, 6) = CDate(val1)           'FECHA DE VENCIMIENTO
        Else
        l(i, 6) = 0
        End If
        val1 = ReemplazaVacioValor(RegExcel.Fields(13).value, 0)
        If IsNumeric(val1) Then
        l(i, 7) = CDbl(val1)            'TITULOS
        Else
        l(i, 7) = 0
        End If
        val1 = ReemplazaVacioValor(RegExcel.Fields(16).value, 0)
        If IsNumeric(val1) Then
           l(i, 8) = CDbl(val1)                                                      'TOTAL DE TITULOS
        Else
           l(i, 8) = 0
        End If
        val1 = ReemplazaVacioValor(RegExcel.Fields(17).value, 0)
        If IsNumeric(val1) Then
        l(i, 9) = CDbl(val1)            'PRECIO DE COMPRA
        Else
        l(i, 9) = 0
        End If
        txtval = ReemplazaCadenaTexto(ReemplazaVacioValor(RegExcel.Fields(30).value, 0), "%", "")
         If IsNumeric(txtval) Then
            l(i, 10) = CDbl(txtval) / 100 'TASA DE REPORTO
         Else
            l(i, 10) = 0
         End If
        RegExcel.MoveNext
    Next i
 RegExcel.Close
  conadoex.Close
LeerHojaCalc1 = l
End Function

Function LeerHojaCalc2(ByVal txtnomarch As String, ByVal txthojacalc As String)

Dim SheetName As String
Dim RS As ADODB.recordset
Dim LI As ListItem
Dim i As Integer
Dim txtcadena As String
Dim l() As Variant
Dim strconexion As String
Dim noreg As Integer
Dim txtval As String
Dim val1 As Variant
Dim conadoex As New ADODB.Connection

strconexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & txtnomarch & "';"
strconexion = strconexion & "Extended Properties=" & Chr(34) & "Excel 12.0 Xml;HDR=YES;IMEX=0" & Chr(34)
conadoex.ConnectionString = strconexion
conadoex.Open
RegExcel.Open "SELECT count(*) FROM [" & txthojacalc & "$]", conadoex
noreg = RegExcel.Fields(0)
RegExcel.Close
RegExcel.Open "SELECT * FROM [" & txthojacalc & "$]", conadoex
    ReDim l(1 To noreg, 1 To 10) As Variant
    For i = 1 To noreg
        l(i, 1) = ReemplazaVacioValor(RegExcel.Fields(0).value, "")                      'TIPO VALOR
        l(i, 2) = ReemplazaVacioValor(RegExcel.Fields(1).value, "")                      'EMISION
        l(i, 3) = ReemplazaVacioValor(RegExcel.Fields(2).value, "")                      'serie
        l(i, 4) = ReemplazaVacioValor(RegExcel.Fields(7).value, "")                      'SUBPORTAFOLIO 2
        val1 = ReemplazaVacioValor(RegExcel.Fields(9).value, 0)
        If IsDate(val1) Then
           l(i, 5) = CDate(val1)                'FECHA DE COMPRA
        Else
           l(i, 5) = 0
        End If
        val1 = ReemplazaVacioValor(RegExcel.Fields(13).value, 0)
        If IsDate(val1) Then
           l(i, 6) = CDate(val1)               'FECHA DE VENCIMIENTO
        Else
           l(i, 6) = 0
        End If
        val1 = ReemplazaVacioValor(RegExcel.Fields(14).value, 0)
        If IsNumeric(val1) Then
           l(i, 7) = CDbl(val1)                'TITULOS
        Else
           l(i, 7) = 0
        End If
        val1 = ReemplazaVacioValor(RegExcel.Fields(17).value, 0)
        If IsNumeric(val1) Then
           l(i, 8) = CDbl(val1)                'TOTAL DE TITULOS
        Else
           l(i, 8) = 0
        End If
        val1 = ReemplazaVacioValor(RegExcel.Fields(18).value, 0)
        If IsNumeric(val1) Then
           l(i, 9) = CDbl(val1)                'PRECIO DE COMPRA
        Else
           l(i, 9) = 0
        End If
        txtval = ReemplazaCadenaTexto(ReemplazaVacioValor(RegExcel.Fields(29).value, 0), "%", "")
        If IsNumeric(txtval) Then
           l(i, 10) = CDbl(txtval) / 100 'TASA DE REPORTO
        Else
           l(i, 10) = 0
        End If
        RegExcel.MoveNext
    Next i
RegExcel.Close
conadoex.Close
LeerHojaCalc2 = l
    
End Function

Function LeerHojaCalc3(ByVal txtnomarch As String, ByVal txthojacalc As String)

Dim SheetName As String
Dim RS As ADODB.recordset
Dim LI As ListItem
Dim i As Integer
Dim txtcadena As String
Dim l() As Variant
Dim strconexion As String
Dim noreg As Integer
Dim txtval As String
Dim val1 As Variant
Dim conadoex As New ADODB.Connection

strconexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & txtnomarch & "';"
strconexion = strconexion & "Extended Properties=" & Chr(34) & "Excel 12.0 Xml;HDR=YES;IMEX=0" & Chr(34)
conadoex.ConnectionString = strconexion
conadoex.Open
RegExcel.Open "SELECT count(*) FROM [" & txthojacalc & "$]", conadoex
noreg = RegExcel.Fields(0)
RegExcel.Close
RegExcel.Open "SELECT * FROM [" & txthojacalc & "$]", conadoex
    ReDim l(1 To noreg, 1 To 10) As Variant
    For i = 1 To noreg
        l(i, 1) = ReemplazaVacioValor(RegExcel.Fields(0).value, "")                      'TIPO VALOR
        l(i, 2) = ReemplazaVacioValor(RegExcel.Fields(1).value, "")                      'EMISION
        l(i, 3) = ReemplazaVacioValor(RegExcel.Fields(2).value, "")                      'serie
        l(i, 4) = ReemplazaVacioValor(RegExcel.Fields(7).value, "")                      'SUBPORTAFOLIO 2
        l(i, 5) = ConvCadFecha(ReemplazaVacioValor(RegExcel.Fields(12).value, 0))        'fecha de compra
        l(i, 6) = ConvCadFecha(ReemplazaVacioValor(RegExcel.Fields(16).value, 0))        'FECHA DE VENCIMIENTO
        val1 = ReemplazaVacioValor(RegExcel.Fields(17).value, 0)
        If IsNumeric(val1) Then
           l(i, 7) = CDbl(val1)                'TITULOS
        Else
           l(i, 7) = 0
        End If
        val1 = ReemplazaVacioValor(RegExcel.Fields(20).value, 0)
        If IsNumeric(val1) Then
           l(i, 8) = CDbl(val1)                                                          'TOTAL DE TITULOS
        Else
           l(i, 8) = 0
        End If
        val1 = ReemplazaVacioValor(RegExcel.Fields(21).value, 0)
        If IsNumeric(val1) Then
           l(i, 9) = CDbl(val1)                                                         'PRECIO DE COMPRA
        Else
           l(i, 9) = 0
        End If
        txtval = ReemplazaCadenaTexto(ReemplazaVacioValor(RegExcel.Fields(34).value, 0), "%", "")
        If IsNumeric(txtval) Then
           l(i, 10) = CDbl(txtval) / 100 'TASA DE REPORTO
        Else
           l(i, 10) = 0
        End If
        RegExcel.MoveNext
    Next i
RegExcel.Close
conadoex.Close
LeerHojaCalc3 = l
    
End Function

Function LeerHojaCalc4(ByVal txtnomarch As String, ByVal txthojacalc As String)
Dim SheetName As String
Dim RS As ADODB.recordset
Dim LI As ListItem
Dim i As Integer
Dim txtcadena As String
Dim l() As Variant
Dim strconexion As String
Dim noreg As Integer
Dim txtval As String
Dim conadoex As New ADODB.Connection

strconexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & txtnomarch & "';"
strconexion = strconexion & "Extended Properties=" & Chr(34) & "Excel 12.0 Xml;HDR=YES;IMEX=0" & Chr(34)
conadoex.ConnectionString = strconexion
conadoex.Open
RegExcel.Open "SELECT count(*) FROM [" & txthojacalc & "$]", conadoex
noreg = RegExcel.Fields(0)
RegExcel.Close
RegExcel.Open "SELECT * FROM [" & txthojacalc & "$]", conadoex
    ReDim l(1 To noreg, 1 To 10) As Variant
    For i = 1 To noreg
        l(i, 1) = ReemplazaVacioValor(RegExcel.Fields(0).value, "")                      'TIPO VALOR
        l(i, 2) = ReemplazaVacioValor(RegExcel.Fields(1).value, "")                      'EMISION
        l(i, 3) = ReemplazaVacioValor(RegExcel.Fields(2).value, "")                      'serie
        l(i, 4) = ReemplazaVacioValor(RegExcel.Fields(7).value, "")                      'SUBPORTAFOLIO 2
        l(i, 5) = ConvCadFecha(ReemplazaVacioValor(RegExcel.Fields(12).value, 0))               'FECHA DE COMPRA
        l(i, 6) = ConvCadFecha(ReemplazaVacioValor(RegExcel.Fields(16).value, 0))               'FECHA DE VENCIMIENTO
        l(i, 7) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(17).value, 0))                'TITULOS
        l(i, 8) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(20).value, 0))                'TOTAL DE TITULOS
        l(i, 9) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(21).value, 0))                'PRECIO DE COMPRA
        txtval = ReemplazaCadenaTexto(ReemplazaVacioValor(RegExcel.Fields(32).value, 0), "%", "")
        If IsNumeric(txtval) Then
              l(i, 10) = CDbl(txtval) / 100 'TASA DE REPORTO
        Else
              l(i, 10) = 0
        End If

        RegExcel.MoveNext
    Next i
RegExcel.Close
conadoex.Close
LeerHojaCalc4 = l
    
End Function

Function ConvCadFecha(ByVal val1 As Variant)
If IsDate(val1) Then
   ConvCadFecha = CDate(val1)
Else
   ConvCadFecha = 0
End If
End Function

Function ConvCadDbl(ByVal val1 As Variant)
If IsNumeric(val1) Then
   ConvCadDbl = CDbl(val1)
Else
   ConvCadDbl = 0
End If
End Function


Function LeerHojaCalc5(ByVal txtnomarch As String, ByVal txthojacalc As String)

Dim SheetName As String
Dim RS As ADODB.recordset
Dim LI As ListItem
Dim i As Integer
Dim txtcadena As String
Dim l() As Variant
Dim strconexion As String
Dim noreg As Integer
Dim txtval As String
Dim val1 As Variant
Dim conadoex As New ADODB.Connection

strconexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & txtnomarch & "';"
strconexion = strconexion & "Extended Properties=" & Chr(34) & "Excel 12.0 Xml;HDR=YES;IMEX=0" & Chr(34)
conadoex.ConnectionString = strconexion
conadoex.Open
RegExcel.Open "SELECT count(*) FROM [" & txthojacalc & "$]", conadoex
noreg = RegExcel.Fields(0)
RegExcel.Close
RegExcel.Open "SELECT * FROM [" & txthojacalc & "$]", conadoex
    ReDim l(1 To noreg, 1 To 12) As Variant
    For i = 1 To noreg
        l(i, 1) = ReemplazaVacioValor(RegExcel.Fields(0).value, "")                      'TIPO VALOR
        l(i, 2) = ReemplazaVacioValor(RegExcel.Fields(1).value, "")                      'EMISION
        l(i, 3) = ReemplazaVacioValor(RegExcel.Fields(2).value, "")                      'serie
        l(i, 4) = ReemplazaVacioValor(RegExcel.Fields(7).value, "")                      'TIPO DE OPERACION
        l(i, 5) = ReemplazaVacioValor(RegExcel.Fields(5).value, "")                      'SUBPORTAFOLIO 1
        l(i, 6) = ReemplazaVacioValor(RegExcel.Fields(6).value, "")                      'SUBPORTAFOLIO 2
        l(i, 7) = CDate(ReemplazaVacioValor(RegExcel.Fields(10).value, 0))               'FECHA DE COMPRA
        l(i, 8) = CDate(ReemplazaVacioValor(RegExcel.Fields(15).value, 0))               'FECHA DE VENCIMIENTO
        l(i, 9) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(11).value, 0))          'TITULOS
        l(i, 10) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(16).value, 0))         'TOTAL DE TITULOS
        l(i, 11) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(19).value, 0))         'PRECIO DE COMPRA
        l(i, 12) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(18).value, 0))         'tasa de reporto
        RegExcel.MoveNext
    Next i
    RegExcel.Close
    conadoex.Close
LeerHojaCalc5 = l
End Function

Function LeerHojaCalc6(ByVal txtnomarch As String, ByVal txthojacalc As String)

Dim SheetName As String
Dim RS As ADODB.recordset
Dim LI As ListItem
Dim i As Integer
Dim txtcadena As String
Dim l() As Variant
Dim strconexion As String
Dim noreg As Integer
Dim txtval As String
Dim val1 As Variant
Dim conadoex As New ADODB.Connection

strconexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & txtnomarch & "';"
strconexion = strconexion & "Extended Properties=" & Chr(34) & "Excel 12.0 Xml;HDR=YES;IMEX=0" & Chr(34)
conadoex.ConnectionString = strconexion
conadoex.Open
RegExcel.Open "SELECT count(*) FROM [" & txthojacalc & "$]", conadoex
noreg = RegExcel.Fields(0)
RegExcel.Close
RegExcel.Open "SELECT * FROM [" & txthojacalc & "$]", conadoex
    ReDim l(1 To noreg, 1 To 12) As Variant
    For i = 1 To noreg
        l(i, 1) = ReemplazaVacioValor(RegExcel.Fields(0).value, "")                      'TIPO VALOR
        l(i, 2) = ReemplazaVacioValor(RegExcel.Fields(1).value, "")                      'EMISION
        l(i, 3) = ReemplazaVacioValor(RegExcel.Fields(2).value, "")                      'serie
        l(i, 4) = ReemplazaVacioValor(RegExcel.Fields(7).value, "")                      'TIPO DE OPERACION
        l(i, 5) = ReemplazaVacioValor(RegExcel.Fields(5).value, "")                      'SUBPORTAFOLIO 1
        l(i, 6) = ReemplazaVacioValor(RegExcel.Fields(6).value, "")                      'SUBPORTAFOLIO 2
        l(i, 7) = CDate(ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(10).value, 0)))   'FECHA DE COMPRA
        l(i, 8) = CDate(ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(15).value, 0)))   'FECHA DE VENCIMIENTO
        l(i, 9) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(11).value, 0))          'TITULOS
        l(i, 10) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(16).value, 0))         'TOTAL DE TITULOS
        l(i, 11) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(19).value, 0))         'PRECIO DE COMPRA
        l(i, 12) = ConvCadDbl(ReemplazaVacioValor(RegExcel.Fields(18).value, 0))         'tasa de reporto
        RegExcel.MoveNext
    Next i
    RegExcel.Close
    conadoex.Close
LeerHojaCalc6 = l
End Function


Function depurartablafp3(ByRef mata() As Variant, ByVal fecha As Date, ByVal cposicion As Integer, ByRef contar As Long)
Dim txtsubportfp As String
Dim i As Integer

ReDim matb(1 To 18, 1 To 1)
txtsubportfp = ""
contar = 0
For i = 1 To UBound(mata, 1)
    If Not EsVariableVacia(mata(i, 1)) And Not EsVariableVacia(mata(i, 2)) And Not EsVariableVacia(mata(i, 3)) Then
       contar = contar + 1
       ReDim Preserve matb(1 To 18, 1 To contar) As Variant
       matb(1, contar) = fecha                                  'fecha de registro
       matb(2, contar) = "N"                                    'intencion
       matb(3, contar) = cposicion                              'clave de la posicion
       matb(4, contar) = contar                                 'clave de operacion
       If mata(i, 4) = "DIRECTO" Then                           'tipo de operacion
          matb(5, contar) = "D"
       Else
          matb(5, contar) = "R"
       End If
       matb(6, contar) = mata(i, 1)                            'TIPO VALOR
       matb(7, contar) = mata(i, 2)                            'EMISION
       matb(8, contar) = CambFechaxSerie(Trim(mata(i, 3)))     'SERIE
       matb(9, contar) = GeneraClaveEmision(mata(i, 1), mata(i, 2), mata(i, 3))
       matb(10, contar) = mata(i, 9)                           'NO DE TITULOS
       matb(11, contar) = mata(i, 7)                           'FECHA DE COMPRA
       matb(12, contar) = mata(i, 8)                           'FECHA DE VENCIMIENTO
       matb(13, contar) = mata(i, 11)                          'P COMPRA
       matb(14, contar) = mata(i, 12)                          'T PREMIO
       matb(15, contar) = mata(i, 5)                           'subportafolio
       matb(16, contar) = mata(i, 6)                           'administradora
       matb(17, contar) = ""                                   'calificacion
       matb(18, contar) = "N"                                  'si flujos
    End If
Next i
If contar = 0 Then
ReDim matb(0 To 0, 0 To 0) As Variant
depurartablafp3 = matb
Else
depurartablafp3 = MTranV(matb)
End If
End Function


Function ObtResCVaREstruc1(ByVal fecha As Date)
Dim noreg As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtport As String
Dim i As Integer
Dim j As Integer
Dim noreg0 As Integer
Dim noesc As Integer
noesc = 500
Dim rmesa As New ADODB.recordset

txtport = "TOTAL"
noreg = UBound(MatPortEstruct, 1)
ReDim matres(1 To noreg, 1 To 7) As Variant
ReDim matport(1 To 3) As String
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
For i = 1 To noreg
    matport(1) = MatPortEstruct(i)
    matport(2) = MatPortEstruct(i) & " Deriv"
    matport(3) = MatPortEstruct(i) & " Oper"
    matres(i, 1) = CLng(fecha) & "_" & MatPortEstruct(i)
    For j = 1 To 3
        txtfiltro2 = "SELECT * FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha
        txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
        txtfiltro2 = txtfiltro2 & " AND PORT_FR = 'Normal'"
        txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & matport(j) & "'"
        txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = 1"
        txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
        rmesa.Open txtfiltro1, ConAdo
        noreg0 = rmesa.Fields(0)
        rmesa.Close
        If noreg0 <> 0 Then
           rmesa.Open txtfiltro2, ConAdo
           matres(i, 2 + 2 * (j - 1)) = rmesa.Fields("MTM_SUCIO")
           rmesa.Close
        End If
        txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
        txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha
        txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha
        txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
        txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = 'Normal'"
        txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & matport(j) & "'"
        txtfiltro2 = txtfiltro2 & " AND TVAR = 'CVARH'"
        txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc
        txtfiltro2 = txtfiltro2 & " AND HTIEMPO = 1"
        txtfiltro2 = txtfiltro2 & " AND NCONF = 0.03"
        txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
        rmesa.Open txtfiltro1, ConAdo
        noreg0 = rmesa.Fields(0)
        rmesa.Close
        If noreg0 Then
           rmesa.Open txtfiltro2, ConAdo
           matres(i, 3 + 2 * (j - 1)) = rmesa.Fields("VALOR")
           rmesa.Close
        End If
    Next j
Next i
ObtResCVaREstruc1 = matres
End Function

Function ObtResCVaREstruc2(ByVal fecha As Date)
Dim noreg As Integer
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtport As String
Dim i As Integer
Dim j As Integer
Dim noreg0 As Integer
Dim noesc As Integer
Dim rmesa As New ADODB.recordset
noesc = 500

txtport = "TOTAL"
ReDim matport(1 To 4) As String
matport(1) = "DERIVADOS ESTRUCTURALES"
matport(2) = "DERIVADOS ESTRUCTURALES Y RELACIONADOS"
matport(3) = "OPER REL A ESTRUCTURALES"
matport(4) = "CONSOLIDADO+OPER REL ESTRUC"
txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
ReDim matres(1 To 4, 1 To 3) As Variant
For j = 1 To 4
    matres(j, 1) = CLng(fecha) & "_" & matport(j)
    txtfiltro2 = "SELECT * FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro2 = txtfiltro2 & " AND PORT_FR = 'Normal'"
    txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & matport(j) & "'"
    txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = 1"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg0 = rmesa.Fields(0)
    rmesa.Close
    If noreg0 <> 0 Then
       rmesa.Open txtfiltro2, ConAdo
       matres(j, 2) = rmesa.Fields("MTM_SUCIO")
       rmesa.Close
    End If
    txtfiltro2 = "SELECT * FROM " & TablaResVaR & " WHERE F_POSICION = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND F_FACTORES = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND F_VALUACION = " & txtfecha
    txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
    txtfiltro2 = txtfiltro2 & " AND ESC_FACTORES = 'Normal'"
    txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & matport(j) & "'"
    txtfiltro2 = txtfiltro2 & " AND TVAR = 'CVARH'"
    txtfiltro2 = txtfiltro2 & " AND NOESC = " & noesc
    txtfiltro2 = txtfiltro2 & " AND HTIEMPO = 1"
    txtfiltro2 = txtfiltro2 & " AND NCONF = 0.03"
    txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
    rmesa.Open txtfiltro1, ConAdo
    noreg0 = rmesa.Fields(0)
    rmesa.Close
    If noreg0 Then
       rmesa.Open txtfiltro2, ConAdo
       matres(j, 3) = rmesa.Fields("VALOR")
       rmesa.Close
    End If
Next j
ObtResCVaREstruc2 = matres
End Function

Sub GenValuacionFP1(ByVal fecha As Date)
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim txtcadena As String
Dim matf() As Date
Dim nofecha As Integer
nofecha = 1
ReDim matf(1 To nofecha) As Date


noreg1 = 5
noreg2 = 7
ReDim matport1(1 To noreg1) As String
ReDim matport2(1 To noreg2) As String

matport1(1) = "983"
matport1(2) = "985"
matport1(3) = "984"
matport1(4) = "987"
matport1(5) = "986"

matport2(1) = "BANOBRAS"
matport2(2) = "EVERCORE"
matport2(3) = "BANORTE"
matport2(4) = "SANTANDER"
matport2(5) = "VECTOR"
matport2(6) = "GBM"
matport2(7) = "BANAMEX"

Open "d:\val_fp " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #1
For i = 1 To nofecha
    ReDim mata(1 To noreg1, 1 To noreg2) As Double
    For j = 1 To noreg1
        For k = 1 To noreg2
            mata(j, k) = ValPosFP(fecha, ClavePosPension1, matport1(j), matport2(k))
        Next k
    Next j
    Print #1, matf(i)
    txtcadena = "" & Chr(9)
    For j = 1 To noreg2
       txtcadena = txtcadena & matport2(j) & Chr(9)
    Next j
    Print #1, txtcadena
    For j = 1 To noreg1
        txtcadena = matport1(j) & Chr(9)
        For k = 1 To noreg2
           txtcadena = txtcadena & mata(j, k) & Chr(9)
        Next k
        Print #1, txtcadena
    Next j
Next i

Close #1

End Sub

Sub GenValuacionFP2(ByVal fecha As Date)
Dim noreg1 As Integer
Dim noreg2 As Integer
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim txtcadena As String
Dim matf() As Date
Dim nofecha As Integer
nofecha = 1
ReDim matf(1 To nofecha) As Date



Screen.MousePointer = 11

noreg1 = 5
noreg2 = 6
ReDim matport1(1 To noreg1) As String
ReDim matport2(1 To noreg2) As String


matport1(1) = "988"
matport1(2) = "989"
matport1(3) = "990"
matport1(4) = "1111"

matport2(1) = "ACTINVER"
matport2(2) = "BANAMEX"
matport2(3) = "BANOBRAS"
matport2(4) = "GBM"
matport2(5) = "VECTOR"

Open "d:\val_fp 2 " & Format(fecha, "yyyy-mm-dd") & ".txt" For Output As #1
For i = 1 To nofecha
    ReDim mata(1 To noreg1, 1 To noreg2) As Double
    For j = 1 To noreg1
        For k = 1 To noreg2
            mata(j, k) = ValPosFP(fecha, ClavePosPension2, matport1(j), matport2(k))
        Next k
    Next j
    Print #1, matf(i)
    txtcadena = "" & Chr(9)
    For j = 1 To noreg2
       txtcadena = txtcadena & matport2(j) & Chr(9)
    Next j
    Print #1, txtcadena
    For j = 1 To noreg1
        txtcadena = matport1(j) & Chr(9)
        For k = 1 To noreg2
           txtcadena = txtcadena & mata(j, k) & Chr(9)
        Next k
        Print #1, txtcadena
    Next j
Next i

Close #1

End Sub



Function ValPosFP(ByVal fecha As Date, ByVal clavepos As Integer, ByVal txtport1 As String, ByVal txtport2 As String)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim valor1 As Double
Dim valor2 As Double
Dim valor3 As Double
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format$(fecha, "DD/MM/YYYY") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT SUM(NO_TITULOS_*VAL_PIP_S*T_OPERACION) AS VALOR FROM " & TablaValPos & " WHERE FECHAP = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 "
txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = 1 "
txtfiltro2 = txtfiltro2 & " AND (CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND CPOSICION = " & clavepos & ""
txtfiltro2 = txtfiltro2 & "  AND SUBPORT_1 = '" & txtport1 & "'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT2 = '" & txtport2 & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   valor1 = ConvValor(rmesa.Fields("VALOR"))
   rmesa.Close
Else
   valor1 = 0
End If

txtfiltro2 = "SELECT SUM(MTM_S) FROM " & TablaValPos & " WHERE FECHAP = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS =1 "
txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = 1 "
txtfiltro2 = txtfiltro2 & " AND (CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION,COPERACION FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND CPOSICION = " & clavepos & ""
txtfiltro2 = txtfiltro2 & "  AND SUBPORT_1 = '" & txtport1 & "'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT2 = '" & txtport2 & "'"
txtfiltro2 = txtfiltro2 & " AND VAL_PIP_S = 0)"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   valor2 = ConvValor(rmesa.Fields("SUM(MTM_S)"))
   rmesa.Close
Else
   valor2 = 0
End If
txtfiltro2 = "SELECT SUM(MTM_S) FROM " & TablaValPos & " WHERE FECHAP = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS =1 "
txtfiltro2 = txtfiltro2 & " AND ID_VALUACION = 1 "
txtfiltro2 = txtfiltro2 & " AND (CPOSICION,COPERACION) IN "
txtfiltro2 = txtfiltro2 & "(SELECT CPOSICION,COPERACION FROM " & TablaPosDiv & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND TIPOPOS = 1 AND CPOSICION = " & clavepos & ""
txtfiltro2 = txtfiltro2 & " AND SUBPORT1 = '" & txtport1 & "'"
txtfiltro2 = txtfiltro2 & " AND SUBPORT2 = '" & txtport2 & "')"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   valor3 = ConvValor(rmesa.Fields("SUM(MTM_S)"))
   rmesa.Close
Else
   valor3 = 0
End If

ValPosFP = valor1 + valor2 + valor3
End Function

Function LeerEscEstresS(ByVal fecha As Date)
Dim txtfecha As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim i As Long
Dim noreg As Long
Dim rmesa As New ADODB.recordset

txtfecha = "TO_DATE('" & Format(fecha, "dd/mm/yyyy") & "','DD/MM/YYYY')"
txtfiltro2 = "SELECT * FROM " & TablaResEscEstres & " WHERE FECHA= " & txtfecha
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg, 1 To 6) As Variant
   For i = 1 To noreg
       mata(i, 1) = rmesa.Fields("PORTAFOLIO")
       mata(i, 2) = rmesa.Fields("ESC_ESTRES")
       mata(i, 3) = rmesa.Fields("CPOSICION")
       mata(i, 4) = rmesa.Fields("FECHAREG")
       mata(i, 5) = rmesa.Fields("COPERACION")
       mata(i, 6) = rmesa.Fields("VALOR")
       rmesa.MoveNext
       DoEvents
   Next i
   rmesa.Close
Else
ReDim mata(0 To 0, 0 To 0) As Variant
End If
LeerEscEstresS = mata
End Function

Sub ActValIKOSValSVM(ByVal fecha As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtcadena As String
Dim noreg As Integer
Dim i As Integer
Dim vmtm As Double
Dim vactiva As Double
Dim vpasiva As Double
Dim cposicion As Integer
Dim coperacion As String
Dim rmesa As New ADODB.recordset

txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT * FROM " & TablaVDerIKOS & " WHERE FECHA = " & txtfecha
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   For i = 1 To noreg
       cposicion = 4
       coperacion = rmesa.Fields("CLAVE")
       vactiva = rmesa.Fields("VAL_ACTIVAIK")
       vpasiva = rmesa.Fields("VAL_PASIVAIK")
       vmtm = rmesa.Fields("MTMIK")
       txtcadena = "UPDATE " & TablaValPos & " SET MTM_IKOS = " & vmtm & ","
       txtcadena = txtcadena & "VAL_ACT_IKOS = " & vactiva & ","
       txtcadena = txtcadena & "VAL_PAS_IKOS = " & vpasiva & " WHERE FECHAP = " & txtfecha
       txtcadena = txtcadena & " AND CPOSICION = " & cposicion
       txtcadena = txtcadena & " AND COPERACION = " & coperacion
       ConAdo.Execute txtcadena
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Actualizando valuaciones de derivados " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   rmesa.Close
End If
End Sub

Sub LeerPlusMinuss(ByRef txtnomarch() As String)
Dim sihayarch As Boolean
Dim indice As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim txtcadena As String
Dim fecha As Date
Dim mata() As Variant
For i = 1 To UBound(txtnomarch, 1)
    sihayarch = VerifAccesoArch(txtnomarch(i))
    If sihayarch Then
       mata = LeerArchTexto(txtnomarch(i), ",", "")
       fecha = BuscarFechaTexto(txtnomarch(i))
       For j = 1 To UBound(mata, 1)
           txtcadena = fecha & Chr(9)
           For k = 1 To UBound(mata, 2)
               txtcadena = txtcadena & mata(j, k) & Chr(9)
           Next k
           Print #2, txtcadena
       Next j
    End If
Next i


End Sub

Function BuscarFechaTexto(txtcadena)
Dim txtcad1 As String
Dim i As Long
txtcad1 = ReemplazaCadenaTexto(txtcadena, ".", "")
txtcad1 = ReemplazaCadenaTexto(txtcad1, " ", "")
txtcad1 = ReemplazaCadenaTexto(txtcad1, "-", "")
For i = 1 To Len(txtcad1)
    If EsCadNumero(Mid(txtcad1, i, 8)) Then
       txtcad1 = Mid(txtcad1, i, 8)
       Exit For
    End If
Next i
If EsCadNumero(txtcad1) Then
   BuscarFechaTexto = CDate(Mid(txtcad1, 7, 2) & "/" & Mid(txtcad1, 5, 2) & "/" & Mid(txtcad1, 1, 4))
Else
   MsgBox "No encontre la cadena de fecha"
   BuscarFechaTexto = Date
End If
End Function
Function EsCadNumero(txtcadena)
Dim exito As Boolean
Dim i As Long
exito = True
For i = 1 To Len(txtcadena)
    If Mid(txtcadena, i, 1) = "0" Then
       exito = exito And True
    ElseIf Mid(txtcadena, i, 1) = "0" Then
       exito = exito And True
    ElseIf Mid(txtcadena, i, 1) = "1" Then
       exito = exito And True
    ElseIf Mid(txtcadena, i, 1) = "2" Then
       exito = exito And True
    ElseIf Mid(txtcadena, i, 1) = "3" Then
       exito = exito And True
    ElseIf Mid(txtcadena, i, 1) = "4" Then
       exito = exito And True
    ElseIf Mid(txtcadena, i, 1) = "5" Then
       exito = exito And True
    ElseIf Mid(txtcadena, i, 1) = "6" Then
       exito = exito And True
    ElseIf Mid(txtcadena, i, 1) = "7" Then
       exito = exito And True
    ElseIf Mid(txtcadena, i, 1) = "8" Then
       exito = exito And True
    ElseIf Mid(txtcadena, i, 1) = "9" Then
       exito = exito And True
    Else
       exito = exito And False
    End If
Next i
EsCadNumero = exito
End Function


Sub ExportarValDerivCF(ByVal fecha1 As Date, ByVal fecha2 As Date)
Dim fecha As Date
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfecha3 As String
Dim txtborra As String
Dim txtinserta As String
Dim i As Long
Dim finicio As Date
Dim fvence As Date
Dim intencion As String
Dim cprod As String
Dim cposicion As Integer
Dim id_contrap  As Integer
Dim mata() As Variant
Dim noreg As Long

fecha = fecha1
Do While fecha <= fecha2
   mata = LeerResValDeriv(fecha, txtportCalc1, 2, noreg)
   If UBound(mata, 1) > 0 Then
   txtfecha1 = Format(fecha, "yyyymmdd")
   txtborra = "DELETE FROM " & TablaValDeriv & " WHERE FECHA = " & txtfecha1
   ConAdo.Execute txtborra
   For i = 1 To UBound(mata, 1)
       finicio = mata(i, 12)
       fvence = mata(i, 13)
       cposicion = ClavePosDeriv
       intencion = mata(i, 14)
       cprod = mata(i, 15)
       id_contrap = mata(i, 17)
       txtfecha2 = "to_date('" & Format(fvence, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtfecha3 = "to_date('" & Format(finicio, "dd/mm/yyyy") & "','dd/mm/yyyy')"
       txtinserta = "INSERT INTO " & TablaValDeriv & " VALUES("
       txtinserta = txtinserta & txtfecha1 & ","              'fecha
       txtinserta = txtinserta & "'" & mata(i, 1) & "',"      'clave de operacion
       txtinserta = txtinserta & "'" & intencion & "',"       'intencion
       txtinserta = txtinserta & "'" & cprod & "',"           'clave de producto
       txtinserta = txtinserta & txtfecha2 & ","              'fecha de vencimiento
       txtinserta = txtinserta & mata(i, 2) & ","             'val act sivarmer
       txtinserta = txtinserta & mata(i, 3) & ","             'val pas sivarmer
       txtinserta = txtinserta & mata(i, 4) & ","             'mtm sivarmer
       txtinserta = txtinserta & mata(i, 5) & ","             'val act ikos
       txtinserta = txtinserta & mata(i, 6) & ","             'val pas ikos
       txtinserta = txtinserta & mata(i, 7) & ","             'mtm ikos
       txtinserta = txtinserta & cposicion & ","              'cposicion
       txtinserta = txtinserta & "'" & id_contrap & "',"      'id contrap
       txtinserta = txtinserta & txtfecha3 & ")"              'fecha de inicio
       ConAdo.Execute txtinserta
       MensajeProc = "Insertando registros a " & TablaValDeriv
  Next i
End If
fecha = fecha + 1
Loop
End Sub

Function ObtenerArrCurvas(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal idcurva As Integer)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Long
Dim j As Long
Dim matb() As String
Dim nomarch As String
Dim rmesa As New ADODB.recordset

txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro1 = "SELECT * FROM " & TablaCurvas & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & " AND IDCURVA = " & idcurva & " ORDER BY FECHA"
txtfiltro2 = "SELECT COUNT(*) FROM (" & txtfiltro1 & ")"
rmesa.Open txtfiltro2, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim matcadena(1 To noreg) As String
   rmesa.Open txtfiltro1, ConAdo
   For i = 1 To noreg
       matcadena(i) = rmesa.Fields(2).GetChunk(rmesa.Fields(2).ActualSize)
       rmesa.MoveNext
       AvanceProc = i / noreg
       MensajeProc = "Leyendo la historia de una curva " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   rmesa.Close
Else
   ReDim matcadena(0 To 0) As String
End If
ObtenerArrCurvas = matcadena
End Function

Sub ObtenerHistCurvas(ByVal fecha1 As Date, ByVal fecha2 As Date, ByVal idcurva As Integer, ByVal nomcurva As String)
Dim txtfecha1 As String
Dim txtfecha2 As String
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim noreg As Integer
Dim i As Long
Dim j As Long
Dim matb() As String
Dim nomarch As String
Dim rmesa As New ADODB.recordset
Dim exitoarch As Boolean

txtfecha1 = "to_date('" & Format(fecha1, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfecha2 = "to_date('" & Format(fecha2, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro1 = "SELECT * FROM " & TablaCurvas & " WHERE FECHA >= " & txtfecha1 & " AND FECHA <= " & txtfecha2 & " AND IDCURVA = " & idcurva & " ORDER BY FECHA"
txtfiltro2 = "SELECT COUNT(*) FROM (" & txtfiltro1 & ")"
rmesa.Open txtfiltro2, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   ReDim matcadena(1 To 12001) As String
   matcadena(1) = "Plazo,"
   For i = 1 To 12000
       matcadena(i + 1) = i & ","
   Next i
   rmesa.Open txtfiltro1, ConAdo
   ReDim mata(1 To 4) As Variant
   For i = 1 To noreg
       mata(1) = rmesa.Fields(0)
       mata(2) = rmesa.Fields(1)
       mata(3) = rmesa.Fields(2).GetChunk(rmesa.Fields(2).ActualSize)
       matb = EncontrarSubCadenas(mata(3), ",")
       rmesa.MoveNext
       matcadena(1) = matcadena(1) & mata(1) & ","
       For j = 1 To 12000
           If j <= UBound(matb, 1) Then
              matcadena(j + 1) = matcadena(j + 1) & matb(j) & ","
           Else
              matcadena(j + 1) = matcadena(j + 1) & "0,"
           End If
       Next j
       AvanceProc = i / noreg
       MensajeProc = "Leyendo la historia de " & nomcurva & " " & Format(AvanceProc, "##0.00 %")
       DoEvents
   Next i
   rmesa.Close
   nomarch = DirResVaR & "\Curvas " & nomcurva & " " & Format(fecha1, "yyyy-mm-dd") & " " & Format(fecha2, "yyyy-mm-dd") & ".csv"
   frmCalVar.CommonDialog1.FileName = nomarch
   frmCalVar.CommonDialog1.ShowSave
   nomarch = frmCalVar.CommonDialog1.FileName
   Call VerificarSalidaArchivo(nomarch, 1, exitoarch)
   If exitoarch Then
      Print #1, matcadena(1)
      For i = 1 To 12000
          Print #1, matcadena(i + 1)
      Next i
      Close #1
   End If
End If
End Sub

Function ObtenerContrapNoFinSwaps(ByVal fecha As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset
Dim noreg As Long
Dim txtfecha As String
Dim i As Long
Dim mata() As String

      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      txtfiltro2 = "SELECT A.ID_CONTRAP,B.NCORTO FROM "
      txtfiltro2 = txtfiltro2 & "(SELECT ID_CONTRAP FROM " & TablaPosSwaps & " WHERE (TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN "
      txtfiltro2 = txtfiltro2 & "(SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPortPosicion
      txtfiltro2 = txtfiltro2 & " WHERE FECHA_PORT = " & txtfecha
      txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = 'DERIV SECT NO FINANCIERO') GROUP BY ID_CONTRAP ORDER BY ID_CONTRAP)"
      txtfiltro2 = txtfiltro2 & " A JOIN " & PrefijoBD & TablaContrapartes & " B ON (A.ID_CONTRAP =B.ID_CONTRAP)"
      txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
      rmesa.Open txtfiltro1, ConAdo
      noreg = rmesa.Fields(0)
      rmesa.Close
      If noreg <> 0 Then
         ReDim mata(1 To noreg, 1 To 2)
         rmesa.Open txtfiltro2, ConAdo
         For i = 1 To noreg
             mata(i, 1) = rmesa.Fields("ID_CONTRAP")
             mata(i, 2) = rmesa.Fields("NCORTO")
             rmesa.MoveNext
         Next i
         rmesa.Close
      Else
         ReDim mata(0 To 0, 0 To 0) As String
      End If
      ObtenerContrapNoFinSwaps = mata
End Function

Function ObtenerContrapNoFinDer(ByVal fecha As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset
Dim noreg As Long
Dim txtfecha As String
Dim i As Long
Dim mata() As String

      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      
      txtfiltro2 = "SELECT A.ID_CONTRAP,B.NCORTO FROM "
      txtfiltro2 = txtfiltro2 & "(SELECT ID_CONTRAP FROM (SELECT ID_CONTRAP FROM " & TablaPosSwaps & " WHERE "
      txtfiltro2 = txtfiltro2 & "(TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN "
      txtfiltro2 = txtfiltro2 & "(SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPortPosicion & " "
      txtfiltro2 = txtfiltro2 & "WHERE FECHA_PORT = " & txtfecha
      txtfiltro2 = txtfiltro2 & "AND PORTAFOLIO = 'DERIVADOS') "
      txtfiltro2 = txtfiltro2 & "UNION "
      txtfiltro2 = txtfiltro2 & "SELECT ID_CONTRAP FROM " & TablaPosFwd & " WHERE TIPOPOS =1 "
      txtfiltro2 = txtfiltro2 & "AND FECHAREG = " & txtfecha & ") "
      txtfiltro2 = txtfiltro2 & "GROUP BY ID_CONTRAP) "
      txtfiltro2 = txtfiltro2 & "A JOIN " & PrefijoBD & TablaContrapartes & " B ON (A.ID_CONTRAP =B.ID_CONTRAP) WHERE B.SECTOR = 'NF' ORDER BY ID_CONTRAP"
      txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
      rmesa.Open txtfiltro1, ConAdo
      noreg = rmesa.Fields(0)
      rmesa.Close
      If noreg <> 0 Then
         ReDim mata(1 To noreg, 1 To 2)
         rmesa.Open txtfiltro2, ConAdo
         For i = 1 To noreg
             mata(i, 1) = rmesa.Fields("ID_CONTRAP")
             mata(i, 2) = rmesa.Fields("NCORTO")
             rmesa.MoveNext
         Next i
         rmesa.Close
      Else
         ReDim mata(0 To 0, 0 To 0) As String
      End If
      ObtenerContrapNoFinDer = mata
End Function

Function ObtenerContrapFinDer(ByVal fecha As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset
Dim noreg As Long
Dim txtfecha As String
Dim i As Long
Dim mata() As String

      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      
      txtfiltro2 = "SELECT A.ID_CONTRAP,B.NCORTO FROM "
      txtfiltro2 = txtfiltro2 & "(SELECT ID_CONTRAP FROM (SELECT ID_CONTRAP FROM " & TablaPosSwaps & " WHERE "
      txtfiltro2 = txtfiltro2 & "(TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN "
      txtfiltro2 = txtfiltro2 & "(SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPortPosicion & " "
      txtfiltro2 = txtfiltro2 & "WHERE FECHA_PORT = " & txtfecha
      txtfiltro2 = txtfiltro2 & "AND PORTAFOLIO = 'DERIVADOS') "
      txtfiltro2 = txtfiltro2 & "UNION "
      txtfiltro2 = txtfiltro2 & "SELECT ID_CONTRAP FROM " & TablaPosFwd & " WHERE TIPOPOS =1 "
      txtfiltro2 = txtfiltro2 & "AND FECHAREG = " & txtfecha & ") "
      txtfiltro2 = txtfiltro2 & "GROUP BY ID_CONTRAP) "
      txtfiltro2 = txtfiltro2 & "A JOIN " & PrefijoBD & TablaContrapartes & " B ON (A.ID_CONTRAP =B.ID_CONTRAP) WHERE B.SECTOR = 'F' ORDER BY ID_CONTRAP"
      txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
      rmesa.Open txtfiltro1, ConAdo
      noreg = rmesa.Fields(0)
      rmesa.Close
      If noreg <> 0 Then
         ReDim mata(1 To noreg, 1 To 2)
         rmesa.Open txtfiltro2, ConAdo
         For i = 1 To noreg
             mata(i, 1) = rmesa.Fields("ID_CONTRAP")
             mata(i, 2) = rmesa.Fields("NCORTO")
             rmesa.MoveNext
         Next i
         rmesa.Close
      Else
         ReDim mata(0 To 0, 0 To 0) As String
      End If
      ObtenerContrapFinDer = mata
End Function

Function ObtenerContrapDer(ByVal fecha As Date)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim rmesa As New ADODB.recordset
Dim noreg As Long
Dim txtfecha As String
Dim i As Long
Dim mata() As String

      txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
      
      txtfiltro2 = "SELECT A.ID_CONTRAP,B.NCORTO FROM "
      txtfiltro2 = txtfiltro2 & "(SELECT ID_CONTRAP FROM (SELECT ID_CONTRAP FROM " & TablaPosSwaps & " WHERE "
      txtfiltro2 = txtfiltro2 & "(TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION) IN "
      txtfiltro2 = txtfiltro2 & "(SELECT TIPOPOS,FECHAREG,NOMPOS,HORAREG,CPOSICION,COPERACION FROM " & TablaPortPosicion & " "
      txtfiltro2 = txtfiltro2 & "WHERE FECHA_PORT = " & txtfecha
      txtfiltro2 = txtfiltro2 & "AND PORTAFOLIO = 'DERIVADOS') "
      txtfiltro2 = txtfiltro2 & "UNION "
      txtfiltro2 = txtfiltro2 & "SELECT ID_CONTRAP FROM " & TablaPosFwd & " WHERE TIPOPOS =1 "
      txtfiltro2 = txtfiltro2 & "AND FECHAREG = " & txtfecha & ") "
      txtfiltro2 = txtfiltro2 & "GROUP BY ID_CONTRAP) "
      txtfiltro2 = txtfiltro2 & "A JOIN " & PrefijoBD & TablaContrapartes & " B ON (A.ID_CONTRAP =B.ID_CONTRAP)ORDER BY ID_CONTRAP"
      txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
      rmesa.Open txtfiltro1, ConAdo
      noreg = rmesa.Fields(0)
      rmesa.Close
      If noreg <> 0 Then
         ReDim mata(1 To noreg, 1 To 2)
         rmesa.Open txtfiltro2, ConAdo
         For i = 1 To noreg
             mata(i, 1) = rmesa.Fields("ID_CONTRAP")
             mata(i, 2) = rmesa.Fields("NCORTO")
             rmesa.MoveNext
         Next i
         rmesa.Close
      Else
         ReDim mata(0 To 0, 0 To 0) As String
      End If
      ObtenerContrapDer = mata
End Function


Function ValidarContrapartesPosMD(ByVal fecha As Date)
'objetivo de la funcion: determinar las emisiones que no esta incluidas en los catalogos
'TablaEmxContrap,
'dato de entrada: fecha

Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim noreg As Long
Dim i As Long
Dim rmesa As New ADODB.recordset
Dim mata() As String
txtfecha = "to_date('" & Format(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT C_EMISION FROM " & TablaPosMD & " WHERE TIPOPOS = 1 AND "
txtfiltro2 = txtfiltro2 & "(TOPERACION =1 OR TOPERACION = 4) AND "
txtfiltro2 = txtfiltro2 & "(TV <> 'BI' AND TV<> 'LD' AND TV<>'IS' AND TV<>'M' AND TV <>'S' AND TV <>'IQ' AND TV<>'IM') AND "
txtfiltro2 = txtfiltro2 & "FECHAREG = " & txtfecha & " AND EMISION NOT IN "
txtfiltro2 = txtfiltro2 & "(SELECT EMISION FROM " & PrefijoBD & TablaEmxContrap & ") GROUP BY C_EMISION ORDER BY C_EMISION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim mata(1 To noreg)
   For i = 1 To noreg
       mata(i) = rmesa.Fields("C_EMISION")
       rmesa.MoveNext
   Next i
   rmesa.Close
Else
  ReDim mata(0 To 0)
End If
ValidarContrapartesPosMD = mata
End Function

Function RepPortPosEm(ByVal fecha As Date, ByVal cposicion As Integer)
Dim txtfiltro1 As String
Dim txtfiltro2 As String
Dim txtfecha As String
Dim txtport As String
Dim noreg As Long
Dim noreg1 As Long
Dim i As Long
Dim j As Long
Dim indice As Long
Dim indice1 As Long
Dim txtsubport As String
Dim txtnomarch As String
Dim txtcadena As String
Dim matvp() As New propVecPrecios
Dim mindvp() As Variant
Dim rmesa As New ADODB.recordset
Dim exito As Boolean

txtport = "TOTAL"
matvp = LeerVPrecios(fecha, mindvp)
txtfecha = "to_date('" & Format$(fecha, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txtfiltro2 = "SELECT C_EMISION,CPOSICION,TOPERACION,CALIF_2 FROM " & TablaPosMD & " WHERE FECHAREG = " & txtfecha
txtfiltro2 = txtfiltro2 & " AND CPOSICION = " & cposicion & " GROUP BY C_EMISION,CPOSICION,TOPERACION,CALIF_2 ORDER BY C_EMISION"
txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
rmesa.Open txtfiltro1, ConAdo
noreg = rmesa.Fields(0)
rmesa.Close
If noreg <> 0 Then
   rmesa.Open txtfiltro2, ConAdo
   ReDim matem(1 To noreg, 1 To 15) As Variant
   For i = 1 To noreg
       matem(i, 1) = rmesa.Fields("C_EMISION")
       matem(i, 2) = rmesa.Fields("CPOSICION")
       matem(i, 3) = rmesa.Fields("TOPERACION")
       matem(i, 4) = rmesa.Fields("CALIF_2")
      rmesa.MoveNext
   Next i
   rmesa.Close
   For i = 1 To noreg
       txtsubport = "EM " & matem(i, 1) & " POS " & matem(i, 2) & " T_OP " & matem(i, 3)
       txtfiltro2 = "SELECT * FROM " & TablaValPosPort & " WHERE FECHAP = " & txtfecha
       txtfiltro2 = txtfiltro2 & " AND PORTAFOLIO = '" & txtport & "'"
       txtfiltro2 = txtfiltro2 & " AND SUBPORT = '" & txtsubport & "'"
       txtfiltro1 = "SELECT COUNT(*) FROM (" & txtfiltro2 & ")"
       rmesa.Open txtfiltro1, ConAdo
       noreg1 = rmesa.Fields(0)
       rmesa.Close
       If noreg1 <> 0 Then
          rmesa.Open txtfiltro2, ConAdo
          matem(i, 5) = rmesa.Fields("MTM_SUCIO")
          matem(i, 6) = rmesa.Fields("DV01_ACT")
          matem(i, 7) = rmesa.Fields("N_TITULOS_A")
          matem(i, 8) = rmesa.Fields("DUR_ACT")
          matem(i, 15) = rmesa.Fields("P_ESPERADA")
          rmesa.Close
       End If
       indice = BuscarValorArray(matem(i, 1), mindvp, 1)
       If indice <> 0 Then
          indice1 = mindvp(indice, 2)
          matem(i, 9) = matvp(indice1).regla_cupon
          matem(i, 10) = matvp(indice1).tcupon
          matem(i, 11) = matvp(indice1).yield
          matem(i, 12) = matvp(indice1).frec_cupon
          matem(i, 13) = matvp(indice1).fvenc
       End If
       matem(i, 14) = LeerResVaR(fecha, txtport, "Normal", txtsubport, 500, 1, 0, 0.03, 0, "CVARH", exito)
    Next i
    ReDim mata(1 To noreg, 1 To 15) As Variant
    For i = 1 To noreg
        mata(i, 1) = fecha
        mata(i, 2) = matem(i, 1)                'clave de emision
        mata(i, 3) = matem(i, 4)                'calificacion
        mata(i, 4) = matem(i, 5) / 1000000      'mtm en millones
        mata(i, 5) = matem(i, 6) / 1000000      'dv01 en millones
        mata(i, 6) = matem(i, 9)                'regla cupon
        mata(i, 7) = matem(i, 10) / 100         'tasa cupon
        mata(i, 8) = matem(i, 11) / 100         'yield
        mata(i, 9) = matem(i, 12)               'frecuencia cupon
        mata(i, 10) = matem(i, 8)               'duracion
        mata(i, 11) = matem(i, 13)              'fecha de vencimiento
        mata(i, 12) = -matem(i, 14) / 1000000   'cvar
        mata(i, 13) = matem(i, 15) / 1000000    'perdida esperada
        mata(i, 14) = matem(i, 7)               'no de titulos
        mata(i, 15) = 0
    Next i
RepPortPosEm = mata
End If

End Function

