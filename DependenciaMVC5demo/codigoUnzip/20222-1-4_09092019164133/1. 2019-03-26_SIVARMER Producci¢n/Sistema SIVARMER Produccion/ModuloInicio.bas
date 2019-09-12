Attribute VB_Name = "ModuloInicio"
Option Explicit

Sub Main()
Dim i As Integer
OpcionBDatos = 1                            '1 produccion, 2 desarrollo, 3 drp
BDIKOS = 1                                  '1 produccion, 2 desarrollo, 3 drp
ActivarControlErrores = False
ValEficiencia = False
SiHacer = True
SiDepuracion = False
VersExcel = "Excel 12.0;"                   'excel 2007
NomServQA = "SIVARMERD"
NomServP = "alm2"
PrefijoBD = "riesgo."                              'prefijo para acceder a la tablas de acuerdo al usuario
PrefijoBD = ""                                     'prefijo para acceder a la tablas de acuerdo al usuario
SiCargaFEmMD = True
ClavePosMD = 1
ClavePosTeso = 2
ClavePosMC = 3
ClavePosDeriv = 4
ClavePosPension1 = 5
ClavePosPension2 = 6
ClavePosDeuda = 7
ClavePosPIDV = 8
ClavePosPICV = 9
ClavePosPID = 10
ClavePosPenMD = 11

ClaveCDirec = 1
ClaveVDirec = 4
BlockSize = 8000

txtCadCarEsp = "!$%&/()=?¿¡'-_.:,;<>|\+*{}Çºª" & Chr(34)
txtCadNum = "1234567890"
txtCadMin = "abcdefghijklmnopqrstuvwxyz"
txtCadMay = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


ReDim MatHora(1 To 2) As Date
MatHora(1) = #9:00:00 AM#
MatHora(2) = #9:30:00 AM#

If ActivarControlErrores Then
  On Error GoTo hayerror
End If
Screen.MousePointer = 11
'se definen algunos parametros iniciales
'Nota la matriz que va a contener todos los datos de var y
'Call DeclararPlazos
'CLAVES DE LAS POSICIONES
'1   MESA DE DINERO
'2   tesorería
'3   MESA DE CAMBIOS
'4   MESA DE DERIVADOS
'5   MESA DE DERIVADOS
'6   MESA DE DERIVADOS
'7   MESA DE DERIVADOS
'8   FONDO DE PENSIONES 1
'9   FONDO DE PENSIONES 2

SRutOrden = 1    '1 quicksort, 2 superorden, 3 metodo de la burbuja
NombrePortFR = "PRUEBA 2"
txtportCalc1 = "TOTAL"
txtportCalc2 = "NEGOCIACION + INVERSION"
txtportBanobras = "CONSOLIDADO"
SiGuardoDBMexico = False
SiAgregarDatosFwd = False
SiAnexarFlujosSwaps = False
SiIncTasaCVig = True                  'usar tasa del cupon vigente de la tabla original
Pi = 3.14159265358979
'variables para el resumen de posicion
'VARIABLES PARA EL RESUMEN DE VAR
Call DeclaraTablas
Call DeclararTablasCatalogos
IdPosPension = 12
IdPosPension2 = 13
ValExacta = False
TipoRendimiento = "l"        '1 rendimientos logaritmicos
PrecioLimpio = 0             'precios limpios
SiConsolidaPos = True
SiCargoFactR = False
SiCargaFechasPosicion = True
SiCargaFechasPrecios = True
SiCargaTasas = True
ContraseñaCatalogos = "dragonfly"

DirBases = "\\STFP3MJ002F23\Bases"
Call IniciarConexOracle(ConAdo, OpcionBDatos)
'esta conexion garantiza la interaccion con ikos
Call IniciarConexOracle(conAdoBD, BDIKOS)

frmLogin.Show 1
For i = 1 To UBound(MatUsuarios, 1)
   If NoIntFallidos(i) <> 0 Then
      Call GuardarIFBit(Date, Time, MatUsuarios(i, 2), NoIntFallidos(i))
   End If
Next i
If LoginSucceeded Then
Call ActUHoraUsuario
SiActTProc = False
'se mantiene una conexion permanente a las bases de oracle
If Not ejecutaSubproc1 And Not ejecutaSubproc2 And Not ejecutaSubproc3 Then Call BloquearUsuario(NomUsuario)
Call RegistrarInicioSesion


'el Servidor de Driesgos
  ConAdo.Execute "ALTER SESSION SET NLS_DATE_FORMAT = 'DD/MM/YYYY'"
  'PODER insertar registros en bases indexadas sin error
  ConAdo.Execute "ALTER SESSION SET SKIP_UNUSABLE_INDEXES = True"

ReDim txtCadenaMensajes(1 To 1) As Variant
If (PerfilUsuario = "ADMINISTRADOR") And Not ejecutaSubproc1 And Not ejecutaSubproc2 And Not ejecutaSubproc3 Then
   frmProgreso.Show
   Call AbrirTablas
   Call LeerCatalogos        'CARGA LOS CATALOGOS DEL SISTEMA
   MatFechasVaR = LeerFechasVaRT()
   MatFechasFR = LeerFechasFRT()
   MatFechasTareas1 = LeerFechasTareasT()
   MatFechasTareas2 = LeerFechasTareasPosT()
   SiCargaTasas = True
'se cargan los plazos de las curvas
   SiCargoFactR = False
   Unload frmProgreso
   frmCalVar.Show
   MensajeProc = NomUsuario & " ha accedido al sistema"
   Call GuardaAccesoBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc)
ElseIf (PerfilUsuario = "ADMINISTRADOR" Or PerfilUsuario = "USUARIO") And ejecutaSubproc1 Then
   frmProgreso.Show
   Call AbrirTablas
   Call LeerCatalogos        'CARGA LOS CATALOGOS DEL SISTEMA
   MatFechasVaR = LeerFechasVaRT()
   MatFechasFR = LeerFechasFRT()
   MatFechasTareas1 = LeerFechasTareasT()
   MatFechasTareas2 = LeerFechasTareasPosT()
   SiCargaTasas = True
'se cargan los plazos de las curvas
   SiCargoFactR = False
   Unload frmProgreso
   frmEjecSubproc1.Show
   MensajeProc = NomUsuario & " ha accedido al sistema"
   Call GuardaAccesoBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc)
   
ElseIf (PerfilUsuario = "ADMINISTRADOR" Or PerfilUsuario = "USUARIO") And ejecutaSubproc2 Then
   frmProgreso.Show
   Call AbrirTablas
   Call LeerCatalogos        'CARGA LOS CATALOGOS DEL SISTEMA
   MatFechasVaR = LeerFechasVaRT()
   MatFechasFR = LeerFechasFRT()
   MatFechasTareas1 = LeerFechasTareasT()
   MatFechasTareas2 = LeerFechasTareasPosT()
   SiCargaTasas = True
'se cargan los plazos de las curvas
   SiCargoFactR = False
   Unload frmProgreso
   frmEjecSubproc2.Show
   MensajeProc = NomUsuario & " ha accedido al sistema"
   Call GuardaAccesoBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc)
ElseIf (PerfilUsuario = "ADMINISTRADOR" Or PerfilUsuario = "USUARIO") And ejecutaSubproc3 Then
   frmProgreso.Show
   Call AbrirTablas
   Call LeerCatalogos        'CARGA LOS CATALOGOS DEL SISTEMA
   MatFechasVaR = LeerFechasVaRT()
   MatFechasFR = LeerFechasFRT()
   MatFechasTareas1 = LeerFechasTareasT()
   MatFechasTareas2 = LeerFechasTareasPosT()
   SiCargaTasas = True
'se cargan los plazos de las curvas
   SiCargoFactR = False
   Unload frmProgreso
   frmEjecSubproc3.Show
   MensajeProc = NomUsuario & " ha accedido al sistema"
   Call GuardaAccesoBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc)
ElseIf PerfilUsuario = "REPORTES" Or PerfilUsuario = "USUARIO" Then
      MatFechasVaR = LeerFechasVaRT
      frmReportes.Show
      MensajeProc = NomUsuario & " ha accedido al sistema"
      Call GuardaAccesoBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc)
ElseIf PerfilUsuario = "ADMUSUARIOS" Then
      frmUsuarios.Show
      MensajeProc = NomUsuario & " ha accedido al sistema"
      Call GuardaAccesoBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc)
ElseIf PerfilUsuario = "BITACORA" Then
      MatFechasVaR = LeerFechasVaRT
      frmBitacora.Show
      MensajeProc = NomUsuario & " ha accedido al sistema"
      Call GuardaAccesoBitacora(1, "Acceso", 0, MensajeProc, NomUsuario, Date, MensajeProc)
End If
Else
 End
End If
Screen.MousePointer = 0
On Error GoTo 0
Exit Sub
hayerror:
 Call TratamientoErrores(Err())
 On Error GoTo 0
End Sub

Sub AbrirTablas()
On Error Resume Next
RGuardarPL.Open "SELECT * FROM " & TablaPLHistOper, ConAdo, 1, 3
RGuardarPLMont.Open "SELECT * FROM " & TablaPyGMontOper, ConAdo, 1, 3
RegResCVA.Open "SELECT * FROM " & TablaPLEscCVA & "", ConAdo, 1, 3
RegResMakeW.Open "SELECT * FROM " & TablaPLEscMW & "", ConAdo, 1, 3
RegResLimC1.Open "SELECT * FROM " & TablaLimContrap1, ConAdo, 1, 3
RegResLimC2.Open "SELECT * FROM " & TablaLimContrap2, ConAdo, 1, 3
RegResCVAMD.Open "SELECT * FROM " & TablaPYGCVAMD, ConAdo, 1, 3

On Error GoTo 0
End Sub

Sub CerrarTablas()
   On Error Resume Next
   RGuardarPL.Close
   RGuardarPLMont.Close
   RegResCVA.Close
   RegResLimC1.Close
   RegResLimC2.Close
   RegResCVAMD.Close
   On Error GoTo 0
End Sub

Sub RegistrarInicioSesion()
Dim txtfecha As String
Dim txtcadena As String
Dim txthora As String
Dim txtipdir As String
txtipdir = RecuperarIP()
Id_Sesion = CLng(Date) & CLng(Time) & Left(NomUsuario, 2)
txtfecha = "to_date('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txthora = "to_date('" & Format(Time, "hh:mm:ss") & "','HH24:MI:SS')"
txtcadena = "INSERT INTO " & TablaSesiones & " VALUES("
txtcadena = txtcadena & "'" & Id_Sesion & "',"
txtcadena = txtcadena & "'" & NomUsuario & "',"
txtcadena = txtcadena & "'" & txtipdir & "',"
txtcadena = txtcadena & txtfecha & ","
txtcadena = txtcadena & txthora & ","
txtcadena = txtcadena & "null,"
txtcadena = txtcadena & "null,"
txtcadena = txtcadena & txtfecha & ","
txtcadena = txtcadena & txthora & ")"
ConAdo.Execute txtcadena

End Sub



Sub BloquearUsuario(ByVal txtusuario As String)
Dim txtfecha As String
Dim txtcadena As String
Dim txthora As String
Dim txtipdir As String
txtfecha = "to_date('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txthora = "to_date('" & Format(Time, "hh:mm:ss") & "','HH24:MI:SS')"
txtipdir = RecuperarIP()
txtcadena = "UPDATE " & TablaUsuarios & " SET ENLINEA = 'S', FENTRADA = " & txtfecha & ", HENTRADA = " & txthora & ", FUREPORTE = " & txtfecha & ", HUREPORTE = " & txthora & ", DIRECCION_IP = '" & txtipdir & "' WHERE USUARIO = '" & txtusuario & "'"
ConAdo.Execute txtcadena
End Sub

Sub DesbloquearUsuario(ByVal txtusuario As String)
Dim txtfecha As String
Dim txtcadena As String
Dim txthora As String
txtfecha = "to_date('" & Format(Date, "dd/mm/yyyy") & "','dd/mm/yyyy')"
txthora = "to_date('" & Format(Time, "hh:mm:ss") & "','HH24:MI:SS')"
txtcadena = "UPDATE " & TablaUsuarios & " SET ENLINEA = 'N', FSALIDA = " & txtfecha & ", HSALIDA = " & txthora & " WHERE USUARIO = '" & txtusuario & "'"
ConAdo.Execute txtcadena
End Sub

Sub DeclaraTablas()

    TablaValDeriv = PrefijoBD & "VAR_VALDERIVADOS1"
    TablaVDerIKOS = PrefijoBD & "VAR_VALDER_IKOS"
    TablaBitacora = PrefijoBD & "VAR_TD_BITACORA"
    TablaBitacoraIF = PrefijoBD & "VAR_BITACORAIF"
    TablaAnalisisFRO = PrefijoBD & "VAR_TD_VOLATIL"
    TablaUsuarios = PrefijoBD & "VAR_USUARIOS1"
    TablaSesiones = PrefijoBD & "VAR_SESIONES_ACTIVAS"
    TablaSensibN = PrefijoBD & "VAR_TD_SENSIBILIDADES"
    TablaDetalleMo = PrefijoBD & "VAR_MO"
    
    'tablas de posicion
    TablaPosMD = PrefijoBD & "VAR_POS_MD_4"
    TablaFlujosMD = PrefijoBD & "VAR_FLUJOS_MD"
    TablaPosDiv = PrefijoBD & "VAR_POS_DIV_6"
    TablaPosSwaps = PrefijoBD & "VAR_POS_SWAPS_3"
    TablaFlujosSwapsO = PrefijoBD & "VAR_FLUJOS_SWAPS_3"
    TablaPosFwd = PrefijoBD & "VAR_POS_FWDS_6"
    TablaPosFwdFrutos = PrefijoBD & "VAR_POS_FWD"
    TablaPosDeuda = PrefijoBD & "VAR_POS_DEUDA_4"
    TablaFlujosDeudaO = PrefijoBD & "VAR_FLUJOS_DEUDA_4"
    TablaFlujosRichard = PrefijoBD & "VAR_FLUJOSD"
    TablaResVaR = PrefijoBD & "VAR_TD_RES_VAR_2"
    TablaFRiesgoO = PrefijoBD & "VAR_F_RIESGO"                   'tabla de trabajo oracle
    TablaCalendSwapsO = PrefijoBD & "VAR_CALEND_SWAPS"
    TablaBaseCalendO = PrefijoBD & "VAR_BASE_CALEND"
    TablaVecPrecios = PrefijoBD & "VAR_TD_VEC_PRECIOS"
    TablaEficSwaps = PrefijoBD & "VAR_EFIC_SWAPS"
    TablaValExtO = PrefijoBD & "VAR_BADHOC1"
    TablaResBack = PrefijoBD & "VAR_TD_PL_BACK"
    TablaBackPort = PrefijoBD & "VAR_TD_PL_BACK_PORT_POS"
    TablaValPos = PrefijoBD & "VAR_TD_VAL_POS"
    TablaValPosPort = PrefijoBD & "VAR_TD_VAL_POS_PORT"
    TablaPLHistOper = PrefijoBD & "VAR_TD_PL_HIST_2"
    TablaPLEscHistPort = PrefijoBD & "VAR_TD_PL_HIST_PORT_2"
    TablaPLEscCVA = PrefijoBD & "VAR_TD_PL_CVA_2"
    TablaPLEscMW = PrefijoBD & "VAR_TD_PL_HIST_MW"
    TablaFactChol = PrefijoBD & "VAR_TD_CHOL_MONT"
    TablaResEfectPros = PrefijoBD & "VAR_TD_RES_EFECT_PROS"
    TablaResMO = PrefijoBD & "VAR_RES_MO"
    
    TablaPYGCVAMD = PrefijoBD & "VAR_TD_PL_CVA_MD"
    TablaPyGMontOper = PrefijoBD & "VAR_TD_PL_MONT"
    TablaPyGMontPort = PrefijoBD & "VAR_TD_PL_MONT_PORT_POS"
    TablaProcesos1 = PrefijoBD & "VAR_TD_PROCESOS"
    TablaProcesos2 = PrefijoBD & "VAR_TD_PROCESOS_2"
    TablaSubProcesos1 = PrefijoBD & "VAR_TD_SUBPROCESOS"
    TablaSubProcesos2 = PrefijoBD & "VAR_TD_SUBPROCESOS_2"
    TablaSubProcesos3 = PrefijoBD & "VAR_TD_SUBPROCESOS_3"
    TablaParamUsuario = PrefijoBD & "VAR_TD_PARAM_USUARIO"
    TablaFechasVaR = PrefijoBD & "VAR_FECHAS_VAR"
    TablaFechasFR = PrefijoBD & "VAR_FECHAS_FR"
    TablaFechasTareas1 = PrefijoBD & "VAR_FECHAS_TAREAS"
    TablaDerivEst2 = PrefijoBD & "VAR_DER_ESTANDAR"

    TablaEficRetro = PrefijoBD & "VAR_EFIC_RETRO"
    TablaSensibPort = PrefijoBD & "VAR_TD_SENSIB_PORT"
    TablaPLHistOperVR = PrefijoBD & "VAR_TD_PL_HIST_VR2"
    TablaPyGHistPVR = PrefijoBD & "VAR_TD_PL_HIST_PORT_VR2"
    TablaResVReemplazo = PrefijoBD & "VAR_TD_RES_V_REEMPLAZO"
    TablaResCalcVReemplazo = PrefijoBD & "VAR_TD_RES_C_REEMPLAZO"
    
    TablaPosDivCon = PrefijoBD & "VAR_POSX_DIV"
    TablaNumDistNormal = PrefijoBD & "VAR_TD_SIM_NORM_MONT"
    TablaValContraparte = PrefijoBD & "VAR_VAL_CONTRAPARTE"
    TablaResCVA = PrefijoBD & "VAR_TD_RES_CVA"
    TablaCurvas = PrefijoBD & "VAR_TD_CURVAS"

    TablaResEscEstres = PrefijoBD & "VAR_TD_ESC_ESTRES"
    TablaResEscEstresPort = PrefijoBD & "VAR_TD_ESC_ESTRES_PORT"
    TablaResEstresAprox = PrefijoBD & "VAR_TD_RES_ESTRES_APROX"
    TablaPortPosicion = PrefijoBD & "VAR_TD_PORT_POS_3"
    TablaResWRW = PrefijoBD & "VAR_TD_REP_WRW"
    TablaEscFR = PrefijoBD & "VAR_TD_ESC_FR"
    TablaLimContrap1 = PrefijoBD & "VAR_TD_RES_LIM_CONTRAPNF_1"
    TablaLimContrap2 = PrefijoBD & "VAR_TD_RES_LIM_CONTRAPNF_2"
    TablaResLimContrap = PrefijoBD & "VAR_TD_LIM_CONTRAP_NF"
    TablaExpFwds = PrefijoBD & "VAR_TD_EXP_FWDS"
    TablaMovBalanza = PrefijoBD & "VAR_TD_MOV_BALANZA"
    
    TablaCatalogosA = "VAR_CATALOGOSN.MDB"
    
    'interfaces
    TablaPosMesaIKOS = "POS_MESA_TESO"
    TablaInterfCarac = "SW_RIESGOS2@ADERIVAD"
    TablaInterfFlujos = "SW_RIESGOS1@ADERIVAD"
    TablaInterfFwd = "OPER_FWDS@aderivad"
    TablaInterfDiv = "cofini.v_posicion_cambiaria@consulta_sicofin"  'NO SE ESTA UTILIZANDO
    TablaInterfSim1 = "VAR_FSWAPSSIM@ADERIVAD"
    TablaInterfSim2 = "FWD_OPERSIM@ADERIVAD"
    TablaDerEstandar = PrefijoBD & "DERIVADO_ESTANDAR"
    
    TablaVaRIKOS = PrefijoBD & "SW_VAR"
    TablaEficienciaCob = PrefijoBD & "EFICIENCIA_COB"
    TablaOperValidada = PrefijoBD & "VAR_TD_OPER_VALIDADA"
  
    'TablaInterfSim1 = "VAR_POS_INT_SWAPS_SIM"      'clon de VAR_FSWAPSSIM@ADERIVAD
    'TablaInterfSim2 = "VAR_POS_INT_FWD_SIM"        'clon de FWD_OPERSIM@ADERIVAD
    'TablaVaRIKOS = "VAR_SW_VAR"                    'clon de SW_VAR
    'TablaEficienciaCob = PrefijoBD & "VAR_EFICIENCIA_COB"     'clon de EFICIENCIA_COB
    'TablaOperValidada = PrefijoBD & "VAR_TD_OPER_VALIDADA_SIM"   'clon de VAR_TD_OPER_VALIDADA

End Sub

Sub DeclararTablasCatalogos()
    TablaCEmision = "VAR_CLAVEEMISION1"
    TablaPEmision = "VAR_PARAMEMISION1"
    TablaLimites = "VAR_LIMITES"
    TablaMTrans = "VAR_MTRANSICION"
    TablaGruposDeriv = "VAR_GRUPOS_DERIV1"
    TablaEmxContrap = "VAR_EMXCONTRAP"
    TablaCatProcesos = "VAR_TC_PROCESOS_3"
    TablaIndVecPreciosO = "VAR_TC_IND_VPRECIOS"
    TablaBLTresh = "VAR_TC_DERIV_SIN_LMARGEN"
    TablaPortPrincipales = "VAR_TC_PORT_PRIN"
    TablaSecProcesos = "VAR_TC_SEC_PROCESOS_3"
    TablaSecSubProc = "VAR_TC_SEC_SUBPROC"
    TablaParamSistema = "VAR_TC_PARAMETROS"
    TablaGruposPortPos = "VAR_TC_GRUPOS_PORT_POS_3"
    TablaReporteCVaR = "VAR_TC_REPORTE_VAR"
    TablaHistCurvas = "VAR_TC_HIST_CURVAS"
    TablaRecNacional = "VAR_TC_RECUPERA_N"
    TablaRecInt = "VAR_TC_RECUPERA_I"
    TablaBlackList = "VAR_BLACKLIST"
    TablaEquivContrap = "VAR_TC_CONTRAP_EQUIV"
    TablaContrapartes = "VAR_TC_CONTRAP"
    TablaTreshCont = "VAR_TC_CONTRAP_PARAM_CVA"
    TablaSectorEscEm = "VAR_TC_SECTOR_ESC_EM"
    TablaEscCortoLargo = "VAR_TC_ESC_CORTO_LARGO"
    TablaCalifContrapF = "VAR_TC_CONTRAP_F_CALIF"
    TablaCalifContrapNF = "VAR_TC_CONTRAP_NF_CALIF"
    TablaCalifContrapEmision = "VAR_TC_EMISIONES_CALIF"
    TablaCatCurvas = "VAR_TC_CURVAS_2"
    TablaMOEmSectorMD = "VAR_TC_MO_EM_SECTOR_MD"
    TablaMOEmSectorPI = "VAR_TC_MO_EM_SECTOR_PI"
    TablaMOContrapCalifPI = "VAR_TC_MO_CONTRAP_CALIF_PI"
    TablaMOContrap = "VAR_TC_MO_CONTRAP"
    TablaMOperCalif = "VAR_TC_MO_CALIF"
    TablaMOEmPriv = "VAR_TC_MO_EM_PRIV"
    TablaMOMon = "VAR_TC_MO_MON"
    TablaMOGub = "VAR_TC_MO_GUB"
    TablaCatPortPos = "VAR_TC_PORT_POS"
    TablaPortPosEstructural = "VAR_TC_PORT_POS_ESTRUCTURAL"
    TablaValBonosC0 = "VAR_TC_VAL_BC0"
    TablaValBonos = "VAR_TC_VAL_BONOS_2"
    TablaValReportos = "VAR_TC_VAL_REPORTOS_2"
    TablaValFwds1 = "VAR_TC_VAL_FWDTC1"
    TablaValFwds2 = "VAR_TC_VAL_FWDTC2"
    TablaSQLPort = "VAR_TC_SQL_PORT_3"
    TablaGruposPapelFP = "VAR_TC_PORT_T_PAPEL_FP"
    TablaPortFR = "VAR_PORT_FR_2"
    TablaCalificaciones = "VAR_TC_CALIFICACION"
    TablaValInds = "VAR_TC_VAL_IND"
    TablaValBSC = "VAR_TC_VAL_ST_CUPON_2"
    TablaValBSD = "VAR_TC_VAL_ST_DESC"
    TablaValSwaps1 = "VAR_TC_VAL_SWAPS1"
    TablaValSwaps2 = "VAR_TC_VAL_SWAPS2"
    TablaValDeuda = "VAR_TC_VAL_DEUDA"
    TablaRSwapsDeuda = "VAR_TC_REL_SWAPS_DEUDA"
    TablaRelSwapIKOSS = "VAR_TC_REL_S_IKOS_S"
    TablaBlEsc = "VAR_TC_BL_ESC"
    TablaFechasEscEstres = "VAR_TC_FECHAS_ESTRES"
    TablaSplits = "VAR_FRSPLIT"
    TablaMonedas = "VAR_MONEDAS"
    TablaNodosCurvas = "VAR_TC_NODOS_CURVAS"
    TablaPosPrimarias = "VAR_TC_POS_PRIM_2"
    TablaRelSwapEm = "VAR_TC_DER_EM_PIDV"
    TablaCalendSwaps = "VAR_REL_SWAP_CALEND"

End Sub


Sub LeerCatalogos()
'los tipos de titulos
MatParamSistema = LeerParametrosSist()
 Call ValidarParamUsuario(NomUsuario)             'parametros establecidos para el usuario
 Call CargaLimites        'limites operativos
 MatListaPortPos = LeerListaPortPos()
 If UBound(MatListaPortPos, 1) > 0 Then
 'este es el desglose del VaR sin considerar la posicion de derivados
    MatEstRepVaRB = CargaGruposPortPos(MatListaPortPos(1, 1))
 End If
 'matclasderivados tiene la finalidad de simplificar la valuacion de instrumentos
 MatPortCurvas = CargaPortCurvas()
 MatContrapartes = CargaContrapartes()
 
 MatClavesContrap = CargaClavesContrap()
 MatPortPosicion = CargaPortPos()
 MatIndVPrecios = CargaIndVPrecios()
 MatHistCurvas = LeerHistCurvas()
 MatCatCurvas = LeerTCCurvasO()
 
    MatFechasEstres = CargaTProd(PrefijoBD & TablaFechasEscEstres, "Fechas de escenarios de estres")
    MatTValBC0 = CargaTProd(PrefijoBD & TablaValBonosC0, "VALUACION BC0")
    MatTValBonos = CargaTProd(PrefijoBD & TablaValBonos, "VALUACION Bonos")
    MatTValReportos = CargaTProd(PrefijoBD & TablaValReportos, "Valuacion de Reportos")
    MatTValSTCupon = CargaTProd(PrefijoBD & TablaValBSC, "Valuación de bonos ST cupon")
    MatTValSTDesc = CargaTProd(PrefijoBD & TablaValBSD, "VALUACION ST DESC")
    MatTValSwaps1 = CargaTablaD(PrefijoBD & TablaValSwaps1, "Valuación de swaps", 1)
    MatTValSwaps2 = CargaTablaD(PrefijoBD & TablaValSwaps2, "Valuación de swaps", 1)
    MatTValFwdsTC1 = CargaTablaD(PrefijoBD & TablaValFwds1, "Valuacion fwd TC", 1)
    MatTValFwdsTC2 = CargaTablaD(PrefijoBD & TablaValFwds2, "Valuacion fwd TC", 1)
    MatReporteCVaR = CargaTablaD(PrefijoBD & TablaReporteCVaR, "Reportes de CVaR", 1)
    MatPortSegRiesgo = DefinePortSegRiesgo
    MatTValDeuda = CargaTablaD(PrefijoBD & TablaValDeuda, "Valuación de posiciones de deuda", 1)
    MatTValInd = CargaTProd(PrefijoBD & TablaValInds, "Valuacion de indices")
    MatMonedas = CargaTablaD(PrefijoBD & TablaMonedas, "Catalogo de Monedas", 1)
    MatFRSplit = CargaTablaD(PrefijoBD & TablaSplits, "Factores riesgo split", 1)
    MatEmxContrap = CargaTProd(PrefijoBD & TablaEmxContrap, "EMISION X CONTRAPARTE")
    MatGruposDeriv = CargaTablaD(PrefijoBD & TablaGruposDeriv, "GRUPOS DERIVADOS", 1)
    MatCalificaciones = CargaTablaD(PrefijoBD & TablaCalificaciones, "Escala de Calificaciones", 1)
    MatPortDeriv = DefPortDerivados
    MatCatProcesos = LeerCatProcesos()
    MatSecProcesos = LeerSecuenciaProcesos()
    MatSecSubproc = LeerSecuenciaSubproc()
    MatClavesEmision = CargaClaveEmision()
    Call LeerNodosCurvas
    SiCargaTasas = False
    MatRelSwapsPrim = CargaRelSwapPrim()
    MatParamEmisiones = CargaParamEmisiones()
    MatDerEstandar = CargaDervEstandar()
    MatRelSwapIS = CargaTablaD(PrefijoBD & TablaRelSwapIKOSS, "Relacion swaps IKOS SIVARMER", 1)
    MatRelSwapsDeuda = CargaTablaD(PrefijoBD & TablaRSwapsDeuda, "Relacion swaps deuda", 1)
    MatBlEsc = CargaTablaD(PrefijoBD & TablaBlEsc, "Escenario hist a omitir", 1)
    MatSQLPort = CargaTablaD(PrefijoBD & TablaSQLPort, "SQL de definicion de portafolios", 1)
    MatGruposPapelFP = CargaTablaD(PrefijoBD & TablaGruposPapelFP, "Grupos de papel fondo de pensiones", 1)
    MatTresholdContrap = CargaTablaD(PrefijoBD & TablaTreshCont, "Parametros de contrapartes para CVA", 1)
    MatRelSwapEm = CargaTablaD(PrefijoBD & TablaRelSwapEm, "Derivados asociados a emisiones", 1)
    MatPortEstruct = LeerPortPosEstruc()
   
    Call LeerPortafolioFRiesgo(NombrePortFR, MatCaracFRiesgo, NoFactores)
    Call CargaCatMO
 'MatCalendSwaps = CargaTablaD(TablaCalendSwapsO , "Base cal swaps", 2)
 'MatBaseCalend = CargaTablaD(TablaBaseCalendO , "Base calculo Calendarios", 2)
 'MatRelSwapsCal = CargaTablaD(prefijobd & TablaCalendSwaps , "Relacion Swaps Calendarios", 2)
    MatPortafolios = CargaPortafolios()
    NoPortafolios = UBound(MatPortafolios, 1)
    MBList = LeerBlackList()
    fechavalIKOS = 0
    
End Sub

Sub CargaCatMO()
    MatMOSectorMD = CargaTablaD(PrefijoBD & TablaMOEmSectorMD, "Sectores de Marco de Operacion", 1)
    MatMOSectorPI = CargaTablaD(PrefijoBD & TablaMOEmSectorPI, "Sectores de Marco de Operacion", 1)
    MatMOCalif = CargaTablaD(PrefijoBD & TablaMOperCalif, "Calificaciones Marco de Operacion", 1)
    MatMOEmPriv = CargaTablaD(PrefijoBD & TablaMOEmPriv, "Calificaciones Marco de Operacion", 1)
    MatMOMon = CargaTablaD(PrefijoBD & TablaMOMon, "Calificaciones Marco de Operacion", 1)
    MatMOContrapCalifPI = CargaTablaD(PrefijoBD & TablaMOContrapCalifPI, "Tipo de inst Calif PI", 1)
    MatMOContrap = CargaTablaD(PrefijoBD & TablaMOContrap, "Tipo de inst Calif PI", 1)
    MatMOGub = CargaTablaD(PrefijoBD & TablaMOGub, "Marco de oper Gub", 1)
End Sub

Function CargaDervEstandar()
'carga la lista de derivados estandar manejados por el sistema IKOS Derivados
' con esta tabla el sistema de riesgos puede determinar el tipo de swap que se quiere validar
'en funcion de la clave nombreswap
'devuelve un array con los datos de esta tabla
On Error GoTo hayerror
    Dim txtfiltro As String, txtfiltro1 As String
    Dim contar    As Integer, i As Integer, noreg As Integer, nocampos As Integer
    Dim rprecios As New ADODB.recordset

    txtfiltro = "select * from " & TablaDerEstandar
    txtfiltro1 = "select count(*) from (" & txtfiltro & ")"
    rprecios.Open txtfiltro1, conAdoBD
    noreg = rprecios.Fields(0)
    rprecios.Close
    If noreg <> 0 Then
        rprecios.Open txtfiltro, conAdoBD
        nocampos = rprecios.Fields.Count
        ReDim mata(1 To noreg, 1 To nocampos) As Variant
        rprecios.MoveFirst
        contar = 0
        For i = 1 To noreg
            mata(i, 1) = rprecios.Fields(0)
            mata(i, 2) = rprecios.Fields("NOMBRESWAP")
            mata(i, 3) = rprecios.Fields(2)
            mata(i, 4) = rprecios.Fields(3)
            mata(i, 5) = rprecios.Fields(4)
            mata(i, 6) = rprecios.Fields(5)
            mata(i, 7) = rprecios.Fields(6)
            mata(i, 8) = Replace(rprecios.Fields(7), " ", "")
            mata(i, 9) = Replace(rprecios.Fields(8), " ", "")
            mata(i, 10) = rprecios.Fields(9)
            mata(i, 11) = rprecios.Fields(10)
            mata(i, 12) = rprecios.Fields(11)
            mata(i, 13) = rprecios.Fields(12)
            mata(i, 14) = rprecios.Fields(13)
            mata(i, 15) = rprecios.Fields(14)
            mata(i, 16) = rprecios.Fields(15)
            mata(i, 17) = rprecios.Fields(16)
            mata(i, 18) = rprecios.Fields(17)
            mata(i, 19) = rprecios.Fields(18)
            mata(i, 20) = rprecios.Fields(19)
            rprecios.MoveNext
            AvanceProc = i / noreg
            MensajeProc = "Catalogo de derivados estandar " & " " & Format$(AvanceProc, "##0.00 %")
        Next i

        rprecios.Close
        mata = RutinaOrden(mata, 2, SRutOrden)
    Else
        ReDim mata(0 To 0, 0 To 0) As Variant
    End If

    CargaDervEstandar = mata
On Error GoTo 0
Exit Function
hayerror:
ReDim mata(0 To 0, 0 To 0) As Variant
CargaDervEstandar = mata
MsgBox error(Err())
End Function

Function CargaPortafolios()
  CargaPortafolios = CargaTablaD(PrefijoBD & TablaPortPrincipales, "Portafolios para el VaR", 1)
End Function


Sub IniciarConexOracle(ByRef conex As ADODB.Connection, ByVal opcion As Integer)
If opcion = 1 Then
   Call IniciarConexOracleP(conex)
ElseIf opcion = 2 Then
   Call IniciarConexOracleD(conex)
End If
End Sub

Sub IniciarConexOracleP(conex As ADODB.Connection)
Dim strMi_Usuario As String
Dim strMi_Password As String
Dim strMi_ODBC As String
Dim strconexion As String

 strMi_Usuario = "riesgo"     'nombre de usuario
 strMi_Password = "riesgo"    'password de usuario
 strMi_ODBC = NomServP   ' "DataSource : Fuente de datos"
 strconexion = "Provider=MSDASQL.1;Password=" & _
 strMi_Password & ";User ID=" & strMi_Usuario & ";Data Source=" & _
 strMi_ODBC & ";Persist Security Info=True"
 conex.ConnectionString = strconexion
 conex.Open
End Sub

Sub ReiniciarConexOracleP(conex As ADODB.Connection)
Dim strMi_Usuario As String
Dim strMi_Password As String
Dim strMi_ODBC As String
Dim strconexion As String
On Error Resume Next
 conex.Close
 strMi_Usuario = "riesgo"     'nombre de usuario
 strMi_Password = "riesgo"    'password de usuario
 strMi_ODBC = NomServP   ' "DataSource : Fuente de datos"
 strconexion = "Provider=MSDASQL.1;Password=" & _
 strMi_Password & ";User ID=" & strMi_Usuario & ";Data Source=" & _
 strMi_ODBC & ";Persist Security Info=True"
 conex.ConnectionString = strconexion
 conex.Open
End Sub


Sub IniciarConexOracleD(conex)
Dim strMi_Usuario As String
Dim strMi_Password As String
Dim strMi_ODBC As String
Dim strconexion As String

  strMi_Usuario = "riesvarm"     'nombre de usuario
  strMi_Password = "r13sV4rM2019"    'password de usuario
  strMi_ODBC = NomServQA   ' "DataSource : Fuente de datos"
  strconexion = "Provider=MSDASQL.1;Password=" & _
  strMi_Password & ";User ID=" & strMi_Usuario & ";Data Source=" & _
  strMi_ODBC & ";Persist Security Info=True"
  conex.ConnectionString = strconexion
  conex.Open
End Sub
  

Sub InConexOracle(txtodbc, objeto1)
Dim strMi_Usuario As String
Dim strMi_Password As String
Dim strMi_ODBC As String
Dim strconexion As String

  strMi_Usuario = "riesgo"     'nombre de usuario
  strMi_Password = "riesgo"    'password de usuario
  strMi_ODBC = txtodbc         'DataSource : Fuente de datos"
  strconexion = "Provider=MSDASQL.1;Password=" & _
  strMi_Password & ";User ID=" & strMi_Usuario & ";Data Source=" & _
  strMi_ODBC & ";Persist Security Info=True"
  objeto1.ConnectionString = strconexion
  objeto1.Open
End Sub

Sub IniciarConexQ(conex)
Dim strMi_Usuario As String
Dim strMi_Password As String
Dim strMi_ODBC As String
Dim strconexion As String

    strMi_Usuario = "sivarmer"     'nombre de usuario
    strMi_Password = "sivarmer3"    'password de usuario
    strMi_ODBC = "qsivarme"         'DataSource : Fuente de datos"
    strconexion = "Provider=MSDASQL.1;Password=" & _
    strMi_Password & ";User ID=" & strMi_Usuario & ";Data Source=" & _
    strMi_ODBC & ";Persist Security Info=True"
    conex.ConnectionString = strconexion
    conex.Open
End Sub


Sub FinalizaConexOracle(objeto1)
objeto1.Close
End Sub
