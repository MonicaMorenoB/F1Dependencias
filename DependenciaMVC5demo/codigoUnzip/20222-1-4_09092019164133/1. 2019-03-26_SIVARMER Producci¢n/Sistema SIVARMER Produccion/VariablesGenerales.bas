Attribute VB_Name = "VariablesGenerales"
Option Explicit

Global MREFProsSwap() As Variant
Global PrefijoBD As String
Global Mat_MesaDinero() As Variant
Global MensajeProc As String
Global AvanceProc As Double
Global NoRutinasS As Integer
Global NoFechasPIP As Long
Global NoPortafolios As Integer

'Global MatFechasPIP() As Date
Global MatFRPos() As Variant
Global Altura As Double
Global alfa As Double

Global CIFlujoActSwap As Long
Global CFFlujoActSwap As Long
Global CIFlujoPasSwap As Long      'inicio flujo pasiva
Global CFFlujoPasSwap As Long      'fin flujo pasiva
Global FechaFinal As Date
Global ComprasDirecto As Double
Global ComprasReporto As Double
Global CoefAsimetria As Double          'coeficiente de asimetria del instrumento
Global Curtosis As Double               'curtosis
Global Dias As Double                   'el mismo valor e NoDias, pero en una variable de precision doble
Global DirReportes As String
Global DirVPrecios As String
Global DirPosMesaD As String
Global DirPosPensiones As String
Global DirPosPensiones2 As String
Global Mat_Sentencia() As Variant
Global Tabla_Excel() As Variant
Global DirSistemVAR As String
Global DirCurvasZ As String
Global DirCurvas As String
Global DirBases As String
Global DirRTexto As String
Global MatDeltas() As Double
Global FechaEval As Date
Global FechasPosRiesgo() As Variant  'matriz con fechas de la posicion Riesgo
Global NoFechasPosRiesgo As Long
Global FEscala As Long
Global FechaPos As Date         'fecha del calculo de la posicion
Global FechaFactor As Date
Global filtro As String        'variable para filtrar un archivo
Global grafico As Object
Global IndiceInstrumento As Long 'instrumento a analizar
Global LimiteTenenciaNeta As Double
Global NombreFRiesgo() As String         'esta es una matriz permanente
Global matfechas() As Variant
Global MatFechasVec1() As Date
Global MatFechasPos() As Date
Global MatPortPosicion() As Variant
Global MatFechasVPrecios() As Date
Global MatPortCurvas() As Variant
Global MatGruposPortPos() As Variant
Global MatEstRepVaRB() As Variant
Global MatPortDeriv() As Variant
Global SiConsolidaPos As Boolean
Global ejecutaSubproc1 As Boolean
Global ejecutaSubproc2 As Boolean
Global ejecutaSubproc3 As Boolean

Global MatGrupo() As String
Global MatLimites() As Double
Global MatOperacion() As Variant

Global MatPlazos() As Long
Global MatDescripFR() As Variant
Global MatPlazosTot() As Long
Global MatPrecios() As New resValIns
Global MatFactR1() As Double
Global FechaMatFactR1 As Date
Global matrendsSH() As Double
Global matBndSH() As Integer
Global noescSH As Long
Global htiempoSH As Long
Global fechaSH As Date
Global MatTasasP() As Double
Global MatIncP() As Double
Global MatVectores() As String
Global matrizch() As Double
Global MatTiempos() As Double      'matriz para evaluar el desempeño de algunas rutinas
Global MatCapitalSist() As Variant
Global MatPosSim() As Variant
Global MatPortEstruct() As String


Global NDiasBack  As Long
Global NoCapitalBase As Long
Global NoDatVecMer As Long
Global nofriesgo As Long    ' No de factores de riesgo
Global NoFechas As Long
Global NoFechasPos As Long
Global NoFechasPos1 As Long
Global NoFechasVPrecios As Long
Global MatResFRiesgo() As New resPropFRiesgo
Global MatResFRiesgo1() As Variant
Global MatNodosCurvas() As Variant
Global NoPlazos As Integer
Global NoPlazosTot As Integer

Global MatPLPos() As Variant
Global noprecios As Long
Global NoInstDirecto As Long      'no de instrumentos que opera en directo
Global NoInstReporto As Long      'No de Instrumentos que operan en reporto
Global NoInstrumentosValuados As Long
Global NoOperaciones As Long
Global NoCompDirecto As Long
Global NoCompReporto As Long
Global NoVenReporto As Long
Global MatVal0Con() As Variant
Global MatPyGCon() As Variant
Global MatPyGTot() As Double
Global MatVal0T() As resValIns
Global MatPyGT() As Double
Global MatResValFlujo() As resValFlujo

Global nombre As String
Global NoPosMesa As Long
Global NoSimulaciones As Long
Global NoVectores As Long
Global NoAciertos As Long
Global Pi As Double

'variables para el sistema estadistico

Global matdatos() As Variant            'matriz con el total de datos de la tabla de datos

Global NomSistem As String              'Nombre por el que el sistema Responde
Global MatHistograma() As Variant       'matriz con los calculos del histograma
Global MatVolatil() As Variant          '
Global nodias As Integer                'no de dias que se van a necesitar para el análisis

Global NoIntervalos As Integer          'no se intervalos en el histograma
Global DiasVolatil As Integer           '
Global DiasEfectivo As Integer          '
Global ValorMaximo As Double
Global ValorMinimo As Double
Global Kolmogorov As Double

Global IndicePos As Long                'indice que apunta a un instrumento de la posicion
Global fmaximo As Double
Global media As Double                  'media del los rendimientos
Global desvest As Double                'desviacion estandar de los rendimientos
Global LambdaMark As Double
Global LambdaMont As Double
Global NConfMark As Double              'nivel de confianza para el calculo de VAR
Global NConfMont As Double
Global NConfHist As Double

Global TVolMark As Integer
Global TVolMont As Integer
Global TVolHist As Integer
Global HorizMark As Integer
Global HorizMont As Integer
Global HorizHist As Integer

Global Procede As String
Global Procede2 As String
Global UltimoDia As Date
Global Fluctuacion As Double
Global ffinal As Date
Global IFecha As Integer                'indice de la tabla de datos donde esta la fecha inicial
Global Vigencia As Integer
Global SiCargaDatos As Boolean
Global SiCargoSerie As Boolean
Global SiCargaFechasPosicion As Boolean
Global SiCargaFechasPrecios As Boolean
Global SiCargaTasas As Boolean
Global SiHab As Boolean
Global finicio As Date
Global ffin As Date
Global matorden() As Variant
Global contador As Long
Global SiPosicionCargada As Boolean
Global CCPortafolio As Integer
Global CCOperacion As Integer
Global CTablaFlujosMD As Integer

'variables que hacen referencias a objetos
Global Cuadro As Object

Global ValorTotalSimulacion As Double
Global ValorMercado As Double
Global ValorCurva As Double
Global ValorVAR As Double
Global VectorPrecios() As Variant
Global MatFactRiesgo() As Variant   'los factores de riesgo ajustados por split
Global MatFactoresRiesgo2() As Variant   'los factores de riesgo ajustados por split
Global SiEncontroFechaPos As Boolean

'VARIABLES MODULO SIMULACION
Global MatSimulacion() As Variant    'matriz donde van los precios simulados
Global MatTasasSim() As Variant      'Matriz donde van las tasas simuladas
Global NoDiasSim As Long
Global ValorMedia As Double
Global ValorDS As Double
Global VecTasasSimulacion() As Double             ' en este vector se guardan las tasas para la generacion de los

'variables para valuacion
Global DVencimiento As Long
Global DVen0 As Long
Global PMercado As Double
Global PVPrecios As Double
Global TInteres As Double
Global TCuponVigente As Double
Global tinterpol As Double
Global st As Double
Global ValorNominal As Double
Global VMercado As Double
Global VCurva As Double
Global MatPosDesglose() As Variant
Global MatDesgloseDivisa() As Variant
Global NoInstrumentos() As Long
Global MatResPosicion() As Variant
Global MatVARMarkowitz() As Variant
Global VARDurMD As Double
Global LimInfVMark As Double
Global LimSupVMark As Double
Global LimInfVMont As Double
Global LimSupVMont As Double
Global LimInfVHist As Variant
Global LimSupVHist As Double
Global MatVARMontecarlo() As Double
Global VARExtMD As Double
Global FechasVPrecios() As Date
Global MatValExt() As Variant
Global MatVARExtremo() As Double
Global MatTotalesExtremo() As Double
Global MatTotalMarkowitz() As Double
Global MatTotalDuracion() As Double
Global ContVPrecios() As Long
Global NoPosRiesgo As Long
Global MatAciertos() As Long
Global MatTasasRef() As Variant        'vector con las tasas de referencia
Global CapitalNeto As Double
Global CapitalBase As Double

Global TipoRendimiento As String
Global MatUsuarios() As Variant
Global LoginSucceeded As Boolean
Global ContraseñaCatalogos As String

Global LimiteExtremo As Double
Global SiCorrerVARMarkowitz As Boolean
Global SiCorrerVARMontecarlo As Boolean
Global SiCorrerVARExtremo As Boolean
Global NoGruposPort As Long  'son los distintos tipos de instrumentos que maneja el sistema
'se definen variables que indican en que columna de la
'matriz se encuentra un factor de riesgo

Global SiFactorRiesgo() As Boolean
Global NDReutersImportados As Long
Global NoPosicionImportados As Long
Global NoPreciosImportados As Long
Global MatPosPension() As Variant
Global NoPosPension As Long
Global DirVPreciosZ As String
Global DirFlujosEm As String
Global DirFlujosEmZ As String
Global DirVAnalitico As String
Global DirVAnaliticoZ As String
Global CDVencimiento As Integer
Global CDuracAct As Integer
Global CDuracPas As Integer
Global CDV01Act As Integer
Global CDV01Pas As Integer
Global CPCuponActSwap As Integer
Global CPCuponPasSwap As Integer
Global CTipoMov As Integer


Global CClaveProd As Integer

Global CFactor1 As Integer
Global CFactor2 As Integer
Global CTReporto As Integer
Global CEmisor As Integer
Global CMoneda As Integer
Global CCPosicion As Integer
Global CDescTitulo As Integer
Global CFechaReg As Integer
Global CSTasa As Integer

Global CSubportafolio1 As Integer



Global matvalIKOS() As Variant
Global fechavalIKOS As Date


Global CFVolatil As Long
Global CSobreT As Long
Global CFValuacion As Long
Global CSector As String

Global MatPosRiesgo() As propPosRiesgo
Global MatFlujosSwaps() As estFlujosDeuda
Global MatFlujosDeuda() As estFlujosDeuda
Global MatFlujosMD() As estFlujosMD

Global ErrorVarianza As Double
Global ValorPosProv As Double
Global ValorPosicion As Double
Global ValorPosPasiva As Double
Global ValorPosActiva As Double
Global TTitCompra As Double
Global TTitVenta As Double

Global NoFactores As Long
Global MatPCTR() As Long
Global MatCaracFRiesgo() As propNodosFRiesgo
Global MatNomFactor1() As Variant
Global MatPreciosSimulados() As Double
Global MatSensib1() As Variant
Global MatSensib2() As Variant
Global MatSensib3() As Variant
Global MatSensib4() As Variant
Global MatSensib5() As Variant
Global MatSensib6() As Variant
Global MatSensib7() As Variant
Global MatSensib8() As Variant

Global MatMediaM1() As Double
Global MatMediaM2() As Double
Global NoGruposFR As Long
Global NoCurvasTot As Long

Global MatSensNum() As Variant
Global MatContrapartes() As Variant
Global MatContrapartes1() As Variant
Global MatTresholdContrap() As Variant
Global MatRelSwapEm() As Variant
Global MatClavesContrap() As Variant
Global MatDerivSinLMargen() As Variant

Global tipotitulos As String

Global ConAdo As New ADODB.Connection
Global conAdo1 As New ADODB.Connection
Global conAdo2 As New ADODB.Connection
Global conAdo3 As New ADODB.Connection
Global conAdoBD As New ADODB.Connection


Global RFlujos As New ADODB.recordset
Global REstadoUsuario As New ADODB.recordset
Global RnMesa As New ADODB.recordset
Global RGuardarPL As New ADODB.recordset
Global RGuardarPLMont As New ADODB.recordset
Global RegResCVA As New ADODB.recordset
Global RegResMakeW As New ADODB.recordset
Global RegResLimC1 As New ADODB.recordset
Global RegResLimC2 As New ADODB.recordset
Global RegResCVAMD As New ADODB.recordset

Global RnPrecios As New ADODB.recordset
Global RegExcel As New ADODB.recordset

Global CDescContra As Integer
Global CContraSistem As Integer
Global SiVerificoPos As Boolean
Global SiVerificoTasas As Boolean

Global SiExisteVecPrecios As Boolean
Global SiExisteVecPreciosMD As Boolean
Global SiExisteCurvas As Boolean
Global MatCholeski() As Double
Global MatCovar1() As Double
Global MatCovar2() As Double
Global MatIndVPrecios() As Variant
Global MatPLPosC() As Double
Global SiGuardoDBMexico As Boolean
Global MatFechasVaR() As Date
Global MatFechasFR() As Date
Global MatFechasTareas1() As Date
Global MatFechasTareas2() As Date
Global MatHistCurvas() As Variant
Global MatCatCurvas() As Variant
Global MatFechasEstres() As Variant
Global SiDActMDinero As Boolean
Global SiDActMCambios As Boolean
Global SiDActMDerivados As Boolean
Global SiDActTPIP As Boolean

Global SiDActVPrecios As Boolean
Global DirSalida As String
Global DirPosDiv As String
Global DirPosSwaps As String
Global DirPosPrimSwaps As String
Global DirPosPrimFwd As String
Global MatClavesFactIKOS()  As Variant
Global MatListaPortPos() As Variant
Global DirResVaR As String
Global NoEscMark As Long
Global NoEscMont As Long
Global NoEscHist As Long
Global NoSimMont As Long

Global PerfilUsuario As String
Global Id_Sesion As String
Global NomUsuario As String
Global idusuario As Integer
Global DirFactReuters As String
Global DirArchBat As String
Global MatDesBack() As Variant
Global MatResumenBack() As Variant
Global MatRendimientos() As Double
Global NoFactRPIP As Long
Global NoRegPIP As Long
Global ICTIMP As Integer
Global IUMS As Integer
Global ITPFB As Integer
Global MatCurvaPIP() As Variant
Global MatFactorRPIP() As String
Global ActivarControlErrores As Boolean
Global CSigno As Integer
Global CVFRiesgo1 As Integer
Global CVFRiesgo2 As Integer
Global CVFRiesgo3 As Integer
Global CVFRiesgo4 As Integer
Global CTInterpol1MD As Integer
Global CTInterpol2MD As Integer
Global CTInterpol3MD As Integer
Global CTInterpol4MD As Integer
Global CTInterCST As Integer
Global ValExacta As Boolean
Global MatNodosFREx() As New propNodosFRiesgo
Global DefCurvasSH() As Variant
Global MatPlazosSH() As Variant
Global CurvaSH() As Variant
Global MatDetCurvaSH() As Variant
Global DirWinRAR As String
Global DirCurvasCSV1 As String
Global DirCurvasCSV2 As String
Global DirTemp As String
Global DirIndFecha As String
Global DirCurvaReal As String
Global DirHistVaRMD As String
Global DirHistREjecutivo As String
Global SiValFR As Boolean
Global NoFRExactos As Long
Global matfechassh() As Variant
Global TPercMont As Variant
Global TPercHist As Variant
Global SiPregGuarda As Boolean
Global MatEscenariosN As Variant
Global MatEscenariosT As Variant
Global MatTPFB() As Variant
Global MatEscenariosMens() As Variant

Global DirPosFwdTC As String
Global DirPosFwdTasa As String
Global MatSimHist() As Variant
Global MatSimMont() As Variant
Global MatClavesEmision() As Variant
Global MatCatProcesos() As Variant
Global MatResEficSwaps() As Variant
Global MatResEficFwd() As Variant
Global MatNumaMont() As String
Global matordenMont() As Variant
Global mmediasMont() As Double
Global mTransicionN() As Double
Global mTransicionI() As Double

Global FechaArchCurvas As Date
Global fechaMatTrans As Date
Global MatHora() As Date
Global SiHacer As Boolean
Global NomArchRVaR As String
Global CDVCuponAct As Integer
Global CDVCuponPas As Integer
Global ICDVCDirecto As Integer
Global ICDVCReporto As Integer
Global ICDVVReporto As Integer
Global ICDVVDirecto As Integer
Global MatEficFwds() As Variant
Global CIntencion As Integer
'bases de datos
Global TablaPosMD As String
Global TablaPosDiv As String
Global TablaPosSwaps As String
Global TablaPosSwapsIDO As String

Global TablaPosDeuda As String
Global TablaFlujosDeudaO  As String
Global TablaFlujosSwapsIDO As String
Global TablaPosFwd As String
Global TablaPosFwdFrutos As String
Global TablaFlujosSwapsO As String
Global TablaFlujosSwapsSimO As String
Global TablaEficienciaO As String
Global TablaPosicionO As String
Global TablaVecPrecios As String
Global TablaSensibN  As String
Global TablaDetalleMo As String
Global TablaAnalisisFRO As String
Global TablaAnalisisFRA As String
Global TablaResultadosA As String
Global TablaFRiesgoA As String
Global TablaFRiesgoO As String
Global TablaFRiesgoNA As String
Global TablaCatalogosA As String
Global TablaValExtO As String
Global TablaValExtA As String
Global TablaResPosA As String
Global TablaResVaR As String
Global TablaBackPort As String
Global TablaResBack As String
Global TablaValPos As String
Global TablaValPosPort As String
Global TablaPLHistOper As String
Global TablaPLEscCVA As String
Global TablaPLEscMakeW As String
Global TablaPLEscMW As String
Global TablaPLEscHistPort As String
Global TablaBacktA As String
Global TablaReportesA As String
Global TablaSQLPort As String
Global TablaGruposPapelFP As String
Global TablaSensibPort As String
Global TablaPLHistOperVR As String
Global TablaPyGHistPVR As String
Global TablaResVReemplazo As String
Global TablaResCalcVReemplazo As String
Global TablaFactChol As String
Global TablaResEfectPros As String
Global TablaResMO As String
Global TablaPosSelecc As String


Global CConsecutivo As Long
Global OpcionBDatos As Integer
Global BDIKOS As Integer
Global SiCargoFactR As Boolean
Global PrecioLimpio As Integer
Global NombrePortFR As String
Global MatCurvasT() As Variant
Global MatCurvasT1() As Variant
Global MatCurvasT2() As Variant
Global MatFactRiesgo1() As Double
Global MatFactRiesgo2() As Double
Global fechaFactR1 As Date
Global fechaFactR2 As Date

Global txtCadenaMensajes() As Variant
Global CSTCuponMD As Variant
Global CDecimalT As Integer
Global MatRelSwapsPrim() As propRelSwapPrim
Global MatDerivEst() As Variant
Global MatParamFwds() As Variant
Global MatValFlujosD() As resValFlujoExt
Global TablaCatProcesos As String
Global MatVPreciosT() As Variant
Global FechaVPrecios As Date
Global FechaPosRiesgo As Date
Global txtNomPosRiesgo As String
Global TablaGruposPortPos As String
Global TablaReporteCVaR As String
Global TablaHistCurvas As String
Global TablaProcesos1 As String
Global TablaProcesos2 As String
Global TablaSubProcesos1 As String
Global TablaSubProcesos2 As String
Global TablaSubProcesos3 As String
Global TablaBitTareasA As String
Global TablaParamSistema As String
Global TablaParamUsuario As String
Global TablaRecNacional As String
Global TablaRecInt As String
Global TablaFechasVaR As String
Global TablaFechasFR As String
Global TablaBlackList As String
Global TablaOperValidada As String
Global TablaEquivContrap As String
Global TablaContrapartes As String
Global TablaFechasTareas1 As String

Global MatEscHistD() As Variant
Global CValActivaIKOS As Integer
Global CValPasivaIKOS As Integer
Global CMtmIKOS As Integer
Global SRutOrden As Integer
Global SiAgregarDatosFwd As Boolean
Global SiAnexarFlujosSwaps As Boolean
Global SiIncTasaCVig As Boolean
Global SiPassNuevo As Boolean
Global txtUsuarioCC As String
Global TablaUsuarios As String
Global TablaSesiones As String
Global MatPortafolios() As Variant
Global TablaInterfCarac As String
Global TablaInterfFlujos As String
Global TablaInterfFwd As String
Global TablaInterfDiv As String
Global TablaInterfSim1 As String
Global TablaInterfSim2 As String
Global TablaPosMesaIKOS As String

Global VersExcel As String
Global NomServQA As String
Global NomServP As String
Global NomSRVPIP As String
Global usersftpPIP As String
Global passsftpPIP As String

Global TablaLimites As String
Global TablaCEmision As String
Global TablaPortFR As String
Global TablaBitacora As String
Global TablaBitacoraIF As String
Global TablaValDeriv As String
Global TablaPEmision As String
Global TablaFlujosRichard As String
Global TablaPYGCVAMD As String
Global TablaPyGMontOper As String
Global TablaPyGMontPort As String
Global MatParamEmisiones() As Variant
Global MatParamSistema() As Variant
Global MatParamUsuario() As Variant
Global NoProcesosT As Integer
Global ICIDevActiva As Integer
Global SiCargaFEmMD As Boolean
Global CSiPosPrim As Integer
Global CIndPosicion As Integer
Global CTabla As Integer
Global CTipoPos As Integer
Global CNomPos As Integer
Global MatDerEstandar() As Variant
Global MatRelSwapIS() As Variant
Global MatCalifEmision() As Variant
Global MatCalendSwaps() As Variant
Global MatBaseCalend() As Variant
Global MatRelSwapsCal() As Variant
Global TablaCalendSwapsO As Variant
Global TablaBaseCalendO As Variant

Global TablaFlujosMDA As String
Global MatSecProcesos() As Variant
Global MatSecSubproc() As Variant

Global ValEficiencia As Variant
Global RespetarSecProc As String
Global TablaEficSwaps As String
Global TablaCatPortPos As String
Global TablaPortPosEstructural As String
Global TablaIndVecPreciosO As String
Global TablaMOEmSectorMD As String
Global TablaMOEmSectorPI As String
Global TablaMOContrapCalifPI As String
Global TablaMOContrap As String
Global TablaMOperCalif As String
Global TablaMOEmPriv As String
Global TablaMOMon As String
Global TablaMOGub As String
Global TablaEficienciaCob As String
Global CHoraRegOp As Integer


Global MatTValBC0() As Variant
Global MatTValBonos() As Variant
Global MatTValReportos() As Variant
Global MatTValSTCupon() As Variant
Global MatTValSTDesc() As Variant
Global MatTValSwaps1() As Variant
Global MatTValSwaps2() As Variant
Global MatTValDeuda() As Variant
Global MatTValFwdsTC1() As Variant
Global MatTValFwdsTC2() As Variant
Global MatReporteCVaR() As Variant
Global MatPortSegRiesgo() As String
Global MatTValInd() As Variant
Global MBList() As Variant
Global ClavePosMD As Integer
Global ClavePosTeso As Integer
Global ClavePosMC As Integer
Global ClavePosDeriv As Integer
Global ClavePosPension1 As Integer
Global ClavePosPension2 As Integer
Global ClavePosDeuda As Integer
Global ClavePosPIDV  As Integer
Global ClavePosPICV As Integer
Global ClavePosPID As Integer
Global ClavePosPenMD As Integer

Global ClaveCDirec As Integer
Global ClaveVDirec As Integer
Global MatMonedas() As Variant
Global BlockSize  As Long
Global MatFRSplit() As Variant
Global MatEmxContrap() As Variant
Global MatGruposDeriv() As Variant
Global MatCalificaciones() As Variant
Global MatTransI() As Double
Global MatTransN() As Double
Global MatRecuperaI() As Variant
Global MatRecuperaN() As Variant
Global TablaVaRIKOS As String
Global TablaDerEstandar As String
Global TablaDerivEst2 As String
Global TablaPortPrincipales As String
Global TablaSecProcesos As String
Global TablaSecSubProc As String
Global TablaTreshCont As String
Global TablaCalifContrapF As String
Global TablaCalifContrapNF As String
Global TablaCalifContrapEmision As String

Global TablaBLTresh As String
Global TablaEficRetro As String
Global TablaValBonosC0 As String
Global TablaFechasEscEstres As String
Global TablaValSwaps1 As String
Global TablaValSwaps2 As String
Global TablaValDeuda As String
Global TablaMTrans As String
Global TablaGruposDeriv As String
Global TablaCalificaciones As String
Global TablaEmxContrap As String
Global TablaSectorEscEm As String
Global TablaEscCortoLargo As String
Global TablaValBonos As String
Global TablaValReportos As String
Global TablaSplits As String
Global TablaMonedas As String
Global TablaValFwds1 As String
Global TablaValFwds2 As String
Global TablaValInds As String
Global TablaValBSC As String
Global TablaValBSD As String
Global TablaNodosCurvas As String
Global TablaPosPrimarias As String
Global TablaRelSwapEm As String
Global TablaCalendSwaps As String
Global TablaFlujosMD As String
Global TablaPosDivCon As String
Global TablaNumDistNormal As String
Global TablaVDerIKOS As String
Global TablaValContraparte As String
Global TablaResCVA As String
Global TablaVARInter As String
Global TablaCurvas As String
Global TablaCatCurvas As String
Global TablaResEscEstres As String
Global TablaResEscEstresPort As String
Global TablaResEstresAprox As String
Global NoIntFallidos() As Integer
Global SiDepuracion As Boolean
Global IdPosPension As Integer
Global IdPosPension2 As Integer
Global SiActTProc As Boolean

Global txtCadCarEsp As String
Global txtCadNum As String
Global txtCadMin As String
Global txtCadMay As String

'Global matposmd() As propPosMD
'Global matposdiv() As propPosDiv
'Global matposswaps() As propPosSwaps
'Global MatPosFwds() As propPosFwd
'Global matposdeuda() As propPosDeuda
Global MatRelSwapsDeuda() As Variant
Global TablaRSwapsDeuda As String
Global TablaRelSwapIKOSS As String
Global TablaBlEsc As String
Global MatBlEsc() As Variant
Global MatSQLPort() As Variant
Global MatGruposPapelFP() As Variant
Global MatMOSectorMD() As Variant
Global MatMOSectorPI() As Variant
Global MatMOGub() As Variant
Global MatMOCalif() As Variant
Global MatMOEmPriv() As Variant
Global MatMOMon() As Variant
Global MatMOContrapCalifPI() As Variant
Global MatMOContrap() As Variant
Global TablaPortPosicion As String
Global TablaResWRW As String
Global TablaEscFR As String
Global TablaLimContrap1 As String
Global TablaLimContrap2 As String
Global TablaResLimContrap As String
Global TablaExpFwds As String
Global TablaMovBalanza As String
Global SiCorrerSubp As Boolean
Global txtportBanobras As String
Global txtportCalc1 As String
Global txtportCalc2 As String
Global fechaFactoresR As Date
Global matEscEstres() As Double
Global matNomEscEstres() As String
Global matResEscEstres() As Variant
Global FechaEscEstres As Date
Global IndOperR As Integer
