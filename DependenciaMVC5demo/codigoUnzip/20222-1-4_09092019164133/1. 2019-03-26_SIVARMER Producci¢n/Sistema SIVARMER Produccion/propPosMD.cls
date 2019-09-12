VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "propPosMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public tipopos As Integer
Public fechareg As Date
Public NomPos As String
Public HoraRegOp As String
Public C_Posicion As Integer
Public c_operacion As String
Public Tipo_Mov As Integer
Public Signo_Op As Integer
Public intencion As String

Public tValorMD As String                     'tipo valor de la emision
Public emisionMD As String                    'emision
Public serieMD As String                      'serie
Public cEmisionMD As String                   'clave de la emision para instrumentos de la md
Public noTitulosMD As Double                  'no de titulos/monto nocional
Public fCompraMD As Date                      'fecha de compra
Public fVencMD As Date                        'fecha de vencimiento de la operacion
Public pAsignadoMD As Double                  'precio de compra/asignado
Public vNominalMD As Double                   'valor nominal
Public tReporto As Double                     'tasa reporto
Public subport1MD As String                   'subportafolio 1 de la posicion
Public pCuponMD As Integer                    'periodo cupon (de la tabla de valuacion)
Public sTCuponMD As Double                    'sobretasa cupon (de la tabla de valuacion
Public SiFlujosMD As String                   'utiliza flujos para valuar (de la tabla de valuacion)
Public escala As String
Public sector As String
Public recupera As Double

Public iFlujoMD As Long                       'inicio de los flujos en la tabla
Public fFlujoMD As Long                       'fin de los flujos en la tabla

'datos que se obtienen del vector de precios
Public pCuponVigenteMD As Integer             'periodo cupon vigente
Public tCuponVigenteMD As Double
Public intDevengMD As Double                  'intereses devengados para los instrumentos tipo BondesD
Public valSucioPIP As Double
Public valLimpioPIP As Double
Public Calif1MD As String
Public Calif2MD As String
Public Calif3MD As String
Public Calif4MD As String
Public CalifTMD As String
Public CalifFP As String
Public CalifLP As Integer
Public reglaCuponMD As String

'factores de riesgo para valuar
Public fRiesgo1MD As String
Public tInterpol1MD As Integer
Public fRiesgo2MD As String
Public tInterpol2MD As Integer
Public fRiesgo3MD As String
Public tInterpol3MD As Integer
Public fRiesgo4MD As String
Public tInterpol4MD As Integer
Public tCambioMD As String
Public fValuacion As String

'datos que se calculan a posterior
Public dVCuponMD As Integer               'dias para vencer cupon vigente
Public dVencMD As Integer
Public tCuponMD As Double

