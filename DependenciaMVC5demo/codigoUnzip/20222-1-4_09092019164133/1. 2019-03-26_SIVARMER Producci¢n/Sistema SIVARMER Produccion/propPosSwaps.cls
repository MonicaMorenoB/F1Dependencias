VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "propPosSwaps"
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

Public IFlujoActSwap As Long            'inicio flujo activa
Public FFlujoActSwap As Long            'fin flujo activa
Public IFlujoPasSwap As Long            'inicio flujo pasiva
Public FFlujoPasSwap As Long            'fin flujo pasiva

Public FCompraSwap As Date
Public FvencSwap As Date
Public IntercIFSwap As String           'intercambio incial de intereses
Public IntercFFSwap As String           'intercambio final de intereses

Public RIntAct As String
Public RIntPas As String
        
Public TCActivaSwap As String
Public TCPasivaSwap As String
Public STActiva As Double
Public STPasiva As Double

Public ConvIntAct As String
Public ConvIntPas As String
Public DVCuponAct As Integer
Public DVCuponPas As Integer
Public ClaveProdSwap As String
Public cProdSwapGen As String
Public c_em_pidv As String
Public EstructuralSwap As String

Public PCuponActSwap As Integer              'periodo cupon activa
Public PCuponPasSwap As Integer              'periodo cupon pasiv

Public FRiesgo1Swap As String                'curva para descontar pata activa
Public TInterpol1Swap As Integer
Public FRiesgo2Swap As String                'curva para descontar pata pasiva
Public TInterpol2Swap As Integer
Public FRiesgo3Swap As String                 'curva para calcular cupon activa
Public TInterpol3Swap As Integer
Public FRiesgo4Swap As String                'curva para calcular cupon pasiva
Public TInterpol4Swap As Integer
Public TCambio1Swap As String
Public TCambio2Swap As String
Public fValuacion As String
Public ID_ContrapSwap As Long

