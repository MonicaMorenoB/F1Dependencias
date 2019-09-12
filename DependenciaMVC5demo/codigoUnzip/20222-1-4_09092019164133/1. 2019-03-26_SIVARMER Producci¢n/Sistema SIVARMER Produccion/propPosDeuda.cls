VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "propPosDeuda"
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
Public Intencion As String

Public FinicioDeuda As Date
Public FVencDeuda As Date
Public InteriDeuda As String
Public InterfDeuda As String
Public RintDeuda As String
Public TRefDeuda As String
Public SpreadDeuda As Double
Public ConvIntDeuda As String
Public ProductoDeuda As String
Public SiposPrimDeuda As String
Public PCuponDeuda As Integer
Public FRiesgo1Deuda As String
Public TInterpol1Deuda As Integer
Public FRiesgo2Deuda As String
Public TInterpol2Deuda As Integer
Public TCambioDeuda As String
Public fValuacion As String
Public IFlujoDeuda As Long
Public FFlujoDeuda As Long

