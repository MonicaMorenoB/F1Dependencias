VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "propPosFwd"
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


Public FRiesgo1Fwd As String
Public TInterpol1Fwd As Integer
Public FRiesgo2Fwd As String
Public TInterpol2Fwd As Integer
Public FRiesgo3Fwd As String
Public TInterpol3Fwd As Integer
Public TCambioFwd As String
Public fValuacion As String
Public MontoNocFwd As Double
Public PAsignadoFwd As Double
Public ClaveProdFwd As String
Public FCompraFwd As Date
Public FVencFwd As Date
Public EstructuralFwd As String
Public ReclasificaFwd As String
Public ID_ContrapFwd As Long

