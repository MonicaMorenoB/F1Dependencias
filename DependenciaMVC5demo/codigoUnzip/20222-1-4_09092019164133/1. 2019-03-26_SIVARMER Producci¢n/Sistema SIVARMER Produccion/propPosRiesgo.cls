VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "propPosRiesgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public IndPosicion As Integer
Public No_tabla As Integer
Public tipopos As Integer
Public fechaReg As Date
Public NomPos As String
Public HoraRegOp As String
Public C_Operacion As String
Public C_Posicion As Integer
Public Tipo_Mov As Integer
Public Signo_Op As Integer
Public Intencion As String
Public FValuacion As String
Public Sector As String
Public ValActivaIKOS As Double
Public ValPasivaIKOS As Double
Public MtmIKOS As Double

