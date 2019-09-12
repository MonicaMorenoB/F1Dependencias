VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "resValFlujo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public c_operacion As String
Public t_pata As String
Public fecha_ini As Date
Public fecha_fin As Date
Public fecha_desc As Date
Public si_paga_int As String
Public int_s_saldo As String
Public saldo_periodo As Double
Public amortizacion As Double
Public t_cupon_per As Double
Public sobretasa As Double
Public dxv1 As Long
Public dxv2 As Long
Public dxv3 As Long
Public p_cupon As Integer
Public f_tiempo_aplicar As Double
Public tc_aplicar As Double
Public saldo_efectivo_aplicar As Double
Public int_gen_periodo As Double
Public int_acum_periodo As Double
Public int_pag_periodo As Double
Public int_acum_sig_periodo As Double
Public pago_total As Double
Public t_desc As Double
Public factor_desc As Double
Public valor_presente As Double



