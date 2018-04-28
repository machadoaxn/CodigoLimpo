VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsConversao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function ConverteValor(Valor As Currency)

Dim msg     As String

On Error GoTo LblErr

   Select Case Valor
      Case 1: msg = "Um"
      Case 2: msg = "Dois"
      Case 3: msg = "Três"
      Case 4: msg = "Quatro"
      Case 5: msg = "Cinco"
      Case 6: msg = "Seis"
      Case 7: msg = "Sete"
      Case 8: msg = "Oito"
      Case 9: msg = "Nove"
      Case 10: msg = "Dez"
   End Select
   
   If Valor = 1 Then
      msg = msg & " Real"
   Else
      msg = msg & " Reais"
   End If
   
   MsgBox msg

GoTo LblEnd

LblErr:
   MsgBox Err.Number & ", " & Err.Description & ", " & Err.Source
   Resume LblEnd

LblEnd:
End Function
