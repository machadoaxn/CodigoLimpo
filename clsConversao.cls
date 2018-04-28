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

Dim aUnidade(20) As String
Dim aDezenas(9) As String


Public Function ConverteValor(ValorAConverter As Currency) As String

Dim Texto         As String
Dim Valor         As Currency

On Error GoTo LblErr

   Call PreencherValorUmANove
   
   If CInt(ValorAConverter) Then
      If ValorAConverter >= 0 And ValorAConverter <= 19 Then
         ConverteValor = aUnidade(ValorAConverter)
      End If
   End If

GoTo LblEnd

LblErr:
   MsgBox Err.Number & ", " & Err.Description & ", " & Err.Source
   Resume LblEnd

LblEnd:
End Function

Public Function PreencherValorUmANove() As String
On Error GoTo LblErr
   
   aUnidade(0) = "zero"
   aUnidade(1) = "um"
   aUnidade(2) = "dois"
   aUnidade(3) = "três"
   aUnidade(4) = "quatro"
   aUnidade(5) = "cinco"
   aUnidade(6) = "seis"
   aUnidade(7) = "sete"
   aUnidade(8) = "oito"
   aUnidade(9) = "nove"
   aUnidade(10) = "dez"
   aUnidade(11) = "onze"
   aUnidade(12) = "doze"
   aUnidade(13) = "treze"
   aUnidade(14) = "quatorze"
   aUnidade(15) = "quinze"
   aUnidade(16) = "dezesseis"
   aUnidade(17) = "dezessete"
   aUnidade(18) = "dezoito"
   aUnidade(19) = "dezenove"
   
GoTo LblEnd

LblErr:
   MsgBox Err.Number & ", " & Err.Description & ", " & Err.Source
   Resume LblEnd

LblEnd:
End Function

Public Function PreencherValorDezADezenove() As String
On Error GoTo LblErr
   
   aDezenas(10) = "dez"
   aDezenas(11) = "onze"
   aDezenas(12) = "doze"
   aDezenas(13) = "treze"
   aDezenas(14) = "quatorze"
   aDezenas(15) = "quinze"
   aDezenas(16) = "dezesseis"
   aDezenas(17) = "dezessete"
   aDezenas(18) = "dezoito"
   aDezenas(19) = "dezenove"
   
GoTo LblEnd

LblErr:
   MsgBox Err.Number & ", " & Err.Description & ", " & Err.Source
   Resume LblEnd

LblEnd:
End Function




