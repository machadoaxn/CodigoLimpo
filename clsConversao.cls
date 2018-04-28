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
Dim aDezenas(9)  As String


Public Function ConverteValor(ValorAConverter As Currency) As String

Dim Texto            As String
Dim ValorInt         As Long
Dim ValorDec         As Long
Dim ControleInt      As Boolean

On Error GoTo LblErr

   Call PreencherValorUmADezenove
   Call PreencherDezenas
   ControleInt = False
   ValorInt = CInt(ValorAConverter)
   ValorDec = Right(CStr((ValorAConverter) - CInt(ValorAConverter)), (Len(CStr((ValorAConverter) - CLng(ValorAConverter))) - 2))
   
   If ValorInt > 0 And ValorInt <= 19 Then
      ConverteValor = aUnidade(ValorInt)
      ControleInt = True
      If ValorInt = 1 Then
         ConverteValor = ConverteValor & " real"
      Else
         ConverteValor = ConverteValor & " reais"
      End If
   End If
   
   If ControleInt And (ValorDec > 0) Then
      If (ValorDec) < 20 Then
         ConverteValor = ConverteValor & " e " & aUnidade(ValorDec) & " centavos"
      Else
         ConverteValor = ConverteValor & " e " & aDezenas(Left(ValorDec, 1)) & " e " & aUnidade(Right(ValorDec, 1)) & " centavos"
      End If
   End If

GoTo LblEnd

LblErr:
   MsgBox Err.Number & ", " & Err.Description & ", " & Err.Source
   Resume LblEnd

LblEnd:
End Function

Public Function PreencherValorUmADezenove() As String
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

Public Function PreencherDezenas() As String
On Error GoTo LblErr
   
   aDezenas(0) = "zero"
   aDezenas(1) = "dez"
   aDezenas(2) = "vinte"
   aDezenas(3) = "trinta"
   aDezenas(4) = "quarenta"
   aDezenas(5) = "cinquenta"
   aDezenas(6) = "sessenta"
   aDezenas(7) = "setenta"
   aDezenas(8) = "oitenta"
   aDezenas(9) = "noventa"
   
GoTo LblEnd

LblErr:
   MsgBox Err.Number & ", " & Err.Description & ", " & Err.Source
   Resume LblEnd

LblEnd:
End Function




