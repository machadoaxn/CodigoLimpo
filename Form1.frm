VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1215
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConverter 
      Caption         =   "&Converter"
      Height          =   435
      Left            =   2970
      TabIndex        =   2
      Top             =   660
      Width           =   1725
   End
   Begin VB.TextBox txtValorAConverter 
      Height          =   435
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3165
   End
   Begin VB.Label lbl 
      Caption         =   "Valor a converter:"
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   1425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objConvert          As New ClsConversao

Private Sub cmdConverter_Click()

Dim msg                 As String

On Error GoTo LblErr

   If (txtValorAConverter.Text = "") Or (Not (IsNumeric(txtValorAConverter.Text))) Then
      msg = "Digite um valor v�lido!"
   Else
      msg = objConvert.ConverteValor(txtValorAConverter.Text)
   End If
   
   MsgBox msg, vbInformation + vbSystemModal

GoTo LblEnd

LblErr:
   MsgBox Err.Number, Err.Description, Err.Source
   Resume LblEnd

LblEnd:
End Sub
