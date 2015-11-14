VERSION 5.00
Begin VB.Form frmSugestao 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio de Sugestões"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "frmSugestao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7230
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sugestões"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdFechar 
         Caption         =   "&Fechar"
         Height          =   375
         Left            =   4800
         TabIndex        =   3
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "&Envia Sugestão"
         Height          =   375
         Left            =   3000
         TabIndex        =   2
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   $"frmSugestao.frx":0442
         Height          =   1575
         Left            =   2880
         TabIndex        =   1
         Top             =   360
         Width           =   3735
      End
      Begin VB.Image Image1 
         Height          =   1230
         Left            =   240
         Picture         =   "frmSugestao.frx":055E
         Top             =   480
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmSugestao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEmail_Click()
    ' Aqui vem a rotina para abrir o programa de email
    ' do usuário.
    ' Incluir campo Para: marciusbezerra@hotmail.com e
    ' luisherrera@ig.com.br
    ' no campo assunto: Sugestão para o Software Click Código
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' Centraliza formulário
    Me.Move frmAmbiente.ScaleWidth / 2 - Me.ScaleWidth / 2, _
            frmAmbiente.ScaleHeight / 2 - Me.ScaleHeight / 2
End Sub
