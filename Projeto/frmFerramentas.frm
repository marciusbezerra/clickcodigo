VERSION 5.00
Begin VB.Form frmFerramentas 
   BackColor       =   &H003A3A59&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro das Ferramentas"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6270
   Begin VB.Frame Frame1 
      BackColor       =   &H003A3A59&
      Caption         =   "Ferramentas de Desenvolvimento:"
      ForeColor       =   &H00C0E0FF&
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   5895
      Begin VB.CommandButton cmdFechar 
         Caption         =   "&Fechar"
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   1800
         Width           =   1200
      End
      Begin VB.CommandButton cmdExcluir 
         Caption         =   "&Excluir"
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   1320
         Width           =   1200
      End
      Begin VB.CommandButton cmdAlterar 
         Caption         =   "&Alterar"
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   840
         Width           =   1200
      End
      Begin VB.CommandButton cmdIncluir 
         Caption         =   "&Incluir"
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   240
         Width           =   1200
      End
      Begin VB.TextBox txtFerramenta 
         BackColor       =   &H00C0E0FF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3975
      End
      Begin VB.ListBox lstFerramentas 
         BackColor       =   &H00C0E0FF&
         Height          =   1425
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   3975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         X1              =   0
         X2              =   5880
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmFerramentas.frx":0000
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmFerramentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFechar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ' Centraliza formulário
    Me.Move frmAmbiente.ScaleWidth / 2 - Me.ScaleWidth / 2, _
            frmAmbiente.ScaleHeight / 2 - Me.ScaleHeight / 2
End Sub
