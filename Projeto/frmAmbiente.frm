VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.MDIForm frmAmbiente 
   BackColor       =   &H8000000C&
   Caption         =   "Click Código V.1.0 - Sistema de Gerenciamento de códigos, artigos, apostilas e projetos"
   ClientHeight    =   5040
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8250
   Icon            =   "frmAmbiente.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4770
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6350
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuSair 
         Caption         =   "Sai&r"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuDados 
      Caption         =   "Base &Dados"
      Begin VB.Menu mnuFerramentas 
         Caption         =   "&Ferramentas"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFontes 
         Caption         =   "F&ontes"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuInterface 
         Caption         =   "&Interface"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "A&juda"
      Begin VB.Menu mnuSugestao 
         Caption         =   "S&ugestões"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuSobre 
         Caption         =   "&Sobre"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuPopFerramentas 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuAdicionarUmaNovaFerramenta 
         Caption         =   "Adicionar nova ferramenta..."
      End
      Begin VB.Menu mnuEditarFerramenta 
         Caption         =   "Renomear %1..."
      End
      Begin VB.Menu mnuExcluirFerramenta 
         Caption         =   "Excluir %1"
      End
   End
End
Attribute VB_Name = "frmAmbiente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuAdicionarUmaNovaFerramenta_Click()
    MsgBox "Tente fazer essa parte Luís..."
End Sub

Private Sub MDIForm_Load()
    CaminhoApp = App.Path
    If Right(CaminhoApp, 1) <> "\" Then CaminhoApp = CaminhoApp & "\"
End Sub

Private Sub mnuFerramentas_Click()
    frmFerramentas.Show
End Sub

Private Sub mnuFontes_Click()
    frmFontes.Show
End Sub

Private Sub mnuInterface_Click()
    frmInterface.Show
End Sub

Private Sub mnuSair_Click()
    End
End Sub

Private Sub mnuSobre_Click()
    frmSobre.Show (1)
End Sub

Private Sub mnuSugestao_Click()
    frmSugestao.Show (1)
End Sub
