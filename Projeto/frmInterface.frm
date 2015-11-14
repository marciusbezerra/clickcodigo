VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInterface 
   BackColor       =   &H003A3A59&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Itens"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9705
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9705
   Begin VB.Frame frame1 
      Height          =   4935
      Index           =   2
      Left            =   9180
      TabIndex        =   47
      Top             =   5865
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CommandButton cmdCopiarApostila 
         Caption         =   "&Copiar"
         Height          =   375
         Left            =   8280
         TabIndex        =   50
         ToolTipText     =   "Copia o Texto"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelecionarApostila 
         Caption         =   "&Selecionar"
         Height          =   375
         Left            =   6840
         TabIndex        =   49
         ToolTipText     =   "Seleciona Todo o Texto"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.Frame Frame6 
         Caption         =   "Word(DOC) / HTML:"
         Height          =   4095
         Index           =   1
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   9495
         Begin RichTextLib.RichTextBox rtfApostila 
            Height          =   3765
            Left            =   105
            TabIndex        =   51
            Top             =   225
            Width           =   9255
            _ExtentX        =   16325
            _ExtentY        =   6641
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmInterface.frx":0000
         End
      End
   End
   Begin VB.Data datProjetos 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   360
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tblProjetos"
      Top             =   120
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.TextBox txtFiltro 
      BackColor       =   &H00C0E0FF&
      DataField       =   "Titulo"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4980
      MaxLength       =   50
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   960
      Width           =   3270
   End
   Begin VB.CommandButton cmdFiltra 
      Caption         =   "&Filtrar"
      Height          =   285
      Left            =   8310
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   945
      Width           =   990
   End
   Begin VB.Frame frame1 
      Height          =   4935
      Index           =   1
      Left            =   8700
      TabIndex        =   37
      Top             =   5595
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Frame Frame5 
         Caption         =   "Código / Artigo:"
         Height          =   4095
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   240
         Width           =   9495
         Begin VB.TextBox txtCodigoArtigo 
            DataSource      =   "datProjetos"
            ForeColor       =   &H00800000&
            Height          =   3735
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Top             =   240
            Width           =   9255
         End
      End
      Begin VB.CommandButton cmdSeleciona 
         Caption         =   "&Selecionar"
         Height          =   375
         Left            =   6840
         TabIndex        =   39
         ToolTipText     =   "Seleciona Todo o Texto"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdCopiarCodigo 
         Caption         =   "&Copiar"
         Height          =   375
         Left            =   8280
         TabIndex        =   38
         ToolTipText     =   "Copia o Texto"
         Top             =   4440
         Width           =   1335
      End
   End
   Begin VB.Frame frame1 
      Height          =   4935
      Index           =   0
      Left            =   0
      TabIndex        =   21
      Top             =   2280
      Width           =   9735
      Begin VB.Data datListFerramentas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT * FROM tblFerramentas ORDER BY Ferramenta"
         Top             =   2160
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Data datListFontes 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   6960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "SELECT * FROM tblFontes ORDER BY Fonte"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Frame frame2 
         Caption         =   "Lista de projetos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3975
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   840
         Width           =   3135
         Begin MSDBCtls.DBList lstProjetos 
            Bindings        =   "frmInterface.frx":0082
            Height          =   3555
            Left            =   120
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   6271
            _Version        =   393216
            IntegralHeight  =   0   'False
            BackColor       =   12640511
            ListField       =   "Titulo"
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Informações deste Registro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3975
         Left            =   3360
         TabIndex        =   23
         Top             =   840
         Width           =   6330
         Begin VB.CommandButton cmdArquivo 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   5960
            TabIndex        =   25
            Top             =   1680
            Width           =   270
         End
         Begin VB.TextBox txtComentário 
            BackColor       =   &H00C0E0FF&
            DataField       =   "Comentario"
            DataSource      =   "datProjetos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1170
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   2700
            Width           =   4410
         End
         Begin VB.TextBox txtArquivo 
            BackColor       =   &H00C0E0FF&
            DataField       =   "Arquivo"
            DataSource      =   "datProjetos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   1680
            Width           =   4140
         End
         Begin VB.TextBox txtTitulo 
            BackColor       =   &H00C0E0FF&
            DataField       =   "Titulo"
            DataSource      =   "datProjetos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   5
            Top             =   600
            Width           =   5370
         End
         Begin VB.TextBox txtCodigo 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            DataField       =   "Codigo"
            DataSource      =   "datProjetos"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   825
            Locked          =   -1  'True
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   240
            Width           =   1000
         End
         Begin VB.TextBox txtAutor 
            BackColor       =   &H00C0E0FF&
            DataField       =   "Autor"
            DataSource      =   "datProjetos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   6
            Top             =   950
            Width           =   2490
         End
         Begin VB.TextBox txtVersao 
            BackColor       =   &H00C0E0FF&
            DataField       =   "Versao"
            DataSource      =   "datProjetos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4200
            MaxLength       =   50
            TabIndex        =   7
            Top             =   950
            Width           =   2010
         End
         Begin MSComCtl2.DTPicker txtCriacao 
            DataField       =   "Criacao"
            DataSource      =   "datProjetos"
            Height          =   285
            Left            =   1800
            TabIndex        =   10
            Top             =   2010
            Width           =   4410
            _ExtentX        =   7779
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   22740992
            CurrentDate     =   37669
            MinDate         =   367
         End
         Begin MSComCtl2.DTPicker txtTermino 
            DataField       =   "Termino"
            DataSource      =   "datProjetos"
            Height          =   285
            Left            =   1800
            TabIndex        =   11
            Top             =   2355
            Width           =   4410
            _ExtentX        =   7779
            _ExtentY        =   503
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CheckBox        =   -1  'True
            Format          =   22740992
            CurrentDate     =   36368
         End
         Begin MSDBCtls.DBCombo cmbFerramentas 
            Bindings        =   "frmInterface.frx":009C
            DataField       =   "codFerramenta"
            DataSource      =   "datProjetos"
            Height          =   315
            Left            =   1800
            TabIndex        =   8
            Top             =   1320
            Width           =   4390
            _ExtentX        =   7752
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            BackColor       =   12640511
            ListField       =   "Ferramenta"
            BoundColumn     =   "ID"
            Text            =   ""
         End
         Begin MSDBCtls.DBCombo cmbFonte 
            Bindings        =   "frmInterface.frx":00BD
            DataField       =   "codFonte"
            DataSource      =   "datProjetos"
            Height          =   315
            Left            =   2650
            TabIndex        =   4
            Top             =   240
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
            Style           =   2
            BackColor       =   12640511
            ListField       =   "Fonte"
            BoundColumn     =   "ID"
            Text            =   ""
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackColor       =   &H003A3A59&
            BackStyle       =   0  'Transparent
            Caption         =   "Ferramenta:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   35
            Top             =   1365
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackColor       =   &H003A3A59&
            BackStyle       =   0  'Transparent
            Caption         =   "Comentário:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   6
            Left            =   480
            TabIndex        =   34
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackColor       =   &H003A3A59&
            BackStyle       =   0  'Transparent
            Caption         =   "Caminho/Arquivo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   33
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackColor       =   &H003A3A59&
            BackStyle       =   0  'Transparent
            Caption         =   "Término Criação:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   32
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackColor       =   &H003A3A59&
            BackStyle       =   0  'Transparent
            Caption         =   "Início Criacao:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   2040
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackColor       =   &H003A3A59&
            BackStyle       =   0  'Transparent
            Caption         =   "Fonte:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1905
            TabIndex        =   30
            Top             =   270
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackColor       =   &H003A3A59&
            BackStyle       =   0  'Transparent
            Caption         =   "Titulo:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   105
            TabIndex        =   29
            Top             =   620
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackColor       =   &H003A3A59&
            BackStyle       =   0  'Transparent
            Caption         =   "Nº:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   28
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackColor       =   &H003A3A59&
            BackStyle       =   0  'Transparent
            Caption         =   "Autor:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   27
            Top             =   950
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            BackColor       =   &H003A3A59&
            BackStyle       =   0  'Transparent
            Caption         =   "Versão:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   3360
            TabIndex        =   26
            Top             =   960
            Width           =   735
         End
      End
      Begin VB.Frame frmOpts 
         Caption         =   "Tipo de dados:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   4815
         Begin VB.OptionButton optTipo 
            Caption         =   "Código"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   0
            ToolTipText     =   "Scripts, Funções, Módulos completos"
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Artigo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   1
            ToolTipText     =   "Texto sobre determinado assunto"
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Apostila"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   2
            ToolTipText     =   "Tutoriais, apostilas e manuais sobre a ferramenta"
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optTipo 
            Caption         =   "Projeto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   3
            ToolTipText     =   "Programa com códifo fonte completo"
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStripDados 
      Height          =   375
      Left            =   0
      TabIndex        =   20
      ToolTipText     =   "Selecione os dados que deseja visualizar"
      Top             =   1920
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   661
      TabWidthStyle   =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Informações"
            Key             =   "aba1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Código / Artigo"
            Key             =   "aba2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Word / HTML"
            Key             =   "aba3"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdEditar 
      Height          =   570
      Left            =   6550
      Picture         =   "frmInterface.frx":00D9
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Editrar Item Selecionado"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   570
      Left            =   5940
      Picture         =   "frmInterface.frx":03E3
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Adicionar Novo Item Neste Grupo"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0C0&
      Height          =   570
      Left            =   7170
      Picture         =   "frmInterface.frx":06ED
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Excluir Item Selecionado"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdCancela 
      Height          =   570
      Left            =   8400
      Picture         =   "frmInterface.frx":09F7
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      ToolTipText     =   "Desfazer Alterações Neste Item"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdUpdate 
      Height          =   570
      Left            =   7780
      Picture         =   "frmInterface.frx":0D01
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      ToolTipText     =   "Gravar Item"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdClose 
      Height          =   570
      Left            =   9000
      Picture         =   "frmInterface.frx":100B
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Sair do Programa"
      Top             =   120
      Width           =   615
   End
   Begin MSComctlLib.TabStrip tabFerramentas 
      Height          =   375
      Left            =   0
      TabIndex        =   42
      ToolTipText     =   "Selecione uma Categoria"
      Top             =   1560
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   661
      TabWidthStyle   =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgArquivo 
      Left            =   3000
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H003A3A59&
      Caption         =   "Pesquisa Campo Comentário:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Index           =   8
      Left            =   4140
      TabIndex        =   45
      Top             =   720
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H003A3A59&
      BorderColor     =   &H00FFFFFF&
      Height          =   480
      Left            =   3960
      Shape           =   4  'Rounded Rectangle
      Top             =   855
      Width           =   5625
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H003A3A59&
      BackStyle       =   0  'Transparent
      Caption         =   "Critério:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   4140
      TabIndex        =   46
      Top             =   975
      Width           =   750
   End
   Begin VB.Image imgLogo 
      Height          =   1245
      Left            =   -120
      Picture         =   "frmInterface.frx":1315
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    ' Marcius [06/03/2003]:
    ' EDITANDO É UMA VARIÁVEL IMPORTANTE NO SISTEMA
    ' ELA NOS INFORMA, DE QUALQUER PARTE DO SISTEMA
    ' SE O REGISTRO ATUAL ESTÁ OU NÃO SENDO EDITADO
Dim Editando As Boolean
'
Dim F As String

Private Sub cmdFechar_Click()
    ' Descarrega formulário
    Unload Me
End Sub

Private Sub cmdCopiarApostila_Click()
    ' Marcius [06/03/2003]:
    ' COPIA O TRECHO SELECIONADO NA CAIXA APOSTILA
    Copiar rtfApostila
End Sub

Private Sub cmdCopiarCodigo_Click()
    ' Marcius [06/03/2003]:
    ' COPIA O TEXTO SELECIONADO NA CAIXA CÓDIGO
    Copiar txtCodigoArtigo
End Sub

Private Sub cmdSeleciona_Click()
    ' Marcius [06/03/2003]:
    '
    Selecionar txtCodigoArtigo
End Sub

Private Sub cmdSelecionarApostila_Click()
    ' Marcius [06/03/2003]:
    '
    Selecionar rtfApostila
End Sub

Private Sub Form_Load()
    ' Luis[16/02/2003]:
    ' Centraliza formulário dentro MDIPai
    Me.Move frmAmbiente.ScaleWidth / 2 - Me.ScaleWidth / 2, _
            frmAmbiente.ScaleHeight / 2 - Me.ScaleHeight / 2
            
    ' Luis[19/02/2003]:
    ' Reposiciona os frames (1 e 2) sobre o frame (0)
    frame1(1).Left = 0
    frame1(1).Top = 2280
    frame1(2).Left = 0
    frame1(2).Top = 2280
    
    ' Marcius [20/02/2003]:
    ' O código para o path foi modificado para servir para todos os
    ' tipos de SO's
    ' Função pega banco (módulo basGlobal)
    PegaBanco Me
    
    ' Marcius[20/02/2003]:
    ' Função preenche Abas da TabStrip Ferramentas
    PreencheTabs
    
    ' Luis[16/02/2003]:
    ' Altera data do calendário para data do sistema (iniciar em)
    
    ' Marcius[06/03/2003]:
    'AQUI VOCÊ ESTARÁ MODIFICANDO O VALOR DE UM REGISTRO A CADA VEZ QUE
    'O FORM FOR INICIADO, SEM QUE O USUÁRIO SAIBA
    
    'txtCriacao.Value = Format(Now, "dd/mm/yyyy")
    'txtTermino.Value = Format(Now, "dd/mm/yyyy")
    
    ' função ativa controles
    Trava True
    
    ' Marcius[20/02/2003]:
    ' ??? - Pode excluir ou será usado depois?
    ' PODE EXCLUIR
    'Filtrar "Visual Basic"
End Sub

' Marcius[20/02/2003]:
' ??? - Tudo (MN, TypeOf, %1, PopupMenu, etc...)
' ESTA FUNÇÃO FAZ COM QUE APAREÇA UM MENU POPUP
' AO CLIQUE DO BOTÃO DIREITO DO MOUSE, SOBRE AS TABS
' TORNANDO POSSÍVEL A ADIÇÃO DE NOVAS TABS (CATEGORIAS)
Private Sub tabFerramentas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim MN As Object
    If Button = vbRightButton Then
        For Each MN In frmAmbiente
            If TypeOf MN Is Menu Then _
                MN.Caption = Replace(MN.Caption, "%1", _
                    tabFerramentas.SelectedItem.Caption)
        Next MN
        PopupMenu frmAmbiente.mnuPopFerramentas
    End If
End Sub

Private Sub TabStripDados_Click()
    ' Luis[19/02/2003]:
    ' Ao alterar a aba do controle TabStripDados, o frame correspondente
    ' é exibido
    With Me.TabStripDados.SelectedItem
        frame1(0).Visible = .Key = "aba1"
        frame1(1).Visible = .Key = "aba2"
        frame1(2).Visible = .Key = "aba3"
    End With
End Sub

Private Sub tabFerramentas_Click()
    ' Marcius[06/03/2003]:
    ' FILTRA DE ACORDO COM O FILTRO SELECIONADO NA CAIXA CRITÉRIO
    Filtrar Trim(Me.txtFiltro.Text)

    ' Luis[21/02/2003]:
    ' Reposicionei a TabStrip na (1ª aba) e deixei o frame1(0) visível.
    
    ' PORQUE NÃO DEIXOU COMO O USUÁRIO HAVERIA DE TER DEIXADO?
    
    TabStripDados.Tabs(1).Selected = True
    Me.frame1(1).Visible = False
    Me.frame1(2).Visible = False
End Sub

' Marcius [06/03/2003]:
' SE AINDA HOUVER UMA EDIÇÃO EM PROCESSO, AO SER FECHADO, O PROGRAMA AVISARÁ
' UMA TAREFA:
' FAÇA COM QUE FIQUE MAIS PROFISSIONAL: O PROGRAMA PERGUNTARÁ SE O USUÁRIO DESEJA
' SALVAR AS ALTERAÇÕES, E ESSE PODERÁ RESPONDER, SIM, NÃO, OU CANCELAR.
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Editando Then
        MsgBox "Existem transações de dados " & _
            "que ainda estão pendentes." & vbCrLf & _
            "Salve-as primeiro.", vbCritical, Caption
            Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Retorna o ícone padrão do mouse
    Screen.MousePointer = vbDefault
End Sub

Sub Trava(Travar As Boolean)
    Dim i As Integer
    ' Marcius[20/02/2003]:
    ' Termine o procedimento Travar aqui, quando o objeto não tiver
    ' a propriedade Locked (preferida) use Enabled.
    ' VOCÊ NÂO FEZ!?
    Me.cmbFonte.Locked = Travar
    Me.txtTitulo.Locked = Travar
    Me.txtAutor.Locked = Travar
    Me.txtVersao.Locked = Travar
    Me.cmbFerramentas.Locked = Travar
    Me.txtArquivo.Locked = Travar
    Me.cmdArquivo.Enabled = Not Travar
    Me.txtCriacao.Enabled = Not Travar
    Me.txtTermino.Enabled = Not Travar
    Me.txtComentário.Locked = Travar
    Me.txtCodigoArtigo.Locked = Travar
    Me.rtfApostila.Locked = Travar
    Me.lstProjetos.Enabled = Travar
    frmOpts.Enabled = Not Travar
    
    ' ??? - Pode excluir?
    ' PODE
    'Me.dbcNacional.Enabled = Not Travar
End Sub

Function Valida() As Boolean
    Valida = True
    
    ' Essa rotina faz a validação de todos os dados
    
    If PegaTipo < 0 Then
        MsgBox "O Tipo é obrigatório para todos os registros.", vbCritical, Caption
        Valida = False
        Me.optTipo(0).SetFocus
        Exit Function
    End If
    
    datProjetos.Recordset("Opcao") = PegaTipo
    
    ' Luis[21/02/2003]:
    ' Alterei as mensagens para ficarem coerentes com os dados cadastrados
    If Trim(Me.txtTitulo.Text) = "" Then
        MsgBox "O Título é obrigatório para todos os registros.", vbCritical, Caption
        Valida = False
        Me.txtTitulo.SetFocus
        Exit Function
    End If
    
    ' ???
    If cmbFerramentas.BoundText = "" Then
        MsgBox "Cadê a droga da ferramanta!?", vbCritical, Caption
        Valida = False
        Me.cmbFerramentas.SetFocus
        Exit Function
    End If
    
    ' Luis[21/02/2003]:
    ' Aqui tive de incluir uma verificação do tipo de dado cadastrado,
    ' e manter esse item para Apostilas e Projetos, além de incrementar
    ' a mensagem conforme o item selecionado.
    
    ' Marcius[06/03/2003]:
    ' ESSA VERIFICAÇÃO FOI OTIMIZADA
    If (optTipo(2).Value = True) Or (optTipo(3).Value = True) Then
        If Trim(Me.txtArquivo.Text) = "" Then
            MsgBox "O path do arquivo de " & IIf(optTipo(2).Value, "apostila", "projeto") & " é obrigatório.", vbCritical, Caption
            Valida = False
            Me.txtArquivo.SetFocus
            Exit Function
        End If
    End If
    
    ' Marcius[20/02/2003]:
    ' POR ENQUANTO, ESSE COMENTÁRIOS AINDA DEVE SER OBRIGATÓRIO
    ' POIS É O OBJETO DA BUSCA
    If Trim(Me.txtComentário.Text) = "" Then
        MsgBox "Você deve inserir um comentário qualquer para os registros.", vbCritical, Caption
        Valida = False
        Me.txtComentário.SetFocus
        Exit Function
    End If
    
    ' ???
    ' por quê duas vezes a mesma mensagem?
    ' por quê a comparação é feita pelo Nº do código, se este é autoincrementado?
    
    'PORQUE, CASO VOCÊ ESTEJA EDITANDO, AO SALVAR, O BANCO DE DADOS E O VB
    'VÃO PENSAR QUE VOCÊ ESTÁ INCLUINDO UM REGISTRO COM O MESMO NOME
    
    Dim RC As Recordset
    Set RC = Me.datProjetos.Recordset.Clone
    RC.FindFirst "Titulo = '" & Trim(Me.txtTitulo.Text) & "'"
    If Not RC.NoMatch Then
        If Me.datProjetos.EditMode = dbEditInProgress Then
            If CLng(RC("Codigo")) <> CLng(Me.datProjetos.Recordset("Codigo")) Then
                MsgBox "Já existe um projeto com este título.", vbInformation, Caption
                Valida = False
                RC.Close
                Exit Function
            Else
                RC.Close
            End If
        Else
            MsgBox "Já existe um projeto com este título.", vbInformation, Caption
            Valida = False
            RC.Close
            Exit Function
        End If
    Else
        RC.Close
    End If
End Function

' Marcius[20/02/2003/:
' ESSA FUNÇÃO FOI ELIMINADA, PODEMOS PENSAR NO USO DELA PARA
' O FUTURO:
'    Function Exten(Texto As String) As String
'        ' Alterei para incluir Doc e HTML, veja se fiz certo
'        Select Case UCase(Texto)
'            Case "VISUAL BASIC"
'                Exten = "Arquivos do Visual Basic|*.vbp;*.frm;*.cls;*.bas;*.dob;*.ctl;*.pag;*.res"
'            Case "MICROSOFT ACCESS"
'                Exten = "Arquivos do Access|*.mdb;*.mda;*.mdw;*.mdt"
'            Case "C / C++ / C#"
'                Exten = "Arquivos do C|*.c;*.h"
'            Case "BORLAND DELPHI"
'                Exten = "Arquivos do Delphi|*.dpr;*.pas;*.tpu"
'            Case "BORLAND PASCAL"
'                Exten = "Arquivos do Pascal|*.pas;*.tpu"
'            Case "MICROSOFT EXCEL"
'                Exten = "Pastas do Excel|*.xls;*.xlt"
'            Case "Apostilas MSWord"
'                Exten = "Pastas do Word|*.doc;*.dot;*.rtf"
'            Case "Apostilas HTML"
'                Exten = "Pastas HTML|*.htm;*.html"
'        End Select
'    End Function


' ???
' Marcius[06/03/2003]:
' FAZ COM QUE, AO CLIQUE NA LISTA, SEJA SELECIONADO O REGISTRO CORRESPONDENTE
Private Sub lstProjetos_Click()
    If Me.datProjetos.Recordset.RecordCount = 0 Then Exit Sub
    Me.datProjetos.Recordset.Bookmark = Me.lstProjetos.SelectedItem
End Sub

'???
Sub Filtrar(Optional Expressao As String)
    Dim Filt As String
    Dim Ferr As String
    
    ' ??? onde o Key desse controle é alimentado? Não localizei isso.
    
    ' Marcius[06/03/2003]:
    ' TEM CERTEZA QUE PROCUROU?
    ' NO PROCEDIMENTO PreencheTabs QUE FAZ COM QUE AS TABS SEJAM GERADAS
    ' DINAMICAMENTE, O QUE, ALIÁS, VOCÊ NEGOU TER SIDO FEITO
    
    Ferr = Mid(tabFerramentas.SelectedItem.Key, 3)
    
    ' ??? onde está essa variável expressao?
    ' Marcius[06/03/2003]:
    ' ESTÁ LOGO ACIMA, COMO STRING, E OPCIONAL, NO CABEÇALHO DO PROCEDIMENTO CORRENTE
    
    If Expressao = "" Then Expressao = "*"
    Me.datProjetos.RecordSource = "SELECT * FROM tblProjetos WHERE (Comentario like '*" & Expressao & "*' AND codFerramenta = " & Ferr & ") ORDER BY Titulo;"
    Me.datProjetos.Refresh
End Sub

Private Sub cmbFerramenta_Click()
    ' ???
    ' Marcius[06/03/2003]:
    ' AO TROCAR DE FERRAMENTA, O ARQUIVO DEVE SER APAGADO.
    ' ISSO É OPCIONAL, E PODE SER EXCLUÍDO
    Me.txtArquivo.Text = vbNullString
End Sub

Private Sub cmbFerramenta_KeyPress(KeyAscii As Integer)
    ' ???
    ' Marcius[06/03/2003]:
    ' NÃO É NECESSÁRIO, TRATA-SE DE SIMPLES TESTE QUE FAZIA
    'KeyAscii = 0
End Sub

Private Sub cmdAdd_Click()
    ' Os controles são desabilitados para que não possa haver edição
    ' ou habilitados
    Trava False
    ' Novo registro
    datProjetos.Recordset.AddNew
    ' Os botões não necessários são desabilitados
    Me.cmdAdd.Enabled = False
    Me.cmdCancela.Enabled = True
    Me.cmdClose.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdUpdate.Enabled = True
    Me.cmdEditar.Enabled = False
    Me.TabStripDados.Tabs(1).Selected = True
    Me.cmbFonte.SetFocus
    ' Marcius[06/03/2003]:
    ' DATA DE HOJE, POR DEFEITO
    Me.txtCriacao.Value = Now
    Me.txtTermino.Value = Me.txtCriacao.Value
    
    HabilitaTipo
    
    'Essa variável pode ser útil futuramente
    Editando = True
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Marcius[20/02/2003]:
    ' Exibe informação na barra de status do MDIPai, através da função Status
    ' do módulo basGlobal
    
    ' Marcius[06/03/2003]:
    'TAREFA:
    'VOCÊ DEVE, COM BASE NO EXEMPLO ABAIXO, FAZER O MESMO COM TODOS OS CONTROLES
    'DE TODOS OS FORMS
    Status "Adiciona novo conteúdo"
End Sub

Private Sub cmdArquivo_Click()
    ' Marcius[20/02/2003]:
    ' Um exemplo de tratamento de erro
    On Error GoTo Erro
    ' ???
    ' Marcius[06/03/2003]:
    ' CANCELAR A COMDLG, É TIDO POR NÓS, PROGRAMADORES, COMO O ERRO CDLCANCEL
    ' NÃO EXISTE UMA FORMA MELHOR
    With Me.dlgArquivo
        ' ???
        .CancelError = True
        ' ???
        ' Marcius[06/03/2003]:
        ' O ARQUIVO PRECISA EXISTIR E NÃO É NECESSÁRIA A CAIXA 'SOMENTE LEITURA'
        .Flags = cdlOFNFileMustExist Or cdlOFNNoReadOnlyReturn Or cdlOFNHideReadOnly
        .Filter = "Todos os arquivos (*.*)|*.*"
        ' ???
        ' Marcius[06/03/2003]:
        ' EXECUTA O DIÁLOGO 'ABRIR ARQUIVO'
        .ShowOpen
        Me.txtArquivo.Text = Me.dlgArquivo.FileName
    End With
    Exit Sub
Erro:
    If Err = cdlCancel Then
        'Nada faz...
    Else
        'Podemos fazer futuramente um tratamento de erro melhor
        MsgBox Err.Number & " # " & Error$, vbCritical, "!@#@$@#$#"
    End If
End Sub

Private Sub cmdCancela_Click()
    Dim MSG As String
    MSG = "Deseja desfazer as alterações ?"
    
    ' ??? - esse primeiro IF não precisa de End IF?
    ' Marcius[06/03/2003]:
    ' NÃO, PORQUE A CONDIÇÃO JÁ FOI FINALIZADA EM UMA SÓ LINHA
    If MsgBox(MSG, vbYesNo, Caption) = vbNo Then Exit Sub
    Trava True
    datProjetos.Recordset.CancelUpdate
    If Me.datProjetos.Recordset.RecordCount = 0 Then
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = False
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = False
    Else
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = True
    End If
    ' Marcius[06/03/2003]:
    ' MODIFICADO
    'Me.txtTitulo.SetFocus
    'Me.txtFonte.SetFocus
    Editando = False
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

' ??? - explicar cada linha por favor
Private Sub cmdDelete_Click()
' Marcius[06/03/2003]:
' SE NÃO EXITE REGISTROS, NADA FAZ..
    If Me.datProjetos.Recordset.RecordCount = 0 Then Exit Sub
    Dim MSG As String
    Dim Cod As Long
    ' TEXTO DA PERGUNTA DA EXLCLUSÃO
    MSG = "Deseja excluir o registro " & _
        Me.datProjetos.Recordset("Codigo") & " ?"
    ' A PERGUNTA SOBRE O ESCLUSÃO, SE RESPONDER NÃO, NADA FAZ
    If MsgBox(MSG, vbYesNo, Caption) = vbNo Then Exit Sub
    ' DEPOIS DE EXCLUÍDO, O REGISTRO DEVE SER POSICIONADO
    ' NA LINHA CORRETA, SENÃO UM ERRO É RETORNADO
  With datProjetos.Recordset
    Cod = CLng(.Fields("Codigo").Value)
    .Delete
    If .RecordCount <> 0 Then .MoveNext
    If .EOF Then .MoveLast
    If .RecordCount = 0 Then Me.datProjetos.Refresh
  End With
  ' ATUALIZA A LISTA
  Me.lstProjetos.ReFill
  Me.txtTitulo.SetFocus
End Sub

Private Sub cmdEditar_Click()
    Trava False
    datProjetos.Recordset.Edit
    Me.cmdAdd.Enabled = False
    Me.cmdCancela.Enabled = True
    Me.cmdClose.Enabled = False
    Me.cmdDelete.Enabled = False
    Me.cmdUpdate.Enabled = True
    Me.cmdEditar.Enabled = False
'    Me.txtTitulo.SetFocus
    Editando = True
    Me.datProjetos.Visible = False
End Sub

Private Sub cmdFiltra_Click()
    Filtrar Trim(Me.txtFiltro.Text)
End Sub

' ???
Private Sub cmdUpdate_Click()
    If Not Valida Then Exit Sub
    datProjetos.UpdateRecord
    datProjetos.Recordset.Bookmark = datProjetos.Recordset.LastModified
    Trava True
    If Me.datProjetos.Recordset.RecordCount = 0 Then
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = False
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = False
    Else
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = True
    End If
    Me.lstProjetos.ReFill
    Me.lstProjetos.Refresh
   ' Me.txtTitulo.SetFocus
    Editando = False
    cmdAdd.SetFocus
    ' Marcius[20/02/2003]:
    ' desativado V.1.0.1 -  pode excluir?
    ' Marcius[06/03/2003]:
    'Filtrar
End Sub

Private Sub datprojetos_Error(DataErr As Integer, Response As Integer)
    ' Luis[16/02/2003]:
    ' Acho que essa rotina precisa melhorar
    ' Marcius[20/02/2003]:
    ' Será melhorada no futuro
  MsgBox "Ocorreu o erro nº " & DataErr & ": " & Error$(DataErr) & _
    vbCrLf & vbCrLf & "Entre em contato com o fornecedor."
  Response = 0
End Sub

'???
Private Sub datProjetos_Reposition()
    If Me.datProjetos.Recordset.RecordCount = 0 Then
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = False
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = False
        HabilitaTipo
    Else
        Me.cmdAdd.Enabled = True
        Me.cmdCancela.Enabled = False
        Me.cmdClose.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdUpdate.Enabled = False
        Me.cmdEditar.Enabled = True
        HabilitaTipo IIf(IsNull(datProjetos.Recordset("Opcao")), -1, datProjetos.Recordset("Opcao"))
    End If
End Sub

' ??? - entendi (+/-) o que faz, mas não os comandos
Private Sub PreencheTabs(Optional TabAtual As String)
    Dim RC As DAO.Recordset
    tabFerramentas.Tabs.Clear
    Set RC = datListFerramentas.Recordset.Clone
    RC.MoveFirst
    Do Until RC.EOF
    'O COMANDO ADD ADICIONA UMA NOVA TAB EM TABFERRAMENTAS
        tabFerramentas.Tabs.Add , "ID" & Trim(RC("ID")), RC("Ferramenta")
        RC.MoveNext
    Loop
    If Trim(TabAtual) = "" Then
        tabFerramentas.Tabs(1).Selected = True
    Else
        tabFerramentas.Tabs(TabAtual).Selected = True
    End If
End Sub

    ' Marcius [06/03/2003]:
    ' SELECIONA O TEXTO EM UMA CAIXA DE TEXTO
Private Sub Selecionar(Caixa As Object)
    With Caixa
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

    ' Marcius [06/03/2003]:
    ' COPIA UM TEXTO SELECIONADO
Private Sub Copiar(Caixa As Object)
    If Caixa.SelLength > 0 Then
        With Clipboard
            .Clear
            .SetText Caixa.SelText
        End With
        Caixa.SetFocus
    End If
End Sub

    ' Marcius [06/03/2003]:
    ' HABILITA A OPÇÃO CORRETA, DEPENDENDO DO BANCO DE DADOS
Private Sub HabilitaTipo(Optional Tipo As Integer = -1)
    Dim i As Integer
    For i = 0 To optTipo.Count - 1
        optTipo(i).Value = False
    Next i
    If Tipo > -1 Then optTipo(Tipo).Value = True
End Sub

    ' Marcius [06/03/2003]:
    ' PEGA O VALOR DA OPÇÃO HABILITADA PARA POR NO BANCO
Private Function PegaTipo() As Integer
    Dim i As Integer
    PegaTipo = -1
    For i = 0 To optTipo.Count - 1
        If optTipo(i).Value = True Then
            PegaTipo = i
            Exit For
        End If
    Next i
End Function
