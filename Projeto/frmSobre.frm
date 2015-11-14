VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSobre 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "frmSobre.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7410
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6960
      Top             =   3720
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   255
      Left            =   3120
      TabIndex        =   12
      Top             =   3840
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.PictureBox picContainer 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   60
         ScaleHeight     =   1500
         ScaleWidth      =   7095
         TabIndex        =   13
         Top             =   2040
         Width           =   7095
         Begin RichTextLib.RichTextBox rtf 
            Height          =   465
            Left            =   0
            TabIndex        =   14
            Top             =   105
            Width           =   7035
            _ExtentX        =   12409
            _ExtentY        =   820
            _Version        =   393217
            BackColor       =   16777215
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            Appearance      =   0
            TextRTF         =   $"frmSobre.frx":0742
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H003A3A59&
         Height          =   2055
         Left            =   0
         Picture         =   "frmSobre.frx":07BE
         ScaleHeight     =   1995
         ScaleWidth      =   7155
         TabIndex        =   1
         Top             =   0
         Width           =   7215
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Atualização:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Index           =   6
            Left            =   3360
            TabIndex        =   11
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "16 de Fevereiro 2003"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   5
            Left            =   4560
            TabIndex        =   10
            Top             =   1080
            Width           =   2175
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Index           =   4
            Left            =   3360
            TabIndex        =   9
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "E-mail:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   8
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Site:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Index           =   2
            Left            =   3360
            TabIndex        =   7
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0FFFF&
            Height          =   255
            Index           =   1
            Left            =   3360
            TabIndex        =   6
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "luisherrera@ig.com.br"
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   0
            Left            =   4080
            TabIndex        =   5
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "marciusbezerra@hotmail.com"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4080
            TabIndex        =   4
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "www.portalbrazuca.kit.net"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4080
            TabIndex        =   3
            Top             =   360
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Marcius Bezerra e Luis Herrera"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   4320
            TabIndex        =   2
            Top             =   120
            Width           =   2295
         End
      End
   End
End
Attribute VB_Name = "frmSobre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thetop As Long
Dim p1hgt As Long
Dim p1wid As Long
Dim theleft As Long
Dim Tempstring As String


Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo Erro
    Dim Linhas As Long
    rtf.LoadFile CaminhoApp & "..\Documentação\Historico ClickCodigo.rtf", rtfRTF
    Linhas = rtf.GetLineFromChar(Len(rtf.Text)) + 1
    rtf.Height = (Linhas * picContainer.TextHeight("A"))
    rtf.Top = picContainer.Height
    Timer1.Enabled = True
    Timer1.Interval = 100
    Exit Sub
Erro:
    If Err = 75 Then
        rtf.Text = "O arquivo Historico ClickCodigo.rtf deve estar na pasta Documentação"
        Resume Next
    Else
        rtf.Text = "Ocorreu o erro " & Err & " - " & Error$
        Resume Next
    End If
End Sub

Sub Timer1_Timer()
     rtf.Move 0, rtf.Top - 15
    If rtf.Top <= -rtf.Height Then rtf.Top = picContainer.Height
End Sub
