VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVisualizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dados de Cliente"
   ClientHeight    =   4755
   ClientLeft      =   135
   ClientTop       =   1005
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8160
   Begin MSComctlLib.ListView lsvPedidos 
      Height          =   1335
      Left            =   5400
      TabIndex        =   18
      Top             =   2160
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   2355
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblObservacoes 
      BackColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1440
      TabIndex        =   17
      Top             =   3720
      Width           =   6615
   End
   Begin VB.Label lblTelefone2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   16
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label lblTelefone1 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   15
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label lblCpfCnpj 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   14
      Top             =   2160
      Width           =   3255
   End
   Begin VB.Label lblDados 
      Caption         =   "Observações:"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lblDados 
      Caption         =   "Telefone 2:"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblDados 
      Caption         =   "Telefone 1:"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label lblDados 
      Caption         =   "CPF / CNPJ:"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblDados 
      Caption         =   "Cidade:"
      Height          =   255
      Index           =   4
      Left            =   4560
      TabIndex        =   9
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDados 
      Caption         =   "Bairro:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDados 
      Caption         =   "Endereço:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblDados 
      Caption         =   "A/c:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblDados 
      Caption         =   "Nome:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblCidade 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblBairro 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblEndereco 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   6975
   End
   Begin VB.Label lblAosCuidados 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   6975
   End
   Begin VB.Label lblNome 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "frmVisualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
