VERSION 5.00
Begin VB.Form frmInserir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inserir Cliente"
   ClientHeight    =   3900
   ClientLeft      =   135
   ClientTop       =   1005
   ClientWidth     =   10515
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   10515
   Begin VB.CommandButton cmdInserir 
      Caption         =   "Inserir"
      Height          =   495
      Left            =   3960
      TabIndex        =   12
      Top             =   2760
      Width           =   1335
   End
   Begin VB.TextBox txtDados 
      Height          =   1245
      Index           =   11
      Left            =   5880
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Text            =   "frmInserir.frx":0000
      Top             =   2160
      Width           =   4215
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   10
      Left            =   7920
      TabIndex        =   10
      Text            =   "Telefone 2"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   9
      Left            =   7920
      TabIndex        =   9
      Text            =   "Telefone 1"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   8
      Left            =   6000
      TabIndex        =   8
      Text            =   "CPF / CNPJ"
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   7
      Text            =   "Nº do Pedido 3"
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Text            =   "Nº do Pedido 2"
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   5
      Text            =   "Nº do Pedido 1"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   4
      Text            =   "Cidade"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Text            =   "Bairro"
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   2
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   2
      Text            =   "Endereço"
      Top             =   1200
      Width           =   5535
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   1
      Text            =   "Aos Cuidados"
      Top             =   720
      Width           =   3375
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   0
      Left            =   1320
      MaxLength       =   255
      TabIndex        =   0
      Text            =   "Nome"
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblDados 
      Caption         =   "Observações:"
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   24
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblDados 
      Caption         =   "Telefone 2:"
      Height          =   255
      Index           =   10
      Left            =   6960
      TabIndex        =   23
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblDados 
      Caption         =   "Telefone 1:"
      Height          =   255
      Index           =   9
      Left            =   6960
      TabIndex        =   22
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblDados 
      Caption         =   "CPF / CNPJ:"
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   21
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblDados 
      Caption         =   "Nº Pedido 3:"
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   20
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblDados 
      Caption         =   "Nº Pedido 2:"
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label lblDados 
      Caption         =   "Nº Pedido 1:"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblDados 
      Caption         =   "Cidade:"
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   17
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDados 
      Caption         =   "Bairro:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   16
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDados 
      Caption         =   "Endereço:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblDados 
      Caption         =   "A/c:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   14
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblDados 
      Caption         =   "Nome:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmInserir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    For x = 0 To txtDados.Count - 1
        txtDados(x).Text = ""
    Next x
    
End Sub
Private Sub cmdInserir_Click()
    
    'txtDados(0).Text  - Nome
    'txtDados(1).Text  - Aos Cuidados
    'txtDados(2).Text  - Endereço
    'txtDados(3).Text  - Bairro
    'txtDados(4).Text  - Cidade
    'txtDados(5).Text  - Nº do Pedido 1
    'txtDados(6).Text  - Nº do Pedido 2
    'txtDados(7).Text  - Nº do Pedido 3
    'txtDados(8).Text  - CPF / CNPJ
    'txtDados(9).Text  - Telefone 1
    'txtDados(10).Text - Telefone 2
    'txtDados(11).Text - Observções
    
    Dim Inserir As Boolean
    
    Inserir = FunçõesDoBanco.InserirCliente(txtDados(0).Text, txtDados(1).Text, txtDados(2).Text, txtDados(3).Text, txtDados(4).Text, _
    txtDados(8).Text, txtDados(9).Text, txtDados(10).Text, txtDados(11).Text, txtDados(5).Text, txtDados(6).Text, txtDados(7).Text)
        
    If (Inserir = True) Then
        For x = 0 To txtDados.Count - 1
            txtDados(x).Text = ""
        Next x
    End If
    
End Sub
Private Sub txtDados_KeyPress(Index As Integer, KeyAscii As Integer)
    
    If (KeyAscii = 13 And Index < txtDados.Count - 1) Then
        txtDados(Index + 1).SetFocus
    End If
    
End Sub
