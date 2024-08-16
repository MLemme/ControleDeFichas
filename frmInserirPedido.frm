VERSION 5.00
Begin VB.Form frmInserirPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inserir Novo Pedido"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton cmdInserirPedido 
      Caption         =   "Inserir"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtNumeroPedido 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Text            =   "000000"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblIdCliente 
      Caption         =   "0"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Nr do Pedido:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblEndereco 
      Caption         =   "Endereço"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label lblNome 
      Caption         =   "Nome"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmInserirPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInserirPedido_Click()
    Dim SQLComand As String
    
    If (txtNumeroPedido.Text = "" Or txtNumeroPedido.Text = Null Or txtNumeroPedido.Text = Empty) Then
        response = MsgBox("Não foi preenchido nenhum número de pedido", vbCritical, "Erro")
        Exit Sub
    End If
    
    SQLComand = "INSERT INTO Pedidos(idCliente,numeropedido) VALUES (" & lblIdCliente.Caption & ",'" & txtNumeroPedido.Text & "')"
    
    connect_banco.Execute SQLComand
    
    response = MsgBox("Pedido inserido com sucesso", vbInformation, "Sucesso !")
    
    Me.Hide
End Sub
