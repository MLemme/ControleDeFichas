VERSION 5.00
Begin VB.Form frmAlterar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alterar Dados de Cliente"
   ClientHeight    =   3600
   ClientLeft      =   135
   ClientTop       =   1005
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   10350
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
      Index           =   3
      Left            =   1320
      TabIndex        =   3
      Text            =   "Bairro"
      Top             =   1680
      Width           =   2535
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
      Index           =   5
      Left            =   6000
      TabIndex        =   5
      Text            =   "CPF / CNPJ"
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   6
      Left            =   7920
      TabIndex        =   6
      Text            =   "Telefone 1"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txtDados 
      Height          =   285
      Index           =   7
      Left            =   7920
      TabIndex        =   7
      Text            =   "Telefone 2"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtDados 
      Height          =   1245
      Index           =   8
      Left            =   5880
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "frmAlterar.frx":0000
      Top             =   2160
      Width           =   4215
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblId 
      Caption         =   "0"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblDados 
      Caption         =   "Nome:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   18
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblDados 
      Caption         =   "A/c:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   17
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblDados 
      Caption         =   "Endereço:"
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblDados 
      Caption         =   "Bairro:"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   15
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDados 
      Caption         =   "Cidade:"
      Height          =   255
      Index           =   4
      Left            =   3960
      TabIndex        =   14
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label lblDados 
      Caption         =   "CPF / CNPJ:"
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   13
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblDados 
      Caption         =   "Telefone 1:"
      Height          =   255
      Index           =   9
      Left            =   6960
      TabIndex        =   12
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblDados 
      Caption         =   "Telefone 2:"
      Height          =   255
      Index           =   10
      Left            =   6960
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblDados 
      Caption         =   "Observações:"
      Height          =   255
      Index           =   11
      Left            =   4680
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "frmAlterar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterar_Click()
    
    'txtDados(0).Text  - Nome
    'txtDados(1).Text  - Aos Cuidados
    'txtDados(2).Text  - Endereço
    'txtDados(3).Text  - Bairro
    'txtDados(4).Text  - Cidade
    'txtDados(5).Text  - CPF / CNPJ
    'txtDados(6).Text  - Telefone 1
    'txtDados(7).Text  - Telefone 2
    'txtDados(8).Text  - Observações
    
    Dim Alterar As Boolean
    Dim Primeiro As Boolean
    
    Alterar = FunçõesDoBanco.AlterarCliente(lblId.Caption, txtDados(0).Text, txtDados(1).Text, txtDados(2).Text, txtDados(3).Text, txtDados(4).Text, _
    txtDados(5).Text, txtDados(6).Text, txtDados(7).Text, txtDados(8).Text)

    If (Alterar = True) Then
        Dim SQLComand As String
    
        SQLComand = "SELECT nome,cpfcnpj,endereco,idcliente FROM Cliente ORDER BY nome"
        
        Set record = CreateObject("ADODB.Recordset")
        record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
        record.MoveLast
        With frmClientes
            .lsvClientes.ListItems.Clear
            
            While Not record.BOF
                Set linhas = .lsvClientes.ListItems.Add(1, , record!Nome)
'                If (record!CpfCnpj <> "") Then
'                    linhas.SubItems(1) = record!CpfCnpj
'                End If
                SQLComand = "SELECT numeropedido FROM Pedidos WHERE idCliente = " & record!idCliente
        
                Set record_pedido = CreateObject("ADODB.Recordset")
                record_pedido.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
                If record_pedido.RecordCount <> 0 Then
                    record_pedido.MoveFirst
                    
                    If (record_pedido.RecordCount <> 1) Then
                        Primeiro = True
                        While Not record_pedido.EOF
                            
                            If (Primeiro = True) Then
                                Pedido = record_pedido!numeropedido
                                Primeiro = False
                            Else
                                Pedido = Pedido & " / " & record_pedido!numeropedido
                            End If
                            
                            record_pedido.MoveNext
                        Wend
                        
                        linhas.SubItems(1) = Pedido
                        
                    Else
                    
                        linhas.SubItems(1) = record_pedido!numeropedido
                        
                    End If
                Else
                    linhas.SubItems(1) = "Não há PEDIDOS"
                End If
                If (record!Endereco <> "") Then
                    linhas.SubItems(2) = record!Endereco
                End If
                record.MovePrevious
            Wend
        End With
        record.Close
        Me.Hide
    End If

End Sub
Private Sub Form_Load()
    
    For x = 0 To txtDados.Count - 1
        txtDados(x).Text = ""
    Next x
    
End Sub
