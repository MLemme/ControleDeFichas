VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes Cadastrados"
   ClientHeight    =   7890
   ClientLeft      =   135
   ClientTop       =   1005
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   9570
   Begin VB.CommandButton cmdInserirPedido 
      Caption         =   "Inserir Pedido"
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   6480
      Width           =   1335
   End
   Begin VB.OptionButton optBusca 
      Caption         =   "Aos Cuidados"
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   5
      Top             =   6360
      Width           =   1815
   End
   Begin VB.OptionButton optBusca 
      Caption         =   "Endereço"
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Top             =   5880
      Width           =   1815
   End
   Begin VB.OptionButton optBusca 
      Caption         =   "CPF / CNPJ"
      Height          =   375
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   6360
      Width           =   1815
   End
   Begin VB.OptionButton optBusca 
      Caption         =   "Nome"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   5880
      Value           =   -1  'True
      Width           =   1815
   End
   Begin VB.TextBox txtBusca 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   5280
      Width           =   4215
   End
   Begin VB.CommandButton cmdTodos 
      Caption         =   "Mostrar Todos"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   4800
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "Visualizar"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdApagar 
      Caption         =   "Apagar"
      Height          =   495
      Left            =   1680
      TabIndex        =   9
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   7200
      Width           =   1335
   End
   Begin MSComctlLib.ListView lsvClientes 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9500
      _ExtentX        =   16748
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterar_Click()
    Dim x As Double
    Dim Nome As String
    Dim Endereco As String
    Dim NaoSelecionado As Boolean
    Dim SQLComand As String
        
    NaoSelecionado = True
    For x = 1 To lsvClientes.ListItems.Count
        If (lsvClientes.ListItems.Item(x).Checked = True) Then
            Nome = lsvClientes.ListItems.Item(x).Text
            Endereco = lsvClientes.ListItems.Item(x).SubItems(2)
            NaoSelecionado = False
            Exit For
        End If
    Next x
    If (NaoSelecionado = True) Then
        response = MsgBox("Não foi selecionado nenhum Cliente", vbCritical, "Erro 02")
        Exit Sub
    End If
    
    SQLComand = "SELECT * FROM Cliente WHERE nome = '" & Nome & "' AND endereco = '" & Endereco & "'"
    
    Set record = CreateObject("ADODB.Recordset")
    record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
    
    With frmAlterar
           
        .lblId = record!idCliente
        .txtDados(0).Text = record!Nome
        If (record!AosCuidados <> "") Then
            .txtDados(1).Text = record!AosCuidados
        End If
        If (record!Endereco <> "") Then
            .txtDados(2).Text = record!Endereco
        End If
        If (record!Bairro <> "") Then
            .txtDados(3).Text = record!Bairro
        End If
        If (record!Cidade <> "") Then
            .txtDados(4).Text = record!Cidade
        End If
        If (record!CpfCnpj <> "") Then
            .txtDados(5).Text = record!CpfCnpj
        End If
        If (record!tel1 <> "") Then
            .txtDados(6).Text = record!tel1
        End If
        If (record!tel2 <> "") Then
            .txtDados(7).Text = record!tel2
        End If
        If (record!Observacoes <> "") Then
            .txtDados(8).Text = record!Observacoes
        End If
        
        record.Close
        .Show vbModal
        
    End With
    
End Sub
Private Sub cmdApagar_Click()
    Dim x As Double
    Dim Nome As String
    Dim Endereco As String
    Dim NaoSelecionado As Boolean
    Dim idCliente As String
    Dim Deletar As Boolean
    Dim Primeiro As Boolean
    
    NaoSelecionado = True
    For x = 1 To lsvClientes.ListItems.Count
        If (lsvClientes.ListItems.Item(x).Checked = True) Then
            Nome = lsvClientes.ListItems.Item(x).Text
            Endereco = lsvClientes.ListItems.Item(x).SubItems(2)
            NaoSelecionado = False
            Exit For
        End If
    Next x
    If (NaoSelecionado = True) Then
        response = MsgBox("Não foi selecionado nenhum Cliente", vbCritical, "Erro 02")
        Exit Sub
    End If
    
    response = MsgBox("Tem certeza que deseja apagar este registro?", vbOKCancel, "Apagar")
    '2 Cancel
    If (response = 2) Then
        lsvClientes.ListItems.Item(x).Checked = False
        Exit Sub
    End If
    '1 OK
    If (response = 1) Then
        SQLComand = "SELECT idCliente FROM Cliente WHERE nome = '" & Nome & "' AND endereco = '" & Endereco & "'"
    
        Set record = CreateObject("ADODB.Recordset")
        record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
        
        idCliente = record!idCliente
        
        record.Close
        
        Deletar = FunçõesDoBanco.DeletarCliente(idCliente)
        
        If (Deletar = True) Then
            SQLComand = "SELECT nome,cpfcnpj,endereco,idCliente FROM Cliente ORDER BY nome"
        
            Set record = CreateObject("ADODB.Recordset")
            record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
            record.MoveLast
                
            lsvClientes.ListItems.Clear
            
            While Not record.BOF
                Set linhas = lsvClientes.ListItems.Add(1, , record!Nome)
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
            
            record.Close
        End If
    End If
   
End Sub
Private Sub cmdBuscar_Click()
    Dim SQLComand As String
    Dim variavel As String
    Dim x As Integer
    Dim Primeiro As Boolean
    
    If (txtBusca.Text = "" Or txtBusca.Text = Null Or txtBusca.Text = Empty) Then
        response = MsgBox("Nenhum valor para busca foi fornecido", vbCritical, "Erro")
        Exit Sub
    End If
    
    For x = 0 To optBusca.Count - 1
        If (optBusca(x).Value = True) Then
            Exit For
        End If
    Next x
    
    Select Case x
        Case 0
            variavel = "nome"
        Case 1
            variavel = "cpfcnpj"
        Case 2
            variavel = "endereco"
        Case 3
            variavel = "aoscuidados"
    End Select
    
    SQLComand = "SELECT nome,cpfcnpj,endereco,idCliente FROM Cliente WHERE " & variavel & " LIKE '%" & txtBusca.Text & "%' ORDER BY nome"
    
    Set record = CreateObject("ADODB.Recordset")
    record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
    On Error GoTo 3021
    record.MoveLast
        
    lsvClientes.ListItems.Clear
    
    While Not record.BOF
        Set linhas = lsvClientes.ListItems.Add(1, , record!Nome)
'        If (record!CpfCnpj <> "") Then
'            linhas.SubItems(1) = record!CpfCnpj
'        End If
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
    
    cmdTodos.Enabled = True
3021:
    If (Err.Number = 3021) Then
        response = MsgBox("Não foi encontrado " & optBusca(x).Caption & " " & txtBusca.Text, vbInformation, "Erro")
    End If
    
    record.Close
    
End Sub

Private Sub cmdInserirPedido_Click()
    Dim x As Double
    Dim Nome As String
    Dim Endereco As String
    Dim NaoSelecionado As Boolean
    Dim idCliente As String
    Dim SQLComand As String
    
    NaoSelecionado = True
    For x = 1 To lsvClientes.ListItems.Count
        If (lsvClientes.ListItems.Item(x).Checked = True) Then
            Nome = lsvClientes.ListItems.Item(x).Text
            Endereco = lsvClientes.ListItems.Item(x).SubItems(2)
            NaoSelecionado = False
            Exit For
        End If
    Next x
    If (NaoSelecionado = True) Then
        response = MsgBox("Não foi selecionado nenhum Cliente", vbCritical, "Erro 02")
        Exit Sub
    End If
    
    SQLComand = "SELECT * FROM Cliente WHERE nome = '" & Nome & "' AND endereco = '" & Endereco & "'"
    
    Set record = CreateObject("ADODB.Recordset")
    record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
    
    idCliente = record!idCliente
    
    record.Close
    
    With frmInserirPedido
        .lblIdCliente.Caption = idCliente
        .lblNome = Nome
        .lblEndereco = Endereco
        .txtNumeroPedido = ""
        
        .Show vbModal
    End With
    
End Sub

Private Sub cmdTodos_Click()
    Dim SQLComand As String
    Dim Primeiro As Boolean
    
    SQLComand = "SELECT nome,cpfcnpj,endereco,idCliente FROM Cliente ORDER BY nome"
        
    Set record = CreateObject("ADODB.Recordset")
    record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
    record.MoveLast
        
    lsvClientes.ListItems.Clear
    
    While Not record.BOF
        Set linhas = lsvClientes.ListItems.Add(1, , record!Nome)
'        If (record!CpfCnpj <> "") Then
'            linhas.SubItems(1) = record!CpfCnpj
'        End If
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
    
    record.Close
    
    cmdTodos.Enabled = False
End Sub
Private Sub cmdVisualizar_Click()
    Dim x As Double
    Dim Nome As String
    Dim Endereco As String
    Dim NaoSelecionado As Boolean
    Dim idCliente As String
    Dim SQLComand As String
    
    NaoSelecionado = True
    For x = 1 To lsvClientes.ListItems.Count
        If (lsvClientes.ListItems.Item(x).Checked = True) Then
            Nome = lsvClientes.ListItems.Item(x).Text
            Endereco = lsvClientes.ListItems.Item(x).SubItems(2)
            NaoSelecionado = False
            Exit For
        End If
    Next x
    If (NaoSelecionado = True) Then
        response = MsgBox("Não foi selecionado nenhum Cliente", vbCritical, "Erro 02")
        Exit Sub
    End If
    
    SQLComand = "SELECT * FROM Cliente WHERE nome = '" & Nome & "' AND endereco = '" & Endereco & "'"
    
    Set record = CreateObject("ADODB.Recordset")
    record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
    
    
    With frmVisualizar
           
        .lblNome.Caption = record!Nome
        If (record!AosCuidados <> "") Then
            .lblAosCuidados.Caption = record!AosCuidados
        End If
        If (record!Endereco <> "") Then
            .lblEndereco.Caption = record!Endereco
        End If
        If (record!Bairro <> "") Then
            .lblBairro.Caption = record!Bairro
        End If
        If (record!Cidade <> "") Then
            .lblCidade.Caption = record!Cidade
        End If
        If (record!CpfCnpj <> "") Then
            .lblCpfCnpj.Caption = record!CpfCnpj
        End If
        If (record!tel1 <> "") Then
            .lblTelefone1.Caption = record!tel1
        End If
        If (record!tel2 <> "") Then
            .lblTelefone2.Caption = record!tel2
        End If
        If (record!Observacoes <> "") Then
            .lblObservacoes.Caption = record!Observacoes
        End If
        
        idCliente = record!idCliente
        
        record.Close
        
        SQLComand = "SELECT * FROM Pedidos WHERE idCliente = " & idCliente & " ORDER BY numeropedido"
    
        Set record = CreateObject("ADODB.Recordset")
        record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
        
        On Error GoTo 3021
        record.MoveLast
            
        Dim colunas As ColumnHeader
        .lsvPedidos.ColumnHeaders.Clear
        Set colunas = .lsvPedidos.ColumnHeaders.Add(1, , "Números de Pedido", 2450)
        
        While Not record.BOF
            Set linhas = .lsvPedidos.ListItems.Add(1, , record!numeropedido)
            record.MovePrevious
        Wend
    
3021:
       
        record.Close
        
        .Show vbModal
        
    End With
    
End Sub
Private Sub Form_Load()
    Dim SQLComand As String
    Dim Controle As Double
    Dim Primeiro As Boolean
    
    Controle = 0
    frmPrincipal.lblProgresso.Caption = "Carregando Clientes existentes no Banco de Dados..."
    frmPrincipal.lblProgresso.Visible = True
    frmPrincipal.Progresso.Visible = True
    
    
    SQLComand = "SELECT nome,cpfcnpj,endereco,idCliente FROM Cliente ORDER BY nome"
    
    Set record = CreateObject("ADODB.Recordset")
    record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
    On Error GoTo 3021
    record.MoveLast
        
    Dim colunas As ColumnHeader
    lsvClientes.ColumnHeaders.Clear
    Set colunas = lsvClientes.ColumnHeaders.Add(1, , "Nome", 3500)
    'Set colunas = lsvClientes.ColumnHeaders.Add(2, , "CPF / CNPJ", 2000)
    Set colunas = lsvClientes.ColumnHeaders.Add(2, , "Pedidos", 2000)
    Set colunas = lsvClientes.ColumnHeaders.Add(3, , "Endereco", 3500)
    
    While Not record.BOF
        Set linhas = lsvClientes.ListItems.Add(1, , record!Nome)
        'If (record!CpfCnpj <> "") Then
        '    linhas.SubItems(1) = record!CpfCnpj
        'End If
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
        
        record_pedido.Close
        
        If (record!Endereco <> "") Then
            linhas.SubItems(2) = record!Endereco
        End If
        
        Controle = Controle + 1
        
        record.MovePrevious
        
        frmPrincipal.Progresso.Value = (((Controle) / record.RecordCount) / 0.001) / 10
        
        Debug.Print frmPrincipal.Progresso.Value
        
        If (Controle = 32700) Then
            response = MsgBox("Atingida capacidade de exibição (32699)", vbInformation, "Parada de processo")
            record.MoveFirst
        End If
    Wend
    
3021:
    frmPrincipal.Progresso.Visible = False
    frmPrincipal.lblProgresso.Visible = False

    cmdTodos.Enabled = False
    record.Close
    
End Sub
Private Sub lsvClientes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim x As Double
      
    For x = Item.Index + 1 To lsvClientes.ListItems.Count Step 1
        lsvClientes.ListItems.Item(x).Checked = False
    Next x
    
    For x = Item.Index - 1 To 1 Step -1
        lsvClientes.ListItems.Item(x).Checked = False
    Next x
        
End Sub
Private Sub lsvClientes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim x As Double
    
    lsvClientes.ListItems.Item(Item.Index).Checked = True
    
    For x = Item.Index + 1 To lsvClientes.ListItems.Count Step 1
        lsvClientes.ListItems.Item(x).Checked = False
    Next x
    
    For x = Item.Index - 1 To 1 Step -1
        lsvClientes.ListItems.Item(x).Checked = False
    Next x
    
End Sub

Private Sub txtBusca_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        cmdBuscar.Value = True
    End If
End Sub
