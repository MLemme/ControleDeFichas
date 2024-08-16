Attribute VB_Name = "FunçõesDoBanco"
Public Function InserirCliente(ByVal Nome As String, ByVal AosCuidados As String, ByVal Endereco As String, ByVal Bairro As String, _
ByVal Cidade As String, ByVal CpfCnpj As String, ByVal Telefone1 As String, ByVal Telefone2 As String, ByVal Observacoes As String, _
ByVal NrPedido1 As String, ByVal NrPedido2 As String, ByVal NrPedido3 As String) As Boolean

    Dim SQLComand As String
    Dim SQLValues As String
    Dim NovoIdCliente As Double
    Dim MaxRecords As Variant
    
    InserirCliente = False
    
    ' Nome
    If (Nome = "" Or Nome = Null Or Nome = Empty) Then
        response = MsgBox("O campo Nome não foi preenchido", vbCritical, "Erro 01")
        Exit Function
    End If
    ' Endereco
    If (Endereco = "" Or Endereco = Null Or Endereco = Empty) Then
        response = MsgBox("O campo Endereço não foi preenchido", vbCritical, "Erro 01")
        Exit Function
    End If
    ' Bairro
    If (Bairro = "" Or Bairro = Null Or Bairro = Empty) Then
        response = MsgBox("O campo Bairro não foi preenchido", vbCritical, "Erro 01")
        Exit Function
    End If
    ' Cidade
    If (Cidade = "" Or Cidade = Null Or Cidade = Empty) Then
        response = MsgBox("O campo Cidade não foi preenchido", vbCritical, "Erro 01")
        Exit Function
    End If
    
    SQLComand = "SELECT nome,endereco,bairro,cidade FROM Cliente WHERE nome = '" & Nome & "' AND endereco = '" & Endereco & "' AND cidade = '" & Cidade & "' AND bairro = '" & Bairro & "'"
    
    Set record = CreateObject("ADODB.Recordset")
    record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
    MaxRecords = record.RecordCount
    If (MaxRecords <> 0) Then
        response = MsgBox("O campo Endereço, Bairro, Cidade e Nome já existentem para o mesmo registro", vbCritical, "Erro!")
        record.Close
        Exit Function
    End If
    record.Close
    
    ' AosCuidados
    If (AosCuidados = "" Or AosCuidados = Null Or AosCuidados = Empty) Then
        response = MsgBox("O campo Aos Cuidados(A/c) não foi preenchido", vbCritical, "Erro 01")
        Exit Function
    End If
    
    ' CpfCnpj
    If (CpfCnpj = "" Or CpfCnpj = Null Or CpfCnpj = Empty) Then
    '    response = MsgBox("O campo Cpf / Cnpj não foi preenchido", vbCritical, "Erro 01")
    '    Exit Function
        CpfCnpj = "Não Informado"
    End If
    ' Telefone1
    If (Telefone1 = "" Or Telefone1 = Null Or Telefone1 = Empty) Then
    '    response = MsgBox("O campo Telefone 1 não foi preenchido", vbCritical, "Erro 01")
    '    Exit Function
        Telefone1 = "Não Informado"
    End If
    ' Telefone2
    If (Telefone2 = "" Or Telefone2 = Null Or Telefone2 = Empty) Then
    '    response = MsgBox("O campo Telefone 2 não foi preenchido", vbCritical, "Erro 01")
    '    Exit Function
        Telefone2 = "Não Informado"
    End If
    ' Observacoes
    If (Observacoes = "" Or Observacoes = Null Or Observacoes = Empty) Then
    '    response = MsgBox("O campo Observações não foi preenchido", vbCritical, "Erro 01")
    '    Exit Function
        Observacoes = "Não Preenchido"
    End If
    ' NrPedido1
    If (NrPedido1 = "" Or NrPedido1 = Null Or NrPedido1 = Empty) Then
    '    response = MsgBox("O campo NrPedido1 não foi preenchido", vbCritical, "Erro 01")
    '    Exit Function
    End If
    ' NrPedido2
    If (NrPedido2 = "" Or NrPedido2 = Null Or NrPedido2 = Empty) Then
    '    response = MsgBox("O campo NrPedido2 não foi preenchido", vbCritical, "Erro 01")
    '    Exit Function
    End If
    ' NrPedido3
    If (NrPedido3 = "" Or NrPedido3 = Null Or NrPedido3 = Empty) Then
    '    response = MsgBox("O campo NrPedido3 não foi preenchido", vbCritical, "Erro 01")
    '    Exit Function
    End If
    
    SQLComand = "SELECT idCliente FROM Cliente"
    
    Set record = CreateObject("ADODB.Recordset")
    record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
    On Error GoTo 3021
    record.MoveLast
    
    NovoIdCliente = record!idCliente + 1
    
3021:
        If (Err.Number = 3021) Then
            NovoIdCliente = 0
        End If
    record.Close
    
    SQLValues = Str(NovoIdCliente) & ",'" & Nome & "','" & AosCuidados & "','" & Endereco & "','" & Bairro & "','" & Cidade & "','" & CpfCnpj & "','" & Telefone1 & "','" & Telefone2 & "','" & Observacoes
    
    SQLComand = "INSERT INTO Cliente(idCliente,nome,aoscuidados,endereco,bairro,cidade,cpfcnpj,tel1,tel2,observacoes) VALUES (" & SQLValues & "')"

    connect_banco.Execute SQLComand
    
    If (NrPedido1 <> "" Or NrPedido1 <> Null Or NrPedido1 <> Empty) Then
        SQLComand = "INSERT INTO Pedidos(idCliente,numeropedido) VALUES (" & Str(NovoIdCliente) & ",'" & NrPedido1 & "')"
        connect_banco.Execute SQLComand
    End If
    
    If (NrPedido2 <> "" Or NrPedido2 = Null Or NrPedido2 <> Empty) Then
        SQLComand = "INSERT INTO Pedidos(idCliente,numeropedido) VALUES (" & Str(NovoIdCliente) & ",'" & NrPedido2 & "')"
        connect_banco.Execute SQLComand
    End If
    
    If (NrPedido3 <> "" Or NrPedido3 <> Null Or NrPedido3 <> Empty) Then
        SQLComand = "INSERT INTO Pedidos(idCliente,numeropedido) VALUES (" & Str(NovoIdCliente) & ",'" & NrPedido3 & "')"
        connect_banco.Execute SQLComand
    End If
    
    response = MsgBox("Cliente inserido com sucesso", vbInformation, "Sucesso !")
    
    InserirCliente = True
    
End Function
Public Function AlterarCliente(ByVal idCliente As Variant, ByVal Nome As String, ByVal AosCuidados As String, ByVal Endereco As String, ByVal Bairro As String, _
ByVal Cidade As String, ByVal CpfCnpj As String, ByVal Telefone1 As String, ByVal Telefone2 As String, ByVal Observacoes As String) As Boolean
   
    Dim SQLComand As String
    Dim SQLValues As String
    'Dim idcliente As String
    
    AlterarCliente = False
    
    'SQLComand = "SELECT idCliente FROM Cliente WHERE nome = '" & Nome & "' AND endereco = '" & Endereco & "'"
    
    'Set record = CreateObject("ADODB.Recordset")
    'record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
    
    'idcliente = record!idcliente
    
    'record.Close
    
    SQLValues = "nome = '" & Nome & "', aoscuidados = '" & AosCuidados & "', endereco = '" & Endereco & "', bairro = '" & Bairro
    
    SQLValues = SQLValues & "', cidade = '" & Cidade & "', cpfcnpj = '" & CpfCnpj & "', tel1 = '" & Telefone1 & "', tel2 = '" & Telefone2 & "', observacoes = '" & Observacoes & "'"
    
    SQLComand = "UPDATE Cliente SET " & SQLValues & " WHERE idCliente = " & idCliente
    
    connect_banco.Execute SQLComand
    
    AlterarCliente = True

End Function
Public Function DeletarCliente(ByVal idCliente As String) As Boolean
    Dim SQLComand As String
    
    DeletarCliente = False
    
    SQLComand = "DELETE FROM Cliente WHERE idCliente = " & idCliente
        
    connect_banco.Execute SQLComand
    
    SQLComand = "DELETE FROM Pedidos WHERE idCliente = " & idCliente
    
    connect_banco.Execute SQLComand
    
    response = MsgBox("Registro apagado com sucesso", vbInformation, "Sucesso !")
    
    DeletarCliente = True

End Function
Public Function EsvaziarCaches() ' As Boolean
    Dim SQLComand As String
    Dim SQLValues As String
    
    For i = LBound(EntradaDados) To UBound(EntradaDados)
        
         
        SQLValues = EntradaDados(i).Id & ",'" & EntradaDados(i).Nome & "','" & EntradaDados(i).AosCuidados & "','" & EntradaDados(i).Endereco & "','" & EntradaDados(i).Bairro & "','" & EntradaDados(i).Cidade & "','" & EntradaDados(i).Telefone1 & "','" & EntradaDados(i).Telefone2 & "','" & EntradaDados(i).Obs
    
        SQLComand = "INSERT INTO Cliente(idCliente,nome,aoscuidados,endereco,bairro,cidade,tel1,tel2,observacoes) VALUES (" & SQLValues & "')"

        connect_banco.Execute SQLComand
        
        frmPrincipal.Progresso.Value = (i * 100) / (UBound(EntradaDados) + UBound(EntradaPedido))
        
    Next i
    
    For i = LBound(EntradaPedido) To UBound(EntradaPedido)
        
        SQLComand = "INSERT INTO Pedidos(idCliente,numeropedido) VALUES (" & EntradaPedido(i).Id & ",'" & EntradaPedido(i).Pedido & "')"
        
        connect_banco.Execute SQLComand
        
        frmPrincipal.Progresso.Value = ((i + UBound(EntradaDados)) * 100) / (UBound(EntradaDados) + UBound(EntradaPedido))
    
    Next i
    
    ReDim Preserve EntradaDados(0)
    ReDim Preserve EntradaPedido(0)
    
    frmPrincipal.lblProgresso.Visible = False
    frmPrincipal.Progresso.Visible = False
    
End Function
'shell (cmd /c del chr(34) xpto chr(34))
