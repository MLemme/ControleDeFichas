Attribute VB_Name = "Global"
Global record As ADODB.Recordset
Global record_pedido As ADODB.Recordset
Global connect_banco As ADODB.Connection
Global flag As Boolean
Global PublicIdCliente As String
Global LocalEscolhido As String
Global ErroDeTipoArquivo As Boolean
Global EntradaDados() As ImportarDados
Global EntradaPedido() As ImportarPedido

Type ImportarDados
    Id As String
    Nome As String
    AosCuidados As String
    Endereco As String
    Bairro As String
    Cidade As String
    Telefone1 As String
    Telefone2 As String
    Obs As String
End Type

Type ImportarPedido
    Id As String
    Pedido As String
End Type
