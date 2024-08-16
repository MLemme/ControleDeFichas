VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   150
   ClientTop       =   975
   ClientWidth     =   9645
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   9645
   Begin MSComctlLib.ProgressBar Progresso 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblProgresso 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Menu mnuInserir 
      Caption         =   "Inserir"
      Begin VB.Menu mnuCliente 
         Caption         =   "Cliente"
      End
   End
   Begin VB.Menu mnuRegistros 
      Caption         =   "Registros"
   End
   Begin VB.Menu mnuImportar 
      Caption         =   "Importar"
   End
   Begin VB.Menu mnuExportar 
      Caption         =   "Exportar"
   End
   Begin VB.Menu mnuSpace1 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSpace2 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "Sobre"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Vers�o As String
Private Sub Form_Load()
    Dim ConectaAcces As String
    Dim ArquivoDB As String
    
    ArquivoDB = "ctrlos.mdb"
        
    ConectaAccess = "Driver={Microsoft Access Driver (*.mdb)};" & _
                "Dbq=" & ArquivoDB & ";" & _
                "DefaultDir=" & App.Path & ";" & _
                "Uid=Admin;Pwd=;"
    
    Set connect_banco = New ADODB.Connection
    
    connect_banco.Open ConectaAccess
    
    Vers�o = "1.1.3"
    
    frmPrincipal.Caption = "Controle de Fichas " & Vers�o
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    connect_banco.Close
    
End Sub
Private Sub mnuCliente_Click()

    frmInserir.Show vbModal
    
End Sub
Private Sub mnuImportar_Click()
    Dim oExcel As New Excel.Application
    Dim SQLComand As String
    Dim CabA1, CabB1, CabC1, CabD1, CabE1, CabF1, CabG1, CabH1 As String
    Dim Conferir, Final As Boolean
    Dim Cont As Double
    Dim ContPedido As Double
    Dim NovoIdCliente As Variant
        
    frmEscolherLocal.Show vbModal
    
    If (ErroDeTipoArquivo = True) Then
            lblProgresso.Visible = False
            Progresso.Visible = False
            Exit Sub
    End If
    'Debug.Print LocalEscolhido
    lblProgresso.Visible = True
    Progresso.Visible = True
        
    lblProgresso.Caption = "Abrindo Planilha..."
        
    Set oExcel = CreateObject("Excel.Application")
    
    Progresso.Value = 5
    
    'oExcel.Workbooks.Open (App.Path & "\Lista.xls")
    oExcel.Workbooks.Open (LocalEscolhido)
    
    Conferir = True
    Final = False
    
    lblProgresso.Caption = "Analisando cabe�alho da Planilha..."
    
    'Conferir Cabe�alho
    CabA1 = oExcel.Application.Cells(1, 1).Value
    CabB1 = oExcel.Application.Cells(1, 2).Value
    CabC1 = oExcel.Application.Cells(1, 3).Value
    CabD1 = oExcel.Application.Cells(1, 4).Value
    CabE1 = oExcel.Application.Cells(1, 5).Value
    CabF1 = oExcel.Application.Cells(1, 6).Value
    CabG1 = oExcel.Application.Cells(1, 7).Value
    CabH1 = oExcel.Application.Cells(1, 8).Value
    
    If (CabA1 <> "Nome") Then
        Conferir = False
    End If
    If (CabB1 <> "A/c") Then
        Conferir = False
    End If
    If (CabC1 <> "Endere�o") Then
        Conferir = False
    End If
    If (CabD1 <> "Bairro") Then
        Conferir = False
    End If
    If (CabE1 <> "Cidade") Then
        Conferir = False
    End If
    If (CabF1 <> "Pedido") Then
        Conferir = False
    End If
    If (CabG1 <> "Telefone") Then
        Conferir = False
    End If
    If (CabH1 <> "Obs") Then
        Conferir = False
    End If
    
    If (Conferir = False) Then
        response = MsgBox("Arquivo com formata��o inv�lida", vbCritical, "Erro")
        
        lblProgresso.Visible = False
        Progresso.Visible = False
        
        oExcel.Workbooks.Close
        oExcel.Quit
        Set oExcel = Nothing
        
        Exit Sub
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
    ContPedido = 0
    Cont = 2
    
    lblProgresso.Caption = "Analisando dados da Planilha..."
    Progresso.Value = 10
        
    While Final = False
        'Cash para tabela Cliente
        ReDim Preserve EntradaDados(Cont - 2)
        EntradaDados(Cont - 2).Id = NovoIdCliente
        EntradaDados(Cont - 2).Nome = oExcel.Application.Cells(Cont, 1).Value
        EntradaDados(Cont - 2).AosCuidados = oExcel.Application.Cells(Cont, 2).Value
        EntradaDados(Cont - 2).Endereco = oExcel.Application.Cells(Cont, 3).Value
        EntradaDados(Cont - 2).Bairro = oExcel.Application.Cells(Cont, 4).Value
        EntradaDados(Cont - 2).Cidade = oExcel.Application.Cells(Cont, 5).Value
        EntradaDados(Cont - 2).Obs = oExcel.Application.Cells(Cont, 8).Value
        
        If (oExcel.Application.Cells(Cont, 7).Value <> "") Then
            Telefone = Split(oExcel.Application.Cells(Cont, 7).Value, "/")
        
            If (UBound(Telefone) = 1) Then
                EntradaDados(Cont - 2).Telefone1 = Telefone(0)
                EntradaDados(Cont - 2).Telefone2 = Telefone(1)
            Else
                EntradaDados(Cont - 2).Telefone1 = Telefone(0)
            End If
            
        Else
        
            EntradaDados(Cont - 2).Telefone1 = ""
            
        End If
        
        'Cash para Tabela pedidos
         
        Pedido = Split(oExcel.Application.Cells(Cont, 6).Value, "/")
        
        If (UBound(Pedido) <> 0) Then
            For i = LBound(Pedido) To UBound(Pedido)
                ReDim Preserve EntradaPedido(ContPedido + 1)
                EntradaPedido(ContPedido).Id = NovoIdCliente
                EntradaPedido(ContPedido).Pedido = Pedido(i)
                ContPedido = ContPedido + 1
            Next
        Else
            ReDim Preserve EntradaPedido(ContPedido + 1)
            EntradaPedido(ContPedido).Id = NovoIdCliente
            EntradaPedido(ContPedido).Pedido = Pedido(0)
            ContPedido = ContPedido + 1
        End If
        
        
        'Fecha loop e redimensiona caches sen�o conta pr�ximo item
        If (EntradaDados(Cont - 2).Nome = "") Then
            If (oExcel.Application.Cells(Cont + 1, 1).Value = "") Then
                Final = True
                ReDim Preserve EntradaDados(UBound(EntradaDados) - 1)
                ReDim Preserve EntradaPedido(UBound(EntradaPedido) - 1)
                
                lblProgresso.Caption = "Transferindo cache de dados da Planilha para o Banco de Dados..."
                Fun��esDoBanco.EsvaziarCaches
                
            Else
                response = MsgBox("Tabela de Importa��o com dados inv�lidos, existe um espa�amento entre dados da coluna Nome", vbCritical, "Erro")
                lblProgresso.Visible = False
                Progresso.Visible = False
                oExcel.Workbooks.Close
                oExcel.Quit
                Set oExcel = Nothing
                
                ReDim Preserve EntradaDados(0)
                ReDim Preserve EntradaPedido(0)
                
                Exit Sub
            End If
        Else
            Cont = Cont + 1
            NovoIdCliente = NovoIdCliente + 1
        End If
        
    Wend
    
    oExcel.Workbooks.Close
    oExcel.Quit
    Set oExcel = Nothing
        
End Sub
Private Sub mnuRegistros_Click()
    
    frmClientes.Show vbModal
    
End Sub
Private Sub mnuSobre_Click()
    
    response = MsgBox("Desenvolvido por: " & Chr(13) & "Mauricio Andrade Lemme" & Chr(13) & "Contato: mauricio.lemme@gmail.com", vbInformation, "Controle de Fichas " & Vers�o)
    
End Sub
Private Sub mnuExportar_Click()
    Dim oExcel As New Excel.Application
    Dim Negrito As Integer
    Dim x As Double
    Dim Celula As String
    Dim SQLComand As String
    Dim Primeiro As Boolean
    
    'Apagar Arquivo Existente
    'response = Shell("cmd /c del teste123.xls")
    
    frmEscolherLocal.Show vbModal
    
    If (ErroDeTipoArquivo = True) Then
            lblProgresso.Visible = False
            Progresso.Visible = False
            Exit Sub
    End If
    
    lblProgresso.Caption = "Exportando dados do Banco de Dados para Planilha..."
    lblProgresso.Visible = True
    Progresso.Visible = True
    
    Set oExcel = CreateObject("Excel.Application")
    oExcel.Workbooks.Add   'inclui o workbook
    Set objExlSht = oExcel.ActiveWorkbook.Sheets(1)
    
    'Escrever Cabe�alho
    oExcel.Range("A1").FormulaR1C1 = "Nome"
    oExcel.Application.Cells(1, 1).ColumnWidth = 50
    oExcel.Range("B1").FormulaR1C1 = "A/c"
    oExcel.Application.Cells(1, 2).ColumnWidth = 15
    oExcel.Range("C1").FormulaR1C1 = "Endere�o"
    oExcel.Application.Cells(1, 3).ColumnWidth = 45
    oExcel.Range("D1").FormulaR1C1 = "Bairro"
    oExcel.Application.Cells(1, 4).ColumnWidth = 15
    oExcel.Range("E1").FormulaR1C1 = "Cidade"
    oExcel.Application.Cells(1, 5).ColumnWidth = 15
    oExcel.Range("F1").FormulaR1C1 = "Pedido"
    oExcel.Application.Cells(1, 6).ColumnWidth = 12
    oExcel.Range("G1").FormulaR1C1 = "Telefone"
    oExcel.Application.Cells(1, 7).ColumnWidth = 20
    oExcel.Range("H1").FormulaR1C1 = "Obs"
    oExcel.Application.Cells(1, 8).ColumnWidth = 50
    
    For Negrito = 1 To 8
        
        oExcel.Application.Cells(1, Negrito).Font.Bold = True
        oExcel.Application.Cells(1, Negrito).Font.Size = 12
        
    Next Negrito
    
    'Escrever Conte�do
    SQLComand = "SELECT * FROM Cliente ORDER BY nome"
    
    Set record = CreateObject("ADODB.Recordset")
    record.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
    record.MoveFirst
    
    x = 2
    
    While Not record.EOF
    
        Celula = "A" & Right(Str(x), Len(Str(x)) - 1)
        oExcel.Range(Celula).Formula = record!Nome
    
        Celula = "B" & Right(Str(x), Len(Str(x)) - 1)
        oExcel.Range(Celula).FormulaR1C1 = record!AosCuidados
        
        Celula = "C" & Right(Str(x), Len(Str(x)) - 1)
        oExcel.Range(Celula).FormulaR1C1 = record!Endereco
        
        Celula = "D" & Right(Str(x), Len(Str(x)) - 1)
        oExcel.Range(Celula).FormulaR1C1 = record!Bairro
        
        Celula = "E" & Right(Str(x), Len(Str(x)) - 1)
        oExcel.Range(Celula).FormulaR1C1 = record!Cidade
        
        Celula = "H" & Right(Str(x), Len(Str(x)) - 1)
        oExcel.Range(Celula).FormulaR1C1 = record!Observacoes
        
        If (record!tel2 <> "") Then
            If (record!tel2 <> "N�o Informado") Then
                Celula = "G" & Right(Str(x), Len(Str(x)) - 1)
                oExcel.Range(Celula).FormulaR1C1 = record!tel1 & " / " & record!tel2
            Else
                Celula = "G" & Right(Str(x), Len(Str(x)) - 1)
                oExcel.Range(Celula).FormulaR1C1 = record!tel1
            End If
        Else
            Celula = "G" & Right(Str(x), Len(Str(x)) - 1)
            oExcel.Range(Celula).FormulaR1C1 = record!tel1
        End If
        
        'F pedido
        SQLComand = "SELECT numeropedido FROM Pedidos WHERE idCliente = " & record!idCliente
        
        Set record_pedido = CreateObject("ADODB.Recordset")
        record_pedido.Open SQLComand, connect_banco, adOpenKeyset, adLockOptimistic
        If record_pedido.RecordCount <> 0 Then 'Corre��o 14.01.2010
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
                
                Celula = "F" & Right(Str(x), Len(Str(x)) - 1)
                oExcel.Range(Celula).FormulaR1C1 = Pedido
                
            Else
            
                Celula = "F" & Right(Str(x), Len(Str(x)) - 1)
                oExcel.Range(Celula).FormulaR1C1 = record_pedido!numeropedido
                
            End If
        Else
            Celula = "F" & Right(Str(x), Len(Str(x)) - 1)
            oExcel.Range(Celula).FormulaR1C1 = "N�o h� PEDIDOS"
        End If
        
        'frmPrincipal.Progresso.Value = Int(((x - 2) * 100) / record.RecordCount)
        frmPrincipal.Progresso.Value = (((x - 2) / record.RecordCount) / 0.001) / 10
        x = x + 1
        record.MoveNext
        
    Wend
         
    record.Close
    
    'objExlSht.SaveAs (App.Path & "\Lista.xls")
    Debug.Print LocalEscolhido
    
    objExlSht.SaveAs (LocalEscolhido)
   
    oExcel.ActiveWorkbook.Saved = True
    oExcel.Workbooks.Close
    oExcel.Quit
    
    Progresso.Visible = False
    lblProgresso.Visible = False
    
    Set oExcel = Nothing
    Set objExlSht = Nothing

End Sub
