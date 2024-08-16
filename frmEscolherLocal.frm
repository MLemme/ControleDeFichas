VERSION 5.00
Begin VB.Form frmEscolherLocal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Escolha o Arquivo"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "Finalizar"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtLocal 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
   End
   Begin VB.FileListBox filArquivo 
      Height          =   1650
      Left            =   2520
      Pattern         =   "*.xls"
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.DirListBox dirPasta 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.DriveListBox drvDisco 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   2175
   End
End
Attribute VB_Name = "frmEscolherLocal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    ErroDeTipoArquivo = True
    Me.Hide
End Sub
Private Sub cmdFinalizar_Click()
    Dim Avalia As String
    
    LocalEscolhido = txtLocal.Text
    Avalia = Right(LocalEscolhido, 3)
    If (Avalia <> "xls") Then
        response = MsgBox("Tipo de Arquivo Inválido!", vbCritical, "Erro")
        ErroDeTipoArquivo = True
    Else
        ErroDeTipoArquivo = False
    End If
    
    Me.Hide
End Sub
Private Sub dirPasta_Change()
    filArquivo.Path = dirPasta.Path
    If (filArquivo.ListCount = 0) Then
        txtLocal.Text = dirPasta.Path & "\Lista.xls"
    Else
        txtLocal.Text = dirPasta.Path
    End If
End Sub
Private Sub drvDisco_Change()
    On Error GoTo 68
    dirPasta.Path = drvDisco.Drive
68:
    If Err.Number = 68 Then
        resposta = MsgBox("Falha ao Abrir Drive", vbCritical, "Erro")
    End If
    Exit Sub
End Sub
Private Sub filArquivo_Click()
    txtLocal.Text = filArquivo.Path + "\" + filArquivo.FileName
End Sub
Private Sub Form_Load()
    txtLocal.Text = App.Path & "\Lista.xls"
End Sub
