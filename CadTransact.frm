VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Sistema de Transações de Cartão"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Casdastro de Transações"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   1
      Top             =   150
      Width           =   8940
      Begin VB.CommandButton btnSair 
         Caption         =   "Sair"
         Height          =   510
         Left            =   6060
         TabIndex        =   15
         Top             =   1545
         Width           =   915
      End
      Begin VB.CommandButton btnEditar 
         Caption         =   "Editar"
         Height          =   510
         Left            =   3372
         TabIndex        =   14
         Top             =   1545
         Width           =   615
      End
      Begin VB.CommandButton btnExcluir 
         Caption         =   "Excluir"
         Height          =   510
         Left            =   2661
         TabIndex        =   13
         Top             =   1545
         Width           =   630
      End
      Begin VB.CommandButton btnSalvar 
         Caption         =   "Salvar"
         Height          =   510
         Left            =   1965
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1545
         Width           =   615
      End
      Begin VB.CommandButton btnBuscar 
         Caption         =   "Pesquisar"
         Height          =   510
         Left            =   4068
         TabIndex        =   11
         Top             =   1545
         Width           =   915
      End
      Begin VB.CommandButton btnRelatorio 
         Caption         =   "Relatório"
         Height          =   510
         Left            =   5064
         TabIndex        =   10
         Top             =   1545
         Width           =   915
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6540
         Top             =   315
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtDescricao 
         Height          =   525
         Left            =   4275
         TabIndex        =   8
         Top             =   900
         Width           =   4365
      End
      Begin VB.TextBox txtData 
         Height          =   375
         Left            =   1395
         TabIndex        =   6
         Top             =   900
         Width           =   1590
      End
      Begin VB.TextBox txtValor 
         Height          =   375
         Left            =   4275
         TabIndex        =   4
         Top             =   495
         Width           =   1590
      End
      Begin VB.TextBox txtCartao 
         Height          =   375
         Left            =   1395
         TabIndex        =   3
         Top             =   495
         Width           =   1590
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "DESCRIÇÃO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3030
         TabIndex        =   9
         Top             =   900
         Width           =   1155
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "DATA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   7
         Top             =   900
         Width           =   1260
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "VALOR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2895
         TabIndex        =   5
         Top             =   495
         Width           =   1260
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Nº CARTÃO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   180
         TabIndex        =   2
         Top             =   495
         Width           =   1260
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGridTransacoes 
      Height          =   2370
      Left            =   360
      TabIndex        =   0
      Top             =   2805
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   4180
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Lista de Transações"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================
' Váriaveis Globais
'========================================
Option Explicit
    
    Dim conn As ADODB.Connection
    Dim IdSelecionado As Long
    Dim EmEdicao As Boolean

'========================================
' Início Sistema e conexão ao banco
'========================================
Private Sub Form_Load()

   Set conn = New ADODB.Connection
   conn.Open "Provider=SQLOLEDB;Data Source=.\SQLEXPRESS;Initial Catalog=XYZAdmCardDB;Integrated Security=SSPI"
  
   btnExcluir.Enabled = False
   btnEditar.Enabled = False

   Me.WindowState = vbMaximized
   
End Sub
'========================================
' Formata automaticamente a data digitada
'========================================
Private Sub txtData_LostFocus()
If Trim(txtData.Text) <> "" Then
    If IsDate(txtData.Text) Then
        txtData.Text = Format(CDate(txtData.Text), "dd/mm/yyyy")
    Else
        MsgBox "Data inválida!", vbExclamation
        txtData.Text = ""
        txtData.SetFocus
    End If
End If
End Sub
'========================================
' Validação de entrada do campo Valor
' Permite apenas:
' - Números (0-9)
' - Vírgula (separador decimal)
' - Backspace (correção)
' Bloqueia qualquer outro caractere
'========================================
Private Sub txtValor_KeyPress(KeyAscii As Integer)

    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub
    If KeyAscii = 44 Then Exit Sub
    If KeyAscii = 8 Then Exit Sub
    KeyAscii = 0
    
End Sub

'========================================
' CONSULTA
' Realiza busca de transações com filtros opcionais
'========================================
Private Sub btnBuscar_Click()

    On Error GoTo TrataErro

    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim sql As String

    sql = "SELECT IdTransacao, NumeroCartao, ValorTransacao, DataTransacao, Descricao " & _
          "FROM Transacoes WHERE 1=1"

    ' Filtro por cartão
    If Trim(txtCartao.Text) <> "" Then
        sql = sql & " AND NumeroCartao = ?"
    End If

    ' Filtro por valor
    If Trim(txtValor.Text) <> "" Then
        If Not IsNumeric(txtValor.Text) Then
            MsgBox "Valor inválido!", vbExclamation
            Exit Sub
        End If
        sql = sql & " AND ValorTransacao = ?"
    End If

    ' Filtro por data
    If Trim(txtData.Text) <> "" Then
        sql = sql & " AND CAST(DataTransacao AS DATE) = ?"
    End If

    ' Filtro por descrição
    If Trim(txtDescricao.Text) <> "" Then
        sql = sql & " AND UPPER(Descricao) LIKE ?"
    End If

    ' Ordenação
    sql = sql & " ORDER BY DataTransacao"

    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = conn
        .CommandText = sql
        .CommandType = adCmdText

        If Trim(txtCartao.Text) <> "" Then
            .Parameters.Append .CreateParameter("pCartao", adVarChar, adParamInput, 20, txtCartao.Text)
        End If
        If Trim(txtValor.Text) <> "" Then
            .Parameters.Append .CreateParameter("pValor", adDouble, adParamInput, , _
                ConverterValorParaBanco(txtValor.Text))
        End If
        If Trim(txtData.Text) <> "" Then
            .Parameters.Append .CreateParameter("pData", adDate, adParamInput, , _
                CDate(txtData.Text))
        End If
        If Trim(txtDescricao.Text) <> "" Then
            .Parameters.Append .CreateParameter("pDesc", adVarChar, adParamInput, 255, _
                "%" & UCase(txtDescricao.Text) & "%")
        End If
    End With

    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenKeyset, adLockReadOnly


    If rs.EOF Then
        MsgBox "A consulta não retornou nenhuma transação!", vbInformation
        Set grdDataGridTransacoes.DataSource = Nothing
        Exit Sub
    End If

    Set grdDataGridTransacoes.DataSource = Nothing
    Set grdDataGridTransacoes.DataSource = rs

    Call FormatarGrid

    btnExcluir.Enabled = True
    btnEditar.Enabled = True

    Exit Sub

TrataErro:
    MsgBox "Erro ao buscar dados: " & Err.Description, vbCritical

End Sub

'========================================
' Carrega os dados do registro selecionado
' Regras:
' - Obtém o registro atual do DataGrid
' - Preenche os campos para visualização/edição
' - Formata valor e data para exibição
' - Habilita botões de edição e exclusão
'========================================

Private Sub grdDataGridTransacoes_DblClick()

    Dim rs As ADODB.Recordset

    If grdDataGridTransacoes.DataSource Is Nothing Then Exit Sub

    Set rs = grdDataGridTransacoes.DataSource

    If rs.EOF Then Exit Sub

    IdSelecionado = rs!IdTransacao

    txtCartao.Text = rs!NumeroCartao

    If Not IsNull(rs!ValorTransacao) Then
        txtValor.Text = Replace(Format(rs!ValorTransacao, "0.00"), ".", ",")
    End If

    If Not IsNull(rs!DataTransacao) Then
        txtData.Text = Format(rs!DataTransacao, "dd/mm/yyyy")
    End If

    txtDescricao.Text = rs!Descricao

    btnExcluir.Enabled = True
    btnEditar.Enabled = True
End Sub

'========================================
' INSERÇÃO
' Insere uma nova transação na base de dados
' Valida duplicidade antes da gravação
'========================================

Private Sub btnSalvar_Click()

   On Error GoTo Erro
   
   Dim cmd As ADODB.Command
   Dim rs As ADODB.Recordset
   Dim valor As Double
   Dim cmdCliente As ADODB.Command
   Dim rsCliente As ADODB.Recordset
   Dim idCliente As Long
   
   Set cmdCliente = New ADODB.Command
   With cmdCliente
       .ActiveConnection = conn
       .CommandText = "SELECT IdCliente FROM Clientes WHERE NumeroCartao = ?"
       .CommandType = adCmdText
       .Parameters.Append .CreateParameter("pCartao", adVarChar, adParamInput, 20, txtCartao.Text)
   End With

   Set rsCliente = New ADODB.Recordset
   rsCliente.CursorLocation = adUseClient
   rsCliente.Open cmdCliente, , adOpenKeyset, adLockReadOnly
   
   If rsCliente.EOF Then
       MsgBox "Cartão não encontrado!", vbExclamation
       Exit Sub
   End If
   
   idCliente = rsCliente!idCliente
   
   rsCliente.Close
   Set rsCliente = Nothing

    If Trim(txtCartao.Text) = "" Then
        MsgBox "Informe um cartão válido!"
        Exit Sub
    End If

    If Trim(txtValor.Text) = "" Or Not IsNumeric(txtValor.Text) Then
        MsgBox "Informe o valor!"
        Exit Sub
    End If

    If Trim(txtDescricao.Text) = "" Then
        MsgBox "Informe a descrição!"
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
            
    Dim cmdCount As ADODB.Command
    Set cmdCount = New ADODB.Command
    With cmdCount
        .ActiveConnection = conn
        .CommandText = "SELECT COUNT(*) AS Total FROM Transacoes WHERE NumeroCartao = ? AND ValorTransacao = ? AND Descricao = ?"
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("pCartao", adVarChar, adParamInput, 20, txtCartao.Text)
        .Parameters.Append .CreateParameter("pValor", adDouble, adParamInput, , ConverterValorParaBanco(txtValor.Text))
        .Parameters.Append .CreateParameter("pDesc", adVarChar, adParamInput, 255, txtDescricao.Text)
    End With
            
    rs.CursorLocation = adUseClient
    rs.Open cmdCount, , adOpenKeyset, adLockReadOnly
           
    If rs.Fields("Total").Value > 0 Then
        MsgBox "Registro já existe!"
        rs.Close

        Exit Sub
    End If

    rs.Close
    
    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = conn
        .CommandText = "INSERT INTO Transacoes (NumeroCartao, IdCliente, ValorTransacao, DataTransacao, Descricao) VALUES (?, ?, CONVERT(money, ?, 2), GETDATE(), ?)"
        .CommandType = adCmdText

        .Parameters.Append .CreateParameter("pCartao", adVarChar, adParamInput, 50, txtCartao.Text)
        .Parameters.Append .CreateParameter("pIdCliente", adInteger, adParamInput, , idCliente)
        .Parameters.Append .CreateParameter("pValor", adVarChar, adParamInput, 50, _
         Replace(Replace(txtValor.Text, ".", ""), ",", "."))
        .Parameters.Append .CreateParameter("pDescricao", adVarChar, adParamInput, 255, txtDescricao.Text)
        
        .Execute
    End With

   Call LimparCampos

   MsgBox "Registro Salvo!"
    
   Call CarregarGrid
    Exit Sub

Erro:
    MsgBox "Erro ao salvar: " & Err.Description

End Sub
'========================================
' ATUALIZAÇÃO
' Atualiza um registro existente
'========================================
Private Sub btnEditar_Click()

   EmEdicao = True
    On Error GoTo Erro

    Dim cmd As ADODB.Command
   Dim valor As Double

    If IdSelecionado = 0 Then
        MsgBox "Selecione um registro!"
        Exit Sub
    End If

    If Trim(txtCartao.Text) = "" Then
        MsgBox "Informe um cartão!"
        Exit Sub
    End If

    If Trim(txtValor.Text) = "" Or Not IsNumeric(txtValor.Text) Then
        MsgBox "Valor inválido!"
        Exit Sub
    End If

    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = conn
        .CommandText = "UPDATE Transacoes SET NumeroCartao = ?, ValorTransacao = CONVERT(money, ?, 2), Descricao = ? WHERE IdTransacao = ?"
        .CommandType = adCmdText

        .Parameters.Append .CreateParameter("pCartao", adVarChar, adParamInput, 50, txtCartao.Text)
        .Parameters.Append .CreateParameter("pValor", adVarChar, adParamInput, 50, _
         Replace(Replace(txtValor.Text, ".", ""), ",", "."))
        
        .Parameters.Append .CreateParameter("pDescricao", adVarChar, adParamInput, 255, txtDescricao.Text)
        .Parameters.Append .CreateParameter("pIdTransacao", adInteger, adParamInput, , IdSelecionado)

        .Execute
    End With

    MsgBox "Registro Atualizado!"

   Set grdDataGridTransacoes.DataSource = Nothing
   EmEdicao = False
   Call LimparCampos
   Call CarregarGrid

    Exit Sub

Erro:
    MsgBox "Erro ao atualizar: " & Err.Description

End Sub
'========================================
' EXCLUSÃO
' Remoção de um registro
'========================================
Private Sub btnExcluir_Click()

    On Error GoTo Erro

    Dim cmd As ADODB.Command

    If IdSelecionado = 0 Then
        MsgBox "Selecione um registro!"
        Exit Sub
    End If

    If MsgBox("Deseja Realmente Excluir?", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    Set cmd = New ADODB.Command

    With cmd
        .ActiveConnection = conn
        .CommandText = "DELETE FROM Transacoes WHERE IdTransacao = ?"
        .CommandType = adCmdText

        .Parameters.Append .CreateParameter("pId", adInteger, adParamInput, , IdSelecionado)

        .Execute
    End With

    MsgBox "Registro Excluído!"

    Call LimparCampos
    Call CarregarGrid
   
    Exit Sub

Erro:
    MsgBox "Erro ao excluir: " & Err.Description
    
        
End Sub


'========================================
' CARREGAMENTO DE DADOS / GRID
'========================================
Private Sub CarregarGrid()

   EmEdicao = True

   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset

   rs.CursorLocation = adUseClient

   rs.Open "SELECT IdTransacao, NumeroCartao, ValorTransacao, DataTransacao, Descricao FROM Transacoes ORDER BY DataTransacao", _
           conn, adOpenKeyset, adLockReadOnly

   Set grdDataGridTransacoes.DataSource = Nothing
   Set grdDataGridTransacoes.DataSource = rs
   
   Call FormatarGrid

   EmEdicao = False

End Sub
'========================================
' FORMATAÇÃO VISUAL DO GRID
'========================================
Private Sub FormatarGrid()
   
   Dim i As Integer
  
   With grdDataGridTransacoes
   
      .Columns(0).Caption = "ID"
      .Columns(1).Caption = "Cartão"
      .Columns(2).Caption = "Valor (R$)"
      .Columns(3).Caption = "Data"
      .Columns(4).Caption = "Descrição"
      .Columns(2).Alignment = dbgRight
      .Columns(3).Alignment = dbgCenter
      .Columns(3).NumberFormat = "dd/mm/yyyy"
      .Columns(2).Locked = True
      
   End With
   grdDataGridTransacoes.AllowUpdate = False
End Sub

'========================================
' FUNÇÃO AUXILIAR - CONVERSÃO DE VALORES
'========================================
Private Function ConverterValorParaBanco(valorTexto As String) As Double

    valorTexto = Trim(valorTexto)

    If valorTexto = "" Then
        ConverterValorParaBanco = 0
        Exit Function
    End If

    valorTexto = Replace(valorTexto, ".", "")
    valorTexto = Replace(valorTexto, ",", ".")

    If Not IsNumeric(valorTexto) Then
        MsgBox "Valor inválido!"
        ConverterValorParaBanco = 0
        Exit Function
    End If

    ConverterValorParaBanco = Round(CDbl(valorTexto), 2)

End Function
'========================================
' RELATÓRIO EXCEL
'========================================
Private Sub btnRelatorio_Click()
   Dim xlApp As Object, xlBook As Object, xlSheet As Object
   Dim rs As ADODB.Recordset
   Dim mesSelecionado As Integer
   Dim anoSelecionado As Integer
   Dim dataIni As Date
   Dim dataFim As Date
   Dim i As Long
   Dim caminhoArquivo As String
   Dim resposta As String
    
      resposta = InputBox("Informe o mês (1 a 12):", "Relatório por Mês", Month(Date))
      
      If Trim(resposta) = "" Then
         MsgBox "Relatório Abortado pelo Usuário!", vbExclamation
      Exit Sub
      End If
      
      mesSelecionado = Val(resposta)
      
      If mesSelecionado < 1 Or mesSelecionado > 12 Then
          MsgBox "Mês inválido!", vbExclamation
          Exit Sub
      End If
      
      anoSelecionado = Year(Date)
      dataIni = DateSerial(anoSelecionado, mesSelecionado, 1)
      dataFim = DateAdd("m", 1, dataIni)
      
   '=======================================================
   'Definindo local arquivo a ser salvo
   '=======================================================
    With CommonDialog1
        .CancelError = True
        On Error GoTo UsuarioCancelou
        .Filter = "Arquivos Excel (*.xlsx)|*.xlsx"
        .DialogTitle = "Salvar Relatório"
        .ShowSave
        caminhoArquivo = .FileName
    End With
   '=======================================================
   
    Set rs = New ADODB.Recordset
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic
    rs.LockType = adLockReadOnly
    
    Dim cmdRel As ADODB.Command
    Set cmdRel = New ADODB.Command
    With cmdRel
        .ActiveConnection = conn
        .CommandText = "SELECT t.NumeroCartao, t.ValorTransacao, t.DataTransacao, t.Descricao, dbo.fn_CATegoriaTransacoes(t.ValorTransacao) AS Categoria FROM Transacoes t WHERE t.DataTransacao >= ? AND t.DataTransacao < ? ORDER BY t.DataTransacao"
        .CommandType = adCmdText
        .Parameters.Append .CreateParameter("pDataIni", adDate, adParamInput, , dataIni)
        .Parameters.Append .CreateParameter("pDataFim", adDate, adParamInput, , dataFim)
    End With
    
    rs.Open cmdRel, , adOpenKeyset, adLockReadOnly
   
      Set xlApp = CreateObject("Excel.Application")
      xlApp.Visible = True
      Set xlBook = xlApp.Workbooks.Add
      Set xlSheet = xlBook.Sheets(1)
      
      
      xlSheet.Cells(1, 1).Value = "Nº Cartão"
      xlSheet.Cells(1, 2).Value = "Valor R$"
      xlSheet.Cells(1, 3).Value = "Data"
      xlSheet.Cells(1, 4).Value = "Descrição"
      xlSheet.Cells(1, 5).Value = "Categoria"
   
   If rs.EOF Then
       MsgBox "Nenhum registro encontrado!", vbExclamation
       Exit Sub
   End If

   i = 2
   Do While Not rs.EOF
     
      xlSheet.Cells(i, 1).NumberFormat = "@"
      xlSheet.Cells(i, 1).Value = "'" & Format(rs.Fields("NumeroCartao").Value, "0000 0000 0000 0000")
      
      xlSheet.Cells(i, 2).Value = rs.Fields("ValorTransacao").Value
      xlSheet.Cells(i, 3).Value = rs.Fields("DataTransacao").Value
      xlSheet.Cells(i, 4).Value = rs.Fields("Descricao").Value
      xlSheet.Cells(i, 5).Value = rs.Fields("Categoria").Value
        
        i = i + 1
        rs.MoveNext
    Loop
    
   '=======================================================
   'Formatação do Relatório
   '=======================================================
   
   Dim ultimaLinha As Long
   ultimaLinha = i - 1
   
   With xlSheet
    
       .Range("A1:E1").Font.Bold = True
       .Range("A1:E1").Interior.Color = RGB(200, 200, 200)
       .Range("A1:E1").HorizontalAlignment = -4108
       .Range("A1:E" & ultimaLinha).Borders.LineStyle = 1
       .Columns(1).HorizontalAlignment = -4108
       .Columns(2).HorizontalAlignment = -4152
       .Columns(3).HorizontalAlignment = -4108
       .Columns(5).HorizontalAlignment = -4108
       .Columns(2).NumberFormat = "#,##0.00"
       .Columns(3).NumberFormat = "dd/mm/yyyy"
       .Columns("A:E").AutoFit
   End With
   '=======================================================
    
    xlBook.SaveAs caminhoArquivo

    
    rs.Close
    Set rs = Nothing
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing

    MsgBox "Relatório gerado com sucesso!", vbInformation
    Exit Sub

UsuarioCancelou:
    MsgBox "Exportação cancelada pelo usuário.", vbExclamation
End Sub
'========================================
' LIMPEZA DE CAMPOS
'========================================
Private Sub LimparCampos()

   txtCartao.Text = ""
   txtValor.Text = ""
   txtDescricao.Text = ""
   txtData.Text = ""

   IdSelecionado = 0

   btnExcluir.Enabled = False
   btnEditar.Enabled = False

End Sub

Private Sub btnSair_Click()
'========================================
' Saindo do sistema
'========================================
    If MsgBox("Deseja realmente sair?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then Exit Sub

    Unload Me
End Sub

'========================================
' Encerramento do Sistema
' Fechando conexão com banco
'========================================
Private Sub Form_Unload(Cancel As Integer)

    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If

End Sub


