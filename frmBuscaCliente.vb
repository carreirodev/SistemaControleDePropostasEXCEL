' Código para o formulário frmBuscaCliente

Private Sub UserForm_Initialize()
    ' Inicializa o formulário
    ConfigurarListBox
End Sub

Private Sub ConfigurarListBox()
    ' Configura as colunas da ListBox sem adicionar cabeçalho
    With lstResultados
        .Clear
        .ColumnCount = 5
        .ColumnWidths = "44;150;130;100;22"
    End With
End Sub

Private Sub btnBuscar_Click()
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim encontrou As Boolean
    
    Set ws = ThisWorkbook.Sheets("CLIENTES")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Limpa a ListBox antes de uma nova busca
    lstResultados.Clear
    
    encontrou = False
    
    For i = 2 To ultimaLinha ' Assumindo que a linha 1 é o cabeçalho na planilha
        If (InStr(1, ws.Cells(i, 1).Value, txtID.Value, vbTextCompare) > 0 And Len(txtID.Value) > 0) Or _
           (InStr(1, ws.Cells(i, 2).Value, txtNomeCliente.Value, vbTextCompare) > 0 And Len(txtNomeCliente.Value) > 0) Then
            
            lstResultados.AddItem
            lstResultados.List(lstResultados.ListCount - 1, 0) = ws.Cells(i, 1).Value ' ID
            lstResultados.List(lstResultados.ListCount - 1, 1) = ws.Cells(i, 2).Value ' Nome do Cliente
            lstResultados.List(lstResultados.ListCount - 1, 2) = ws.Cells(i, 3).Value ' Pessoa de Contato
            lstResultados.List(lstResultados.ListCount - 1, 3) = ws.Cells(i, 5).Value ' Cidade
            lstResultados.List(lstResultados.ListCount - 1, 4) = ws.Cells(i, 6).Value ' Estado
            
            encontrou = True
        End If
    Next i
    
    If Not encontrou Then
        MsgBox "Nenhum cliente encontrado com os critérios fornecidos.", vbInformation
    End If
End Sub

Private Sub lstResultados_Click()
    If lstResultados.ListIndex >= 0 Then ' Alterado para >= 0 já que não há mais linha de cabeçalho
        PreencherCamposCliente lstResultados.List(lstResultados.ListIndex, 0) ' Passa o ID do cliente selecionado
    End If
End Sub

Private Sub PreencherCamposCliente(clienteID As String)
    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Sheets("CLIENTES")
    ultimaLinha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To ultimaLinha
        If ws.Cells(i, 1).Value = clienteID Then
            txtID.Value = ws.Cells(i, 1).Value
            txtNomeCliente.Value = ws.Cells(i, 2).Value
            txtPessoaContato.Value = ws.Cells(i, 3).Value
            txtEndereco.Value = ws.Cells(i, 4).Value
            txtCidade.Value = ws.Cells(i, 5).Value
            txtEstado.Value = ws.Cells(i, 6).Value
            txtTelefone.Value = ws.Cells(i, 7).Value
            txtEmail.Value = ws.Cells(i, 8).Value
            Exit For
        End If
    Next i
End Sub


