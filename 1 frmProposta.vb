' Código para o formulário frmProposta

Private Sub btnBuscaCliente_Click()
    ' Abre o formulário frmCliente
    frmCliente.Show
End Sub

' Esta sub será chamada pelo frmCliente quando um cliente for selecionado
Public Sub PreencherDadosCliente(nome As String, cidade As String, estado As String)
    txtNomeCliente.Value = nome
    txtCidade.Value = cidade
    txtEstado.Value = estado
End Sub