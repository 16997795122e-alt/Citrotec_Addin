' Arquivo: ResetPasswordWindow.xaml.vb
Imports System.Windows

Public Class ResetPasswordWindow
    Private ReadOnly _currentUser As User

    ' O construtor da janela recebe o objeto do usuário que precisa resetar a senha
    Public Sub New(userToReset As User)
        InitializeComponent()
        _currentUser = userToReset
        ' Foca no primeiro campo de senha quando a janela abre
        NewPasswordBox.Focus()
    End Sub

    Private Sub SaveButton_Click(sender As Object, e As RoutedEventArgs)
        ErrorTextBlock.Visibility = Visibility.Collapsed

        If String.IsNullOrWhiteSpace(NewPasswordBox.Password) OrElse String.IsNullOrWhiteSpace(ConfirmPasswordBox.Password) Then
            ShowError("Ambos os campos de senha devem ser preenchidos.")
            Return
        End If

        ' MUDANÇA: A validação de 6 caracteres foi REMOVIDA.
        ' If NewPasswordBox.Password.Length < 6 Then
        '     ShowError("A nova senha deve ter pelo menos 6 caracteres.")
        '     Return
        ' End If

        If NewPasswordBox.Password <> ConfirmPasswordBox.Password Then
            ShowError("As senhas não coincidem.")
            Return
        End If

        ' Tenta atualizar a senha no banco de dados usando a nova função
        Dim success As Boolean = DatabaseManager.UpdateUserPassword(_currentUser.Id, NewPasswordBox.Password)

        If success Then
            MessageBox.Show("Senha atualizada com sucesso! Por favor, faça o login novamente com sua nova senha.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True ' Sinaliza que a operação foi bem-sucedida
            Me.Close()
        Else
            ShowError("Ocorreu um erro ao tentar atualizar sua senha. Por favor, contate um administrador.")
        End If
    End Sub

    Private Sub ShowError(message As String)
        ErrorTextBlock.Text = message
        ErrorTextBlock.Visibility = Visibility.Visible
    End Sub
End Class