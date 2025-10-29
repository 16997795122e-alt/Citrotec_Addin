' Arquivo: AdminPage.xaml.vb
Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Controls

Public Class AdminPage

    Public Sub New()
        InitializeComponent()
        LoadUsers()
    End Sub

    Private Sub LoadUsers()
        ' Busca todos os usuários do banco de dados
        Dim userList = DatabaseManager.GetAllUsers()
        ' Define a lista como a fonte de dados do DataGrid
        UsersDataGrid.ItemsSource = userList
    End Sub

    Private Sub SaveChangesButton_Click(sender As Object, e As RoutedEventArgs)
        ' Pega a lista de usuários diretamente do DataGrid (que pode ter sido editada)
        Dim usersToSave = CType(UsersDataGrid.ItemsSource, List(Of User))

        If usersToSave IsNot Nothing Then
            Try
                For Each user In usersToSave
                    ' Manda cada usuário para a rotina de atualização no banco de dados
                    DatabaseManager.UpdateUser(user)
                Next
                MessageBox.Show("Alterações salvas com sucesso!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro ao salvar as alterações: " & ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Sub DeleteUserButton_Click(sender As Object, e As RoutedEventArgs)
        ' Pega o botão que foi clicado
        Dim button = CType(sender, Button)
        ' Pega o usuário associado a essa linha do DataGrid
        Dim userToDelete = CType(button.DataContext, User)

        If userToDelete Is Nothing Then Return

        ' Medida de segurança: não permitir deletar o usuário 'admin' principal
        If userToDelete.NomeUsuario.ToLower() = "admin" Then
            MessageBox.Show("Não é possível deletar o administrador principal do sistema.", "Ação Proibida", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Pede confirmação antes de deletar
        Dim result = MessageBox.Show($"Tem certeza que deseja deletar o usuário '{userToDelete.NomeUsuario}'?" & vbCrLf & "Esta ação não pode ser desfeita.",
                                     "Confirmar Exclusão", MessageBoxButton.YesNo, MessageBoxImage.Warning)

        If result = MessageBoxResult.Yes Then
            Try
                ' Chama a função para deletar o usuário do banco de dados
                DatabaseManager.DeleteUser(userToDelete.Id)
                MessageBox.Show("Usuário deletado com sucesso!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
                ' Recarrega a lista de usuários no DataGrid para refletir a mudança
                LoadUsers()
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro ao deletar o usuário: " & ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    ' MUDANÇA: Nova sub-rotina para o clique do botão de reset de senha.
    Private Sub ResetPasswordButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim userToReset = CType(button.DataContext, User)

        If userToReset Is Nothing Then Return

        ' Medida de segurança: não permitir resetar a senha do 'admin' principal
        If userToReset.NomeUsuario.ToLower() = "admin" Then
            MessageBox.Show("Não é possível resetar a senha do administrador principal por aqui.", "Ação Proibida", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim result = MessageBox.Show($"Tem certeza que deseja forçar o usuário '{userToReset.NomeUsuario}' a resetar a senha no próximo login?",
                                     "Confirmar Reset de Senha", MessageBoxButton.YesNo, MessageBoxImage.Question)

        If result = MessageBoxResult.Yes Then
            Try
                ' Chama a nova função no DatabaseManager para marcar o usuário
                DatabaseManager.FlagUserForPasswordReset(userToReset.Id)
                MessageBox.Show("Usuário marcado para reset de senha. A alteração será exigida no próximo login.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
            Catch ex As Exception
                MessageBox.Show("Ocorreu um erro ao marcar o usuário para reset: " & ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End If
    End Sub

    Private Sub BackButton_Click(sender As Object, e As RoutedEventArgs)
        ' Navega de volta para a página anterior (a HomePage)
        If Me.NavigationService.CanGoBack Then
            Me.NavigationService.GoBack()
        End If
    End Sub

End Class