' Arquivo: PostAnnouncementWindow.xaml.vb
Imports System.Windows

Public Class PostAnnouncementWindow

    Public Sub New()
        InitializeComponent()
        AddHandler Me.Loaded, Sub(sender, e) ContentTextBox.Focus()
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub btnPost_Click(sender As Object, e As RoutedEventArgs)
        Dim content = ContentTextBox.Text.Trim()

        If String.IsNullOrWhiteSpace(content) Then
            MessageBox.Show("O conteúdo do aviso não pode estar vazio.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Try
            ' Pega o nome do admin logado
            Dim adminUserName = SessionManager.CurrentUser?.NomeCompleto ' Usando NomeCompleto para exibição
            If String.IsNullOrEmpty(adminUserName) Then
                MessageBox.Show("Não foi possível identificar o administrador logado.", "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
                Return
            End If

            DatabaseManager.AddAnnouncement(adminUserName, content)
            MessageBox.Show("Aviso postado com sucesso!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro ao postar o aviso: " & ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub
End Class