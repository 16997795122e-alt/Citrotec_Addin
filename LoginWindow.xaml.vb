' Arquivo: LoginWindow.xaml.vb (VERSÃO COM AUTO-LOGIN)
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.Windows

Public Class LoginWindow
    Public Sub New()
        InitializeComponent()
        AddHandler Loaded, AddressOf LoginWindow_Loaded
    End Sub

    Private Sub LoginWindow_Loaded(sender As Object, e As RoutedEventArgs)
        Dim settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CitrotecAddin", "user.setting")

        If File.Exists(settingsPath) Then
            Try
                Dim encryptedBytes As Byte() = File.ReadAllBytes(settingsPath)
                Dim originalBytes As Byte() = ProtectedData.Unprotect(encryptedBytes, Nothing, DataProtectionScope.CurrentUser)
                Dim credentials As String = Encoding.UTF8.GetString(originalBytes)
                Dim parts() As String = credentials.Split(vbTab)
                If parts.Length = 2 Then
                    UsernameTextBox.Text = parts(0)
                    PasswordBox.Password = parts(1)
                    RememberMeCheckBox.IsChecked = True

                    ' NOVO: Tenta fazer o login automaticamente
                    AttemptLogin()
                End If
            Catch ex As Exception
                UsernameTextBox.Focus()
            End Try
        Else
            UsernameTextBox.Focus()
        End If
    End Sub

    ' NOVO: Lógica de login refatorada para uma sub-rotina separada
    Private Sub AttemptLogin()
        ' Se a janela já estiver fechando (por um login automático bem-sucedido), não faz nada
        If Not Me.IsVisible Then Return

        ErrorTextBlock.Visibility = Visibility.Collapsed
        Dim username = UsernameTextBox.Text

        If String.IsNullOrWhiteSpace(username) Then
            ShowError("Por favor, informe o nome de usuário.")
            Return
        End If

        Dim userToCheck = DatabaseManager.GetUserByUsername(username)

        If userToCheck Is Nothing Then
            ShowError("Usuário não encontrado.")
            Return
        End If

        If userToCheck.ResetSenhaObrigatorio Then
            MessageBox.Show("Por motivos de segurança, você precisa definir uma nova senha para continuar.", "Redefinição de Senha Necessária", MessageBoxButton.OK, MessageBoxImage.Information)
            Dim resetWindow As New ResetPasswordWindow(userToCheck)
            resetWindow.Owner = Me
            resetWindow.ShowDialog()
            PasswordBox.Clear()
            Return
        End If

        Dim authenticatedUser = DatabaseManager.AuthenticateUser(username, PasswordBox.Password)

        If authenticatedUser IsNot Nothing Then
            SessionManager.Login(authenticatedUser)
            Dim settingsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CitrotecAddin", "user.setting")

            If RememberMeCheckBox.IsChecked.GetValueOrDefault(False) Then
                Dim credentials As String = $"{authenticatedUser.NomeUsuario}{vbTab}{PasswordBox.Password}"
                Dim originalBytes As Byte() = Encoding.UTF8.GetBytes(credentials)
                Dim encryptedBytes As Byte() = ProtectedData.Protect(originalBytes, Nothing, DataProtectionScope.CurrentUser)
                File.WriteAllBytes(settingsPath, encryptedBytes)
            Else
                If File.Exists(settingsPath) Then
                    File.Delete(settingsPath)
                End If
            End If

            Me.DialogResult = True
            Me.Close()
        Else
            ShowError("Senha incorreta.")
        End If
    End Sub

    Private Sub ShowError(message As String)
        ErrorTextBlock.Text = message
        ErrorTextBlock.Visibility = Visibility.Visible
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub LoginButton_Click(sender As Object, e As RoutedEventArgs)
        ' O botão agora simplesmente chama a nova rotina de login
        AttemptLogin()
    End Sub

    Private Sub RegisterButton_Click(sender As Object, e As RoutedEventArgs)
        Dim registerWindow = New RegisterWindow()
        registerWindow.Owner = Me
        registerWindow.ShowDialog()
    End Sub
End Class