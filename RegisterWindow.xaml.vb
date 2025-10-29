Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Windows
Imports System.Windows.Controls

Public Class RegisterWindow
    ' REMOVIDO: A lista de avatares e a variável do avatar selecionado.
    ' Public Avatars As New List(Of String)
    ' Private SelectedAvatarPath As String

    Public Sub New()
        InitializeComponent()
        ' REMOVIDO: A lógica que carregava os avatares na tela[cite: 60].
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    ' REMOVIDO: A função Avatar_Checked não é mais necessária[cite: 61].
    ' Private Sub Avatar_Checked(sender As Object, e As RoutedEventArgs)
    '     ...
    ' End Sub

    Private Sub RegisterButton_Click(sender As Object, e As RoutedEventArgs)
        ErrorTextBlock.Visibility = Visibility.Collapsed

        If String.IsNullOrWhiteSpace(UsernameTextBox.Text) OrElse String.IsNullOrWhiteSpace(FullNameTextBox.Text) OrElse String.IsNullOrWhiteSpace(PasswordBox.Password) Then
            ShowError("Todos os campos devem ser preenchidos.")
            Return
        End If

        If PasswordBox.Password <> ConfirmPasswordBox.Password Then
            ShowError("As senhas não coincidem.")
            Return
        End If

        Dim selectedSexo As String = ""
        If MascRadioButton.IsChecked.GetValueOrDefault(False) Then
            selectedSexo = "M"
        ElseIf FemRadioButton.IsChecked.GetValueOrDefault(False) Then
            selectedSexo = "F"
        Else
            ShowError("Por favor, selecione o sexo.")
            Return
        End If

        Dim userTag = UserTagTextBox.Text.ToUpper().Trim()
        If Not Regex.IsMatch(userTag, "^([A-Z]\.)+$") Then
            ShowError("O formato da TAG do usuário é inválido (ex: N.O.M.).")
            Return
        End If


        ' REMOVIDO: A validação que verificava se um avatar foi selecionado[cite: 63].

        ' MUDANÇA: Passamos uma string vazia "" como avatar, pois ele não é mais definido aqui[cite: 64].
        Dim success As Boolean = DatabaseManager.AddUser(UsernameTextBox.Text.Trim(), FullNameTextBox.Text.Trim(), PasswordBox.Password, selectedSexo, userTag)

        If success Then
            ' MUDANÇA: A mensagem de sucesso foi simplificada [cite: 64-65].
            MessageBox.Show("Cadastro realizado com sucesso! Você já pode fazer o login.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
            Me.DialogResult = True
            Me.Close()
        Else
            ShowError("Este nome de usuário já existe. Por favor, escolha outro.")
        End If
    End Sub

    Private Sub ShowError(message As String)
        ErrorTextBlock.Text = message
        ErrorTextBlock.Visibility = Visibility.Visible
    End Sub
End Class