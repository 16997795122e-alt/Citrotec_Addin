Imports System.Windows
Imports System.Text.RegularExpressions

Public Class AddPropertyWindow

    Public Property InternalName As String
    Public Property DisplayName As String

    Public Sub New()
        InitializeComponent()
        AddHandler Me.Loaded, Sub(sender, e) InternalNameTextBox.Focus()
    End Sub

    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs)
        ' Validação 1: Nenhum campo pode estar vazio
        If String.IsNullOrWhiteSpace(InternalNameTextBox.Text) OrElse String.IsNullOrWhiteSpace(DisplayNameTextBox.Text) Then
            MessageBox.Show("Ambos os campos devem ser preenchidos.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        ' Validação 2: Nome interno não pode conter espaços ou caracteres especiais (exceto underscore)
        If Not Regex.IsMatch(InternalNameTextBox.Text, "^[a-zA-Z0-9_]+$") Then
            MessageBox.Show("O Nome Interno da propriedade só pode conter letras, números e o caractere underscore (_).", "Formato Inválido", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Me.InternalName = InternalNameTextBox.Text.Trim()
        Me.DisplayName = DisplayNameTextBox.Text.Trim()
        Me.DialogResult = True
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub
End Class