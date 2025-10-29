' Arquivo: InputBoxWindow.xaml.vb (VERSÃO CORRIGIDA)
Imports System.Windows

Public Class InputBoxWindow

    ' Propriedade pública para retornar o texto digitado pelo usuário
    Public Property InputText As String

    Public Sub New()
        InitializeComponent()
        ' Foca na caixa de texto assim que a janela abre
        ' ===================== CORREÇÃO AQUI =====================
        ' Trocamos "Loaded += Sub(...)" por "AddHandler Me.Loaded, Sub(...)"
        AddHandler Me.Loaded, Sub(sender, e) InputTextBox.Focus()
        ' =======================================================
    End Sub

    ' Construtor que permite customizar o título e o prompt
    Public Sub New(prompt As String, title As String)
        Me.New()
        Me.TitleTextBlock.Text = title
        Me.PromptTextBlock.Text = prompt
    End Sub

    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs)
        ' Verifica se o usuário digitou algo antes de confirmar
        If Not String.IsNullOrWhiteSpace(InputTextBox.Text) Then
            Me.InputText = InputTextBox.Text.Trim()
            Me.DialogResult = True
            Me.Close()
        Else
            MessageBox.Show("O nome da propriedade não pode estar vazio.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning)
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub
End Class