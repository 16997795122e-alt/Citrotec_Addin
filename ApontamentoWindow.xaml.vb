Imports System.ComponentModel
Imports System.Windows

' Adiciona o Namespace para corresponder ao XAML
Namespace CitrotecAddin

    ' A classe PRECISA ser "Partial" para se conectar com o XAML
    Partial Public Class ApontamentoWindow
        Inherits System.Windows.Window

        ' Flag para permitir o fechamento apenas via botão
        Private _permitirFechar As Boolean = False

        Public Sub New()
            ' Este método conecta o XAML com este arquivo de código
            InitializeComponent()
        End Sub

        ' Handler para o evento Click do botão
        Private Sub btnConfirmar_Click(sender As Object, e As RoutedEventArgs)
            ' 1. Permite que a janela seja fechada
            _permitirFechar = True

            ' 2. Fecha a janela
            Me.Close()
        End Sub

        ' Handler para o evento Closing da janela
        Private Sub Window_Closing(sender As Object, e As CancelEventArgs)
            ' Se o usuário tentar fechar (Ex: Alt+F4) e NÃO for pelo botão "Confirmar",
            ' nós cancelamos a ação de fechamento.
            If Not _permitirFechar Then
                e.Cancel = True
            End If
        End Sub
    End Class

End Namespace