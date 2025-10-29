Imports System.Windows

Public Enum ModernMessageBoxButtons
    OK
    YesNo
End Enum

Public Class MessageBoxWindow
    Public Result As MessageBoxResult

    Public Sub New(message As String, title As String, buttons As ModernMessageBoxButtons)
        InitializeComponent()
        MessageTextBlock.Text = message
        TitleTextBlock.Text = title

        Select Case buttons
            Case ModernMessageBoxButtons.OK
                btnOK.Visibility = Visibility.Visible
                btnOK.IsDefault = True
            Case ModernMessageBoxButtons.YesNo
                btnYes.Visibility = Visibility.Visible
                btnNo.Visibility = Visibility.Visible
                btnYes.IsDefault = True
                btnNo.IsCancel = True
        End Select
    End Sub

    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs)
        Me.Result = MessageBoxResult.OK
        Me.Close()
    End Sub

    Private Sub btnYes_Click(sender As Object, e As RoutedEventArgs)
        Me.Result = MessageBoxResult.Yes
        Me.Close()
    End Sub

    Private Sub btnNo_Click(sender As Object, e As RoutedEventArgs)
        Me.Result = MessageBoxResult.No
        Me.Close()
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub
End Class