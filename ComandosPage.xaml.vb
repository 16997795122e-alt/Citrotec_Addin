' Arquivo: ComandosPage.xaml.vb
Imports System.Windows
Imports System.Windows.Controls
Imports System.Linq

Public Class ComandosPage
    Private m_allCommandsInView As New List(Of InventorCommand)

    Public Sub UpdateCommands(ByVal commands As List(Of InventorCommand))
        m_allCommandsInView = commands
        ' Carrega o estado 'Fixado' dos comandos a partir do arquivo de configurações
        Dim settings = UserSettingsManager.LoadSettings()
        For Each cmd In m_allCommandsInView
            cmd.IsPinned = settings.PinnedCommands.Contains(cmd.DisplayName)
        Next
        FilterAndDisplayCommands()
    End Sub

    Public Sub UpdateEnvironmentTitle(ByVal title As String)
        EnvironmentTitleLabel.Text = title
    End Sub

    Private Sub FilterAndDisplayCommands()
        Dim searchText = SearchTextBox.Text.ToLower().Trim()
        Dim filteredCommands As New List(Of InventorCommand)
        Dim currentUserLevel As Integer = If(SessionManager.IsUserLoggedIn, SessionManager.CurrentUser.Nivel, 0)

        For Each cmd In m_allCommandsInView
            If cmd.MinimumUserLevel <= currentUserLevel Then
                If String.IsNullOrWhiteSpace(searchText) OrElse
                   cmd.DisplayName.ToLower().Contains(searchText) OrElse
                   cmd.Keywords.Exists(Function(kw) kw.ToLower().Contains(searchText)) Then
                    filteredCommands.Add(cmd)
                End If
            End If
        Next

        ' NOVO: Ordena a lista para que os comandos fixados apareçam primeiro
        Dim sortedCommands = filteredCommands.OrderByDescending(Function(c) c.IsPinned).ToList()

        CommandsItemsControl.ItemsSource = Nothing
        CommandsItemsControl.ItemsSource = sortedCommands
    End Sub

    Private Sub SearchTextBox_TextChanged(sender As Object, e As TextChangedEventArgs)
        FilterAndDisplayCommands()
    End Sub

    Private Sub CommandButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim selectedCommand = CType(button.DataContext, InventorCommand)

        If selectedCommand IsNot Nothing Then
            selectedCommand.Action.Invoke()
        End If
    End Sub

    Private Sub PinButton_Click(sender As Object, e As RoutedEventArgs)
        Dim toggleButton = CType(sender, Controls.Primitives.ToggleButton)
        Dim commandToPin = CType(toggleButton.DataContext, InventorCommand)

        If commandToPin IsNot Nothing Then
            If commandToPin.IsPinned Then
                UserSettingsManager.PinCommand(commandToPin.DisplayName)
            Else
                UserSettingsManager.UnpinCommand(commandToPin.DisplayName)
            End If
            ' Reordena a lista para refletir a mudança visualmente
            FilterAndDisplayCommands()
        End If
    End Sub
End Class