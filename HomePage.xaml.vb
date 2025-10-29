' Arquivo: HomePage.xaml.vb
Imports System.Collections.ObjectModel
Imports System.Linq
Imports System.Windows
Imports System.Windows.Controls
Imports Inventor

Public Class HomePage
    Private _activeDoc As Document

    Public ReadOnly Property IsUserAdmin As Boolean
        Get
            Return SessionManager.IsAdmin
        End Get
    End Property

    Public ReadOnly PinnedCommands As New ObservableCollection(Of InventorCommand)
    Public ReadOnly Announcements As New ObservableCollection(Of Announcement)

    Public Sub New()
        InitializeComponent()
        Me.DataContext = Me ' Necessário para Bindings como IsUserAdmin
        PinnedCommandsItemsControl.ItemsSource = PinnedCommands
        AnnouncementsItemsControl.ItemsSource = Announcements
        UpdateWelcomeMessage()
        LoadAnnouncements() ' Carrega anúncios ao iniciar a página
    End Sub

    Public Sub UpdateForActiveDocument(ByVal doc As Document, ByVal environmentName As String)
        _activeDoc = doc

        If environmentName = "Nenhum ambiente ativo" Then
            PinnedCommandsTitle.Text = "Comandos Fixados"
        Else
            Dim cleanEnvName = environmentName.Replace("Ambiente de ", "")
            PinnedCommandsTitle.Text = $"Comandos Fixados - {cleanEnvName}"
        End If

        LoadPinnedCommands()
        LoadAnnouncements() ' Recarrega anúncios para garantir a visibilidade correta
    End Sub

    Public Sub LoadPinnedCommands()
        PinnedCommands.Clear()
        Dim settings = UserSettingsManager.LoadSettings()
        Dim pinnedNames = settings.PinnedCommands
        Dim currentDocType = If(_activeDoc IsNot Nothing, _activeDoc.DocumentType, DocumentTypeEnum.kUnknownDocumentObject)

        If pinnedNames.Any() Then
            Dim commandsToShow = CommandManager.AllCommands.Where(
                Function(cmd) pinnedNames.Contains(cmd.DisplayName) AndAlso
                              cmd.ApplicableDocumentTypes.Contains(currentDocType)
            ).ToList()

            For Each cmd In commandsToShow
                cmd.IsPinned = True
                PinnedCommands.Add(cmd)
            Next
        End If

        If PinnedCommands.Count = 0 Then
            NoPinnedCommandsMessage.Visibility = Visibility.Visible
        Else
            NoPinnedCommandsMessage.Visibility = Visibility.Collapsed
        End If
    End Sub

    ' ***** CORREÇÃO: Filtra anúncios ocultos *****
    ' ***** CORREÇÃO: Carrega para todos, filtra ocultos só para não-admins *****
    Public Sub LoadAnnouncements()
        Announcements.Clear()
        Dim allRecentAnnouncements = DatabaseManager.GetRecentAnnouncements()
        Dim announcementsToShow As List(Of Announcement)

        ' Verifica se o usuário é admin
        If SessionManager.IsAdmin Then
            ' Admins veem todos os anúncios recentes
            announcementsToShow = allRecentAnnouncements
        Else
            ' Usuários normais não veem os que foram ocultos
            Dim userSettings = UserSettingsManager.LoadSettings()
            announcementsToShow = allRecentAnnouncements.Where(
                Function(ann) Not userSettings.HiddenAnnouncementIds.Contains(ann.Id)
            ).ToList()
        End If

        ' Popula a lista na UI e controla a visibilidade do painel/mensagem
        If announcementsToShow.Any() Then
            For Each announcement In announcementsToShow
                Announcements.Add(announcement)
            Next
            AnnouncementsPanel.Visibility = Visibility.Visible
            NoAnnouncementsMessage.Visibility = Visibility.Collapsed
        Else
            ' Se for admin, pode não haver anúncios, mas o painel deve aparecer se a intenção for mostrar a msg "nenhum aviso"
            ' Se não for admin, pode não haver anúncios visíveis (ou pq não existem ou pq foram ocultos)
            AnnouncementsPanel.Visibility = Visibility.Visible ' Mantém o painel visível para mostrar a mensagem
            NoAnnouncementsMessage.Visibility = Visibility.Visible
        End If

        ' Garante que o painel fique escondido se o usuário for admin E não houver nenhum anúncio recente
        If SessionManager.IsAdmin AndAlso Not allRecentAnnouncements.Any() Then
            AnnouncementsPanel.Visibility = Visibility.Collapsed
        End If

    End Sub

    Private Sub UpdateWelcomeMessage()
        If SessionManager.IsUserLoggedIn Then
            WelcomeTitleText.Text = $"Bem-vindo, {SessionManager.CurrentUser.NomeCompleto}!"
        Else
            WelcomeTitleText.Text = "Bem-vindo!"
        End If
    End Sub

    Private Sub AdminPanelButton_Click(sender As Object, e As RoutedEventArgs)
        If Me.NavigationService IsNot Nothing Then
            Me.NavigationService.Navigate(New AdminPage())
        End If
    End Sub

    Private Sub PostAnnouncementButton_Click(sender As Object, e As RoutedEventArgs)
        WindowManager.ShowWindow(Of PostAnnouncementWindow)()
    End Sub

    Private Sub CommandButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim selectedCommand = CType(button.DataContext, InventorCommand)

        If selectedCommand IsNot Nothing Then
            selectedCommand.Action.Invoke()
        End If
    End Sub

    Public Sub PinButton_Click(sender As Object, e As RoutedEventArgs)
        Dim toggleButton = CType(sender, Controls.Primitives.ToggleButton)
        Dim commandToUnpin = CType(toggleButton.DataContext, InventorCommand)

        If commandToUnpin IsNot Nothing Then
            UserSettingsManager.UnpinCommand(commandToUnpin.DisplayName)
            LoadPinnedCommands()
        End If
    End Sub

    ' ***** NOVO: Lógica para Ocultar Anúncio *****
    Private Sub HideAnnouncementButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim announcementToHide = CType(button.DataContext, Announcement)

        If announcementToHide IsNot Nothing Then
            UserSettingsManager.HideAnnouncement(announcementToHide.Id)
            LoadAnnouncements() ' Recarrega a lista para remover o item oculto
        End If
    End Sub

    Private Sub ViewHistoryButton_Click(sender As Object, e As RoutedEventArgs)
        WindowManager.ShowWindow(Of AnnouncementHistoryWindow)()
    End Sub

    ' ***** NOVO: Lógica para Excluir Anúncio *****
    Private Sub DeleteAnnouncementButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim announcementToDelete = CType(button.DataContext, Announcement)

        If announcementToDelete IsNot Nothing Then
            Dim result = MessageBox.Show($"Tem certeza que deseja excluir permanentemente este aviso?" & vbCrLf & $"'{announcementToDelete.Content}'",
                                         "Confirmar Exclusão", MessageBoxButton.YesNo, MessageBoxImage.Warning)

            If result = MessageBoxResult.Yes Then
                Try
                    DatabaseManager.DeleteAnnouncement(announcementToDelete.Id)
                    MessageBox.Show("Aviso excluído com sucesso.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
                    LoadAnnouncements() ' Recarrega a lista para remover o item excluído
                Catch ex As Exception
                    MessageBox.Show("Erro ao excluir o aviso: " & ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
                End Try
            End If
        End If
    End Sub

End Class