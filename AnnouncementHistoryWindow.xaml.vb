' Arquivo: AnnouncementHistoryWindow.xaml.vb
Imports System.Collections.ObjectModel
Imports System.Windows
Imports System.Windows.Controls

Public Class AnnouncementHistoryWindow
    ' Coleção para vincular ao DataGrid
    Public ReadOnly AnnouncementsHistory As New ObservableCollection(Of Announcement)

    Public Sub New()
        InitializeComponent()
        HistoryDataGrid.ItemsSource = AnnouncementsHistory ' Define a fonte de dados
        LoadHistory() ' Carrega os dados ao abrir
    End Sub

    Private Sub LoadHistory()
        AnnouncementsHistory.Clear()
        Dim allAnnouncements = DatabaseManager.GetAllAnnouncements() ' Usa a nova função
        For Each announcement In allAnnouncements
            AnnouncementsHistory.Add(announcement)
        Next
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    ' Lógica para excluir um aviso do histórico
    Private Sub DeleteHistoryAnnouncementButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button = TryCast(sender, Button)
        Dim announcementToDelete = TryCast(button?.DataContext, Announcement)

        If announcementToDelete IsNot Nothing Then
            Dim result = MessageBox.Show($"Tem certeza que deseja excluir permanentemente este aviso do histórico?" & vbCrLf & $"'{announcementToDelete.Content}'",
                                         "Confirmar Exclusão", MessageBoxButton.YesNo, MessageBoxImage.Warning)

            If result = MessageBoxResult.Yes Then
                Try
                    DatabaseManager.DeleteAnnouncement(announcementToDelete.Id)
                    ' Remove da coleção visível e recarrega (ou apenas remove se preferir)
                    AnnouncementsHistory.Remove(announcementToDelete)
                    ' Opcional: Mostrar uma mensagem de sucesso
                    ' MessageBox.Show("Aviso excluído do histórico.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
                Catch ex As Exception
                    MessageBox.Show("Erro ao excluir o aviso: " & ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
                End Try
            End If
        End If
    End Sub
End Class