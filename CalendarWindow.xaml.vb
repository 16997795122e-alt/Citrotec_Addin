' Arquivo: CalendarWindow.xaml.vb
Imports System.Windows
Imports System.Windows.Controls ' Mantemos este para outros controles (Button, etc.)
Imports System.Globalization ' Mantemos este para Clipboard e CultureInfo

Public Class CalendarWindow
    Inherits System.Windows.Window

    Public Enum CalendarMode
        CopyToClipboard
        ReturnValue
    End Enum

    Public Property OperationMode As CalendarMode = CalendarMode.CopyToClipboard
    Public Property SelectedDateValue As Nullable(Of DateTime)

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub MainCalendar_SelectedDatesChanged(sender As Object, e As SelectionChangedEventArgs)
        ' ***** CORREÇÃO AQUI: Especifica o namespace completo *****
        Dim calendar = TryCast(sender, System.Windows.Controls.Calendar)
        ' ***** FIM CORREÇÃO *****

        If calendar Is Nothing OrElse Not calendar.SelectedDate.HasValue Then Return

        Dim selectedDate As DateTime = calendar.SelectedDate.Value

        Select Case Me.OperationMode
            Case CalendarMode.CopyToClipboard
                Dim formattedDate As String = selectedDate.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
                Try
                    Clipboard.SetText(formattedDate)
                Catch ex As Exception
                    Debug.WriteLine("Erro ao copiar data para Clipboard: " & ex.Message)
                End Try
                Me.Close()

            Case CalendarMode.ReturnValue
                SelectedDateValue = selectedDate
                Try
                    Me.DialogResult = True
                Catch ex As InvalidOperationException
                    Debug.WriteLine("AVISO: Tentativa de definir DialogResult em janela não modal no modo ReturnValue.")
                End Try
                Me.Close()
        End Select
    End Sub

End Class