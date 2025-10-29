' Arquivo: SelectModelStateWindow.xaml.vb
Imports System.Collections.Generic
Imports System.Windows

Public Class SelectModelStateWindow
    Private _SelectedModelStateName As String

    Public ReadOnly Property SelectedModelStateName As String
        Get
            Return _SelectedModelStateName
        End Get
    End Property

    Public Sub New(modelStateNames As List(Of String), Optional activeStateName As String = Nothing)
        InitializeComponent()

        ModelStateComboBox.ItemsSource = modelStateNames

        If Not String.IsNullOrEmpty(activeStateName) AndAlso modelStateNames.Contains(activeStateName) Then
            ModelStateComboBox.SelectedItem = activeStateName
        ElseIf modelStateNames.Count > 0 Then
            ModelStateComboBox.SelectedIndex = 0
        End If
    End Sub

    Public Sub New(title As String, prompt As String, options As List(Of String), Optional defaultOption As String = Nothing)
        InitializeComponent()

        ' Define os textos customizados da UI
        TitleTextBlock.Text = title
        PromptTextBlock.Text = prompt

        ' Popula o ComboBox com as opções
        ModelStateComboBox.ItemsSource = options

        ' Define a seleção padrão
        If Not String.IsNullOrEmpty(defaultOption) AndAlso options.Contains(defaultOption) Then
            ModelStateComboBox.SelectedItem = defaultOption
        ElseIf options.Count > 0 Then
            ModelStateComboBox.SelectedIndex = 0
        End If
    End Sub

    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs)
        If ModelStateComboBox.SelectedItem IsNot Nothing Then
            _SelectedModelStateName = ModelStateComboBox.SelectedItem.ToString()
            Me.DialogResult = True
        Else
            MessageBox.Show("Por favor, selecione um item.", "Seleção Necessária", MessageBoxButton.OK, MessageBoxImage.Warning)
        End If
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.DialogResult = False
    End Sub
End Class