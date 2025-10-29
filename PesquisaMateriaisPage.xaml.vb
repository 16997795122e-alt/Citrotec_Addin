' Arquivo: PesquisaMateriaisPage.xaml.vb
Imports System.Windows.Controls

Public Class PesquisaMateriaisPage
    Inherits Page

    Public Sub New()
        InitializeComponent()
        ' Define o DataContext para o ViewModel, que contém toda a lógica
        Me.DataContext = New PesquisaMateriaisViewModel()
    End Sub

End Class