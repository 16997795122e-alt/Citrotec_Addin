' Arquivo: InventorCommand.vb (VERSÃO COM SUPORTE A "FIXAR")
Imports Inventor
Imports System.ComponentModel
Imports System.Runtime.CompilerServices

Public Class InventorCommand
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

    Private Sub OnPropertyChanged(<CallerMemberName> Optional propertyName As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    Public Property DisplayName As String
    Public Property Keywords As List(Of String)
    Public Property ApplicableDocumentTypes As List(Of DocumentTypeEnum)
    Public Property Action As Action
    Public Property IconAcronym As String
    Public Property Description As String
    Public Property IconImagePath As String
    Public Property MinimumUserLevel As Integer

    ' NOVO: Propriedade para controlar o estado "Fixado"
    Private _isPinned As Boolean
    Public Property IsPinned As Boolean
        Get
            Return _isPinned
        End Get
        Set(value As Boolean)
            If _isPinned <> value Then
                _isPinned = value
                OnPropertyChanged() ' Notifica a UI da mudança
            End If
        End Set
    End Property


    Public Sub New(displayName As String, iconAcronym As String, keywords As List(Of String),
                   applicableTypes As List(Of DocumentTypeEnum), action As Action,
                   Optional description As String = "Nenhuma descrição disponível.",
                   Optional minimumLevel As Integer = 1,
                   Optional iconImagePath As String = Nothing)

        Me.DisplayName = displayName
        Me.IconAcronym = iconAcronym
        Me.Keywords = keywords
        Me.ApplicableDocumentTypes = applicableTypes
        Me.Action = action
        Me.Description = description
        Me.MinimumUserLevel = minimumLevel
        Me.IsPinned = False ' Padrão é não fixado
    End Sub
End Class