' Arquivo: Announcement.vb
Public Class Announcement
    Public Property Id As Integer
    Public Property AdminUserName As String ' Quem postou
    Public Property Content As String     ' A mensagem
    Public Property Timestamp As DateTime ' Quando foi postado

    ' Propriedade auxiliar para formatar a data/hora para exibição
    Public ReadOnly Property DisplayTimestamp As String
        Get
            Return Timestamp.ToString("dd/MM/yyyy HH:mm")
        End Get
    End Property
End Class