' Arquivo: SessionManager.vb
Public Module SessionManager
    Public CurrentUser As User

    Public Sub Login(user As User)
        CurrentUser = user
    End Sub

    Public Sub Logout()
        CurrentUser = Nothing
    End Sub

    Public ReadOnly Property IsUserLoggedIn As Boolean
        Get
            Return CurrentUser IsNot Nothing
        End Get
    End Property

    Public ReadOnly Property IsAdmin As Boolean
        Get
            If IsUserLoggedIn Then
                Return CurrentUser.Nivel = 3
            End If
            Return False
        End Get
    End Property
End Module