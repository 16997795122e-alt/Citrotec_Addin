' Arquivo: JanelaPrincipal.xaml.vb
Imports System
Imports System.ComponentModel
Imports System.IO
Imports System.Linq ' Adicionado para a função .Cast(Of ...).FirstOrDefault
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Threading
Imports CitrotecAddin.CitrotecAddin
Imports Inventor

Public Class JanelaPrincipal
    Inherits Window

    Public ReadOnly m_homePage As New HomePage()
    Public ReadOnly m_comandosPage As New ComandosPage()
    Public ReadOnly m_bibliotecaPage As New BibliotecaPage()
    Public ReadOnly m_propriedadesPage As New PropriedadesPage()
    Public ReadOnly m_bomPage As New BOMPage()
    ' Declaração da nova página
    Public ReadOnly m_pesquisaPage As New PesquisaMateriaisPage()

    Private m_clockTimer As DispatcherTimer
    Private m_apontamentoTimer As DispatcherTimer
    Private _isApontamentoRunning As Boolean = False
    Private _isApontamentoTimerEnabled As Boolean = False

    Public WasLogoutRequested As Boolean = False
    Private Const VAULT_ADDIN_DISPLAY_NAME As String = "Inventor Vault"

    Public Property IsApontamentoTimerEnabled As Boolean
        Get
            Return _isApontamentoTimerEnabled
        End Get
        Set(value As Boolean)
            _isApontamentoTimerEnabled = value
            If m_apontamentoTimer Is Nothing Then Return

            If value = True Then
                If Not m_apontamentoTimer.IsEnabled AndAlso Not _isApontamentoRunning Then
                    m_apontamentoTimer.Start()
                    LogStatus("Timer de Apontamento ATIVADO.")
                End If
            Else
                If m_apontamentoTimer.IsEnabled Then
                    m_apontamentoTimer.Stop()
                    LogStatus("Timer de Apontamento DESATIVADO.")
                End If
            End If
        End Set
    End Property

    Public Sub New()
        InitializeComponent()
        CommandManager.PopulateAllCommands()
        MainFrame.Content = m_homePage
        InitializeClock()
        InitializeApontamentoTimer()

        If SessionManager.IsUserLoggedIn Then
            UserPanel.DataContext = SessionManager.CurrentUser
        End If

        UpdateVaultToggleState()
        AddHandler Me.Closing, AddressOf JanelaPrincipal_Closing
    End Sub

#Region "Lógica da Janela Principal (Relógio, Navegação, Update, etc.)"

    Private Sub InitializeClock()
        m_clockTimer = New DispatcherTimer()
        m_clockTimer.Interval = TimeSpan.FromSeconds(1)
        AddHandler m_clockTimer.Tick, AddressOf ClockTimer_Tick
        m_clockTimer.Start()
        ClockTimer_Tick(Nothing, Nothing)
    End Sub

    Private Sub ClockTimer_Tick(sender As Object, e As EventArgs)
        DateTimeTextBlock.Text = DateTime.Now.ToString("dd/MM/yyyy - HH:mm")
    End Sub

    Private Sub CalendarButton_Click(sender As Object, e As RoutedEventArgs)
        WindowManager.ShowWindow(Of CalendarWindow)()
    End Sub

    Private Sub LogoutButton_Click(sender As Object, e As RoutedEventArgs)
        SessionManager.Logout()
        Dim settingsPath = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData), "CitrotecAddin", "user.setting")
        If System.IO.File.Exists(settingsPath) Then
            System.IO.File.Delete(settingsPath)
        End If
        WasLogoutRequested = True
        Me.Close()
    End Sub

    Public Sub OnActiveDocumentChanged(ByVal activeDoc As Document)
        If Not Me.IsVisible Then Return

        Dim currentDocType = If(activeDoc IsNot Nothing, activeDoc.DocumentType, DocumentTypeEnum.kUnknownDocumentObject)
        Dim environmentTitle As String = "Nenhum ambiente ativo"

        If activeDoc IsNot Nothing Then
            Select Case activeDoc.DocumentType
                Case DocumentTypeEnum.kPartDocumentObject : environmentTitle = "Ambiente de Peça"
                Case DocumentTypeEnum.kAssemblyDocumentObject : environmentTitle = "Ambiente de Montagem"
                Case DocumentTypeEnum.kDrawingDocumentObject : environmentTitle = "Ambiente de Desenho"
                Case DocumentTypeEnum.kPresentationDocumentObject : environmentTitle = "Ambiente de Apresentação"
            End Select
        End If

        Dim commandsForDoc = CommandManager.AllCommands.FindAll(Function(cmd) cmd.ApplicableDocumentTypes.Contains(currentDocType))

        Me.Dispatcher.Invoke(Sub()
                                 m_comandosPage.UpdateEnvironmentTitle(environmentTitle)
                                 m_comandosPage.UpdateCommands(commandsForDoc)
                                 m_propriedadesPage.UpdateProperties(activeDoc)
                                 m_bomPage.UpdateActiveDocument(activeDoc)
                                 m_homePage.UpdateForActiveDocument(activeDoc, environmentTitle)
                             End Sub)
    End Sub

    Private Sub UpdateButton_Click(sender As Object, e As RoutedEventArgs)
        Dim sourceDir As String = "\\SERVIDOR20\eng\Engenharia\iLogic\CitrotecAddin"
        Dim appDataPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)
        Dim destDir As String = System.IO.Path.Combine(appDataPath, "Autodesk", "Inventor 2025", "Addins", "CitrotecAddin")

        Try
            DirectoryCopy(sourceDir, destDir, True)
            ShowModernMessageBox("Add-in atualizado com sucesso!" & vbCrLf & vbCrLf & "Por favor, reinicie o Autodesk Inventor para aplicar as alterações.",
                                 "Atualização Concluída", ModernMessageBoxButtons.OK)
        Catch ex As Exception
            ShowModernMessageBox("Ocorreu um erro ao tentar atualizar o add-in:" & vbCrLf & vbCrLf & ex.Message,
                                 "Erro de Atualização", ModernMessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub DirectoryCopy(sourceDirName As String, destDirName As String, copySubDirs As Boolean)
        Dim dir As New System.IO.DirectoryInfo(sourceDirName)
        If Not dir.Exists Then Throw New System.IO.DirectoryNotFoundException("Diretório não encontrado: " & sourceDirName)

        If Not System.IO.Directory.Exists(destDirName) Then
            System.IO.Directory.CreateDirectory(destDirName)
        End If

        For Each file As System.IO.FileInfo In dir.GetFiles()
            Dim tempPath As String = System.IO.Path.Combine(destDirName, file.Name)
            file.CopyTo(tempPath, True)
        Next

        If copySubDirs Then
            For Each subdir In dir.GetDirectories()
                Dim tempPath As String = System.IO.Path.Combine(destDirName, subdir.Name)
                DirectoryCopy(subdir.FullName, tempPath, copySubDirs)
            Next
        End If
    End Sub

    Private Sub NavRadio_Checked(sender As Object, e As RoutedEventArgs)
        If MainFrame Is Nothing Then Return
        Dim radio As System.Windows.Controls.RadioButton = CType(sender, System.Windows.Controls.RadioButton)

        Select Case radio.Name
            Case "HomeRadio"
                MainFrame.Content = m_homePage
                m_homePage.LoadPinnedCommands()
                m_homePage.LoadAnnouncements()
            Case "ComandosRadio"
                MainFrame.Content = m_comandosPage
            Case "BibliotecaRadio"
                MainFrame.Content = m_bibliotecaPage
            Case "PropriedadesRadio"
                MainFrame.Content = m_propriedadesPage
            Case "BOMRadio"
                MainFrame.Content = m_bomPage
                m_bomPage.UpdateActiveDocument(g_inventorApplication.ActiveDocument)
            ' Case Adicionado
            Case "PesquisaRadio"
                MainFrame.Content = m_pesquisaPage
                ' Fim do Case Adicionado
        End Select
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub btnMinimize_Click(sender As Object, e As RoutedEventArgs)
        Me.WindowState = WindowState.Minimized
    End Sub
#End Region

#Region "Lógica do Vault Toggle"

    Private Sub VaultToggleButton_Click(sender As Object, e As RoutedEventArgs)
        Dim toggleButton As System.Windows.Controls.Primitives.ToggleButton = CType(sender, System.Windows.Controls.Primitives.ToggleButton)
        Dim desiredStateActivated As Boolean = toggleButton.IsChecked.GetValueOrDefault()

        Dim vaultAddin As ApplicationAddIn = Nothing

        Try
            LogStatus($"Tentando alterar estado do Vault para: {(If(desiredStateActivated, "Ativado", "Desativado"))}")

            vaultAddin = g_inventorApplication.ApplicationAddIns _
                .Cast(Of ApplicationAddIn) _
                .FirstOrDefault(Function(a) a.DisplayName.Equals(VAULT_ADDIN_DISPLAY_NAME, StringComparison.InvariantCultureIgnoreCase))

            If vaultAddin Is Nothing Then
                Throw New Exception($"Add-In '{VAULT_ADDIN_DISPLAY_NAME}' não encontrado.")
            End If

            If desiredStateActivated Then
                If Not vaultAddin.Activated Then
                    LogStatus("Ativando o Vault...")
                    vaultAddin.Activate()
                    If Not vaultAddin.Activated Then
                        Throw New Exception("Falha ao ativar o Vault. O Add-in não respondeu ao comando.")
                    End If
                    LogStatus("Vault ativado.")
                End If
            Else
                If vaultAddin.Activated Then
                    LogStatus("Desativando o Vault...")
                    vaultAddin.Deactivate()
                    If vaultAddin.Activated Then
                        Throw New Exception("Falha ao desativar o Vault. O Add-in não respondeu ao comando.")
                    End If
                    LogStatus("Vault desativado.")
                End If
            End If

        Catch ex As Exception
            LogStatus($"ERRO ao tentar ativar/desativar o Vault: {ex.Message}")
            ShowModernMessageBox($"Ocorreu um erro ao tentar alterar o estado do Vault Add-in:{vbCrLf}{ex.Message}", "Erro Vault", ModernMessageBoxButtons.OK)
        Finally
            UpdateVaultToggleState()
        End Try
    End Sub

    Private Sub UpdateVaultToggleState()
        Dim vaultAddin As ApplicationAddIn = Nothing
        Dim isCurrentlyActivated As Boolean = False

        Try
            vaultAddin = g_inventorApplication.ApplicationAddIns _
                .Cast(Of ApplicationAddIn) _
                .FirstOrDefault(Function(a) a.DisplayName.Equals(VAULT_ADDIN_DISPLAY_NAME, StringComparison.InvariantCultureIgnoreCase))

            If vaultAddin IsNot Nothing Then
                isCurrentlyActivated = vaultAddin.Activated
                VaultToggleButton.IsEnabled = True
            Else
                VaultToggleButton.IsEnabled = False
                VaultStatusText.Text = "Vault não encontrado"
                LogStatus("AVISO: Add-in do Vault não encontrado.")
                Return
            End If

        Catch ex As Exception
            LogStatus($"Erro ao verificar estado do Vault: {ex.Message}")
            VaultToggleButton.IsEnabled = False
            VaultStatusText.Text = "Erro ao verificar"
            Return
        End Try

        VaultToggleButton.IsChecked = isCurrentlyActivated
        VaultStatusText.Text = If(isCurrentlyActivated, "Vault Ligado", "Vault Desligado")
        LogStatus($"Estado atual do Vault verificado: {(If(isCurrentlyActivated, "Ligado", "Desligado"))}")
    End Sub

    Private Sub ShowModernMessageBox(message As String, title As String, buttons As ModernMessageBoxButtons)
        Dim msgBox = New MessageBoxWindow(message, title, buttons)
        msgBox.Owner = Me
        msgBox.ShowDialog()
    End Sub

    Private Sub LogStatus(message As String)
        Debug.WriteLine($"{DateTime.Now:HH:mm:ss} - Vault Toggle: {message}")
    End Sub
#End Region

#Region "Lógica de Apontamento"

    Private Sub InitializeApontamentoTimer()
        m_apontamentoTimer = New DispatcherTimer()
        m_apontamentoTimer.Interval = TimeSpan.FromMinutes(1)
        AddHandler m_apontamentoTimer.Tick, AddressOf ApontamentoTimer_Tick

        If _isApontamentoTimerEnabled Then
            m_apontamentoTimer.Start()
        End If
    End Sub

    Private Sub ApontamentoTimer_Tick(sender As Object, e As EventArgs)
        If Not _isApontamentoTimerEnabled OrElse _isApontamentoRunning Then Return

        Try
            _isApontamentoRunning = True
            m_apontamentoTimer.Stop()

            Dim apontamentoWin As New CitrotecAddin.ApontamentoWindow()
            apontamentoWin.Owner = Me
            apontamentoWin.ShowDialog()

        Catch ex As Exception
            LogStatus($"ERRO ao abrir ApontamentoWindow: {ex.Message}")
        Finally
            _isApontamentoRunning = False
            If _isApontamentoTimerEnabled Then m_apontamentoTimer.Start()
        End Try
    End Sub

    Private Sub JanelaPrincipal_Closing(sender As Object, e As CancelEventArgs)
        If m_clockTimer IsNot Nothing Then
            m_clockTimer.Stop()
            RemoveHandler m_clockTimer.Tick, AddressOf ClockTimer_Tick
        End If

        If m_apontamentoTimer IsNot Nothing Then
            m_apontamentoTimer.Stop()
            RemoveHandler m_apontamentoTimer.Tick, AddressOf ApontamentoTimer_Tick
        End If
    End Sub

#End Region

End Class