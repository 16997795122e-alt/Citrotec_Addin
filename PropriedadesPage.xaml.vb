' Arquivo: PropriedadesPage.xaml.vb (VERSÃO CORRIGIDA)
Imports System.IO
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Inventor
Imports System.Threading.Tasks
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports Newtonsoft.Json
Imports System.Text.RegularExpressions
Imports System.Reflection
Imports System.Linq
Imports System.Data.SqlClient
Imports System.Windows.Input

' --- Classes de Configuração (com DisplayName) ---
Public Class PropertyDefinition
    Public Property Name As String ' Nome interno
    Public Property DisplayName As String
    Public Property IsCustom As Boolean
    Public Property Type As String = "Text"
End Class
Public Class DocumentTypeConfig
    Public Property Properties As New List(Of PropertyDefinition)
End Class
Public Class PropertyConfig
    ' REMOVIDO: A propriedade LastUserTag não é mais necessária.
    Public Property DocumentConfigs As New Dictionary(Of String, DocumentTypeConfig)
End Class

' --- Classe que representa uma propriedade na interface (com DisplayName) ---
Public Class UserProperty
    Implements INotifyPropertyChanged
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Property InternalName As String ' Nome interno, programático
    Public Property DisplayName As String ' Nome de exibição, para a UI
    Public Property IsCustom As Boolean
    Public Property Type As String
    Private _value As String
    Public Property Value As String
        Get
            Return _value
        End Get
        Set(value As String)
            _value = value
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Value"))
        End Set
    End Property
End Class

' --- CLASSE PRINCIPAL DA PÁGINA ---
Public Class PropriedadesPage
    Private _activeDoc As Document
    Private _config As PropertyConfig
    Private ReadOnly _configFilePath As String
    Public ReadOnly LeftColumnProperties As New ObservableCollection(Of UserProperty)
    Public ReadOnly RightColumnProperties As New ObservableCollection(Of UserProperty)
    Private _isDrawingLogicRunning As Boolean = False

    Public Sub New()
        InitializeComponent()
        Dim appDataPath As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData)
        Dim addinFolderPath As String = System.IO.Path.Combine(appDataPath, "CitrotecAddin")
        System.IO.Directory.CreateDirectory(addinFolderPath)
        _configFilePath = System.IO.Path.Combine(addinFolderPath, "PropriedadesConfig.json")
        LeftItemsControl.ItemsSource = LeftColumnProperties
        RightItemsControl.ItemsSource = RightColumnProperties
        LoadConfiguration()
    End Sub

#Region "Lógica Principal e Carregamento"
    Public Sub UpdateProperties(ByVal doc As Document)
        _activeDoc = doc
        LoadPropertiesFromConfig(doc)
        If doc IsNot Nothing AndAlso doc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            SyncSectraButton.Visibility = Visibility.Visible
            ExecuteDrawingUpdateLogic(doc, overwriteExisting:=False)
        Else
            SyncSectraButton.Visibility = Visibility.Collapsed
        End If
    End Sub

    Private Sub LoadPropertiesFromConfig(ByVal doc As Document)
        LeftColumnProperties.Clear()
        RightColumnProperties.Clear()
        If doc Is Nothing Then
            ClearAllFields()
            Return
        End If
        ApplyButton.IsEnabled = True
        AddPropertyButton.IsEnabled = True
        NoPropertiesMessage.Visibility = Visibility.Collapsed
        Dim docTypeKey As String = GetDocTypeKey(doc.DocumentType)
        Dim docTypeName As String = GetDocTypeName(doc.DocumentType)
        SectionHeaderLabel.Content = $"PROPRIEDADES DE {docTypeName}".ToUpper()
        If Not _config.DocumentConfigs.ContainsKey(docTypeKey) Then
            NoPropertiesMessage.Visibility = Visibility.Visible
            Return
        End If
        Dim configForDoc = _config.DocumentConfigs(docTypeKey)
        If configForDoc.Properties.Count = 0 Then
            NoPropertiesMessage.Visibility = Visibility.Visible
        End If
        For i As Integer = 0 To configForDoc.Properties.Count - 1
            Dim propDef As PropertyDefinition = configForDoc.Properties(i)
            Dim propValue As String = GetPropertyValueFromDoc(doc, propDef.Name)
            Dim finalDisplayName = If(String.IsNullOrWhiteSpace(propDef.DisplayName), propDef.Name, propDef.DisplayName)
            Dim userProp = New UserProperty With {
                .InternalName = propDef.Name,
                .DisplayName = finalDisplayName,
                .Value = propValue,
                .IsCustom = propDef.IsCustom,
                .Type = propDef.Type
            }
            If i Mod 2 = 0 Then
                LeftColumnProperties.Add(userProp)
            Else
                RightColumnProperties.Add(userProp)
            End If
        Next
    End Sub
#End Region

#Region "Lógica de Sincronização (Drawing)"
    Private Async Function ExecuteDrawingUpdateLogic(doc As Document, overwriteExisting As Boolean) As Task
        If _isDrawingLogicRunning Then Return
        If Not doc.FullFileName.Contains("\") Then
            If overwriteExisting Then
                ShowModernMessageBox("É necessário salvar o desenho antes de sincronizar com o Sectra.", "Aviso", ModernMessageBoxButtons.OK)
            End If
            Return
        End If
        _isDrawingLogicRunning = True
        Try
            Await FillDesignerAndAuthority(doc, overwriteExisting)

            Dim nroDesenhoDoArquivo As String = IO.Path.GetFileNameWithoutExtension(doc.FullFileName).Split(" "c)(0).Trim()
            Dim nroDesenhoAtual As String = GetPropertyValueFromDoc(doc, "NRODESENHO")
            Dim forcePropertyOverwrite As Boolean = False

            If String.IsNullOrWhiteSpace(nroDesenhoAtual) OrElse Not nroDesenhoAtual.Equals(nroDesenhoDoArquivo, StringComparison.OrdinalIgnoreCase) Then
                Dim deveAtualizarNomeArquivo As Boolean = False
                If Not String.IsNullOrWhiteSpace(nroDesenhoAtual) AndAlso Not overwriteExisting Then
                    Dim msgBox = New MessageBoxWindow($"O Nº do desenho na propriedade ({nroDesenhoAtual}) é diferente do nome do arquivo ({nroDesenhoDoArquivo}). Deseja corrigir a propriedade?", "Divergência Encontrada", ModernMessageBoxButtons.YesNo)
                    msgBox.Owner = Window.GetWindow(Me)
                    msgBox.ShowDialog()
                    If msgBox.Result = MessageBoxResult.Yes Then
                        deveAtualizarNomeArquivo = True
                    End If
                Else
                    deveAtualizarNomeArquivo = True
                End If
                If deveAtualizarNomeArquivo Then
                    SetInventorProperty(doc, "NRODESENHO", nroDesenhoDoArquivo)
                    UpdateDisplayedProperty("NRODESENHO", nroDesenhoDoArquivo)
                    forcePropertyOverwrite = True
                End If
            End If

            Dim nroDesenhoFinal As String = GetPropertyValueFromDoc(doc, "NRODESENHO")
            ShowStatusMessage("Sincronizando com Sectra...", isError:=False)
            Await Task.Run(Sub()
                               Try
                                   Dim connectionString As String = "Server=172.16.0.120;Database=INDUSTRIAL;User Id=ORCAMENTOS;Password=orc1248;Connection Timeout=5;"
                                   Using connection As New SqlConnection(connectionString)
                                       connection.Open()
                                       Dim nroDesenhoSemHifen As String = nroDesenhoFinal.Replace("-", "")
                                       Dim queryPrDesenho As String = "SELECT DESCRICAO, OBRA, CLIENTE FROM PR_DESENHO WHERE DESENHO = @Desenho"
                                       Using command As New SqlCommand(queryPrDesenho, connection)
                                           command.Parameters.AddWithValue("@Desenho", nroDesenhoSemHifen)
                                           Using reader As SqlDataReader = command.ExecuteReader()
                                               If reader.Read() Then
                                                   Dim descricao As String = reader("DESCRICAO").ToString()
                                                   Dim obra As String = reader("OBRA").ToString()
                                                   Dim cliente As String = reader("CLIENTE").ToString()
                                                   reader.Close()
                                                   Dispatcher.Invoke(Sub()
                                                                         If overwriteExisting OrElse forcePropertyOverwrite OrElse String.IsNullOrWhiteSpace(GetPropertyValueFromDoc(doc, "TITULO2")) Then
                                                                             SetInventorProperty(doc, "TITULO2", descricao)
                                                                             UpdateDisplayedProperty("TITULO2", descricao)
                                                                         End If
                                                                     End Sub)
                                                   Dim partesObra() As String = obra.Split("."c)
                                                   If partesObra.Length = 2 Then
                                                       Dim nroOtFormatado As String = partesObra(0).Substring(2) & "-" & partesObra(1)
                                                       Dispatcher.Invoke(Sub()
                                                                             If overwriteExisting OrElse forcePropertyOverwrite OrElse String.IsNullOrWhiteSpace(GetPropertyValueFromDoc(doc, "NROOT")) Then
                                                                                 SetInventorProperty(doc, "NROOT", nroOtFormatado)
                                                                                 UpdateDisplayedProperty("NROOT", nroOtFormatado)
                                                                             End If
                                                                         End Sub)
                                                   End If
                                                   Dim queryVeObra As String = "SELECT DESCRICAO FROM VE_OBRA WHERE CODIGO = @Obra"
                                                   Using commandObra As New SqlCommand(queryVeObra, connection)
                                                       commandObra.Parameters.AddWithValue("@Obra", obra)
                                                       Using readerObra As SqlDataReader = commandObra.ExecuteReader()
                                                           If readerObra.Read() Then
                                                               Dim descricaoObra As String = readerObra("DESCRICAO").ToString()
                                                               Dispatcher.Invoke(Sub()
                                                                                     If overwriteExisting OrElse forcePropertyOverwrite OrElse String.IsNullOrWhiteSpace(GetPropertyValueFromDoc(doc, "TITULO1")) Then
                                                                                         SetInventorProperty(doc, "TITULO1", descricaoObra)
                                                                                         UpdateDisplayedProperty("TITULO1", descricaoObra)
                                                                                     End If
                                                                                 End Sub)
                                                           End If
                                                       End Using
                                                   End Using
                                                   Dim queryFornecedores As String = "SELECT FANTASIA FROM FN_FORNECEDORES WHERE CODIGO = @Cliente"
                                                   Using commandFornec As New SqlCommand(queryFornecedores, connection)
                                                       commandFornec.Parameters.AddWithValue("@Cliente", cliente)
                                                       Using readerFornec As SqlDataReader = commandFornec.ExecuteReader()
                                                           If readerFornec.Read() Then
                                                               Dim fantasia As String = readerFornec("FANTASIA").ToString()
                                                               Dispatcher.Invoke(Sub()
                                                                                     If overwriteExisting OrElse forcePropertyOverwrite OrElse String.IsNullOrWhiteSpace(GetPropertyValueFromDoc(doc, "Company")) Then
                                                                                         SetInventorProperty(doc, "Company", fantasia)
                                                                                         UpdateDisplayedProperty("Company", fantasia)
                                                                                     End If
                                                                                 End Sub)
                                                           End If
                                                       End Using
                                                   End Using
                                               Else
                                                   ShowModernMessageBox("Desenho " & nroDesenhoSemHifen & " não encontrado no banco de dados.", "Aviso", ModernMessageBoxButtons.OK)
                                               End If
                                           End Using
                                       End Using
                                   End Using
                                   ShowStatusMessage("Sincronização concluída!", isError:=False)
                               Catch ex As Exception
                                   ShowModernMessageBox("Não foi possível conectar ao banco de dados: " & ex.Message, "Erro de Conexão", ModernMessageBoxButtons.OK)
                               End Try
                           End Sub)
        Catch ex As Exception
            ShowModernMessageBox("Ocorreu um erro na lógica de atualização: " & ex.Message, "Erro Crítico", ModernMessageBoxButtons.OK)
        Finally
            _isDrawingLogicRunning = False
        End Try
    End Function
#End Region

#Region "Lógica de Atualização (Peça/Montagem)"
    Private Async Function UpdateDescriptionFromCodigo(doc As Document) As Task
        ShowStatusMessage("Atualizando descrição a partir do Código...", isError:=False)
        Try
            Dim codigoValue As String = GetPropertyValueFromDoc(doc, "Código")
            If String.IsNullOrWhiteSpace(codigoValue) Then
                ShowStatusMessage("Campo 'Código' está vazio. Nenhuma ação foi tomada.", isError:=True)
                Return
            End If
            Dim oDescricao As String = Await Task.Run(Function()
                                                          Dim server As String = "172.16.0.120"
                                                          Dim database As String = "INDUSTRIAL"
                                                          Dim connectionString As String = $"Server={server};Database={database};User Id=ORCAMENTOS;Password=orc1248;Connection Timeout=5;"
                                                          Dim desc As String = "[ATENÇÃO] Código Inválido"
                                                          Try
                                                              Using connection As New SqlConnection(connectionString)
                                                                  connection.Open()
                                                                  Dim query As String
                                                                  If codigoValue.Length >= 8 Then
                                                                      query = "SELECT DESCRICAO FROM MT_MATERIAL WHERE CODIGO = @Codigo AND CCKATIVO = 'S'"
                                                                  Else
                                                                      query = "SELECT DESCRICAO FROM PR_DESENHO WHERE DESENHO = @Codigo"
                                                                  End If
                                                                  Using command As New SqlCommand(query, connection)
                                                                      command.Parameters.AddWithValue("@Codigo", codigoValue)
                                                                      Dim result As Object = command.ExecuteScalar()
                                                                      If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                                                                          desc = result.ToString()
                                                                      End If
                                                                  End Using
                                                              End Using
                                                          Catch ex As Exception
                                                              desc = "[ERRO] Falha de conexão"
                                                          End Try
                                                          Return desc
                                                      End Function)
            If oDescricao.Length > 60 Then
                oDescricao = oDescricao.Substring(0, 60)
            End If
            Dim finalDescription As String
            If oDescricao.StartsWith("[") Then
                finalDescription = oDescricao
            Else
                Dim dim1 As String = GetPropertyValueFromDoc(doc, "DIMENSAO1").Trim()
                Dim dim2 As String = GetPropertyValueFromDoc(doc, "DIMENSAO2").Trim()
                Select Case True
                    Case Not String.IsNullOrEmpty(dim1) AndAlso Not String.IsNullOrEmpty(dim2)
                        finalDescription = $"{oDescricao} x {dim1} x {dim2}"
                    Case Not String.IsNullOrEmpty(dim1)
                        finalDescription = $"{oDescricao} x {dim1}"
                    Case Else
                        finalDescription = oDescricao
                End Select
            End If
            SetInventorProperty(doc, "Descrição", finalDescription)
            UpdateDisplayedProperty("Descrição", finalDescription)
            ShowStatusMessage("Descrição atualizada com sucesso!", isError:=False)
        Catch ex As Exception
            ShowModernMessageBox("Ocorreu um erro inesperado ao atualizar a descrição: " & ex.Message, "Erro Crítico", ModernMessageBoxButtons.OK)
        End Try
    End Function
#End Region

#Region "Eventos de Botões e UI"
    Private Sub SyncSectraButton_Click(sender As Object, e As RoutedEventArgs)
        If _activeDoc IsNot Nothing AndAlso _activeDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            ExecuteDrawingUpdateLogic(_activeDoc, overwriteExisting:=True)
        End If
    End Sub

    Private Async Sub ApplyButton_Click(sender As Object, e As RoutedEventArgs)
        If _activeDoc Is Nothing Then Return
        Try
            For Each prop As UserProperty In LeftColumnProperties.Concat(RightColumnProperties)
                SetInventorProperty(_activeDoc, prop.InternalName, prop.Value)
            Next
            If _activeDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject OrElse _activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Await UpdateDescriptionFromCodigo(_activeDoc)
            Else
                ShowStatusMessage("Propriedades aplicadas com sucesso!", isError:=False)
            End If
        Catch ex As Exception
            ShowStatusMessage("Erro ao aplicar propriedades: " & ex.Message, isError:=True)
        End Try
    End Sub

    Private Sub AddPropertyButton_Click(sender As Object, e As RoutedEventArgs)
        If _activeDoc Is Nothing Then Return
        Dim addPropWindow = New AddPropertyWindow()
        addPropWindow.Owner = Window.GetWindow(Me)
        If Not addPropWindow.ShowDialog().GetValueOrDefault(False) Then Return
        Dim newPropInternalName = addPropWindow.InternalName
        Dim newPropDisplayName = addPropWindow.DisplayName
        If LeftColumnProperties.Concat(RightColumnProperties).Any(Function(p) p.InternalName.Equals(newPropInternalName, StringComparison.OrdinalIgnoreCase)) Then
            ShowModernMessageBox("Uma propriedade com este nome interno já existe.", "Aviso", ModernMessageBoxButtons.OK)
            Return
        End If
        Dim docTypeKey As String = GetDocTypeKey(_activeDoc.DocumentType)
        If _config.DocumentConfigs.ContainsKey(docTypeKey) Then
            _config.DocumentConfigs(docTypeKey).Properties.Add(New PropertyDefinition With {
               .Name = newPropInternalName,
                .DisplayName = newPropDisplayName,
                .IsCustom = True,
                .Type = "Text"
            })
            SaveConfiguration()
            LoadPropertiesFromConfig(_activeDoc)
        End If
    End Sub
    Private Sub DeletePropertyButton_Click(sender As Object, e As RoutedEventArgs)
        If _activeDoc Is Nothing Then Return
        Dim button = CType(sender, Button)
        Dim propToDelete = CType(button.CommandParameter, UserProperty)
        If propToDelete Is Nothing Then Return
        Dim msgBox = New MessageBoxWindow($"Tem certeza que deseja remover a propriedade '{propToDelete.DisplayName}' da configuração?", "Confirmar Exclusão", ModernMessageBoxButtons.YesNo)
        msgBox.Owner = Window.GetWindow(Me)
        msgBox.ShowDialog()
        If msgBox.Result = MessageBoxResult.Yes Then
            Dim docTypeKey As String = GetDocTypeKey(_activeDoc.DocumentType)
            If _config.DocumentConfigs.ContainsKey(docTypeKey) Then
                Dim propDefToRemove = _config.DocumentConfigs(docTypeKey).Properties.FirstOrDefault(Function(p) p.Name = propToDelete.InternalName)
                If propDefToRemove IsNot Nothing Then
                    _config.DocumentConfigs(docTypeKey).Properties.Remove(propDefToRemove)
                    SaveConfiguration()
                End If
            End If
            LoadPropertiesFromConfig(_activeDoc)
        End If
    End Sub
    Private Sub CalendarButton_Click(sender As Object, e As RoutedEventArgs)
        Dim button = CType(sender, Button)
        Dim propToUpdate = CType(button.DataContext, UserProperty)
        If propToUpdate Is Nothing Then Return
        Dim calendarView As New CalendarWindow()
        calendarView.OperationMode = CalendarWindow.CalendarMode.ReturnValue
        calendarView.Owner = Window.GetWindow(Me)
        Dim result As Boolean = calendarView.ShowDialog().GetValueOrDefault(False)
        If result AndAlso calendarView.SelectedDateValue.HasValue Then
            propToUpdate.Value = calendarView.SelectedDateValue.Value.ToString("dd/MM/yyyy")
        End If
    End Sub
#End Region

#Region "Lógica de Configuração e Funções Auxiliares"
    Private Sub UpdateDisplayedProperty(propName As String, value As String)
        Dim prop = LeftColumnProperties.FirstOrDefault(Function(p) p.InternalName.Equals(propName, StringComparison.OrdinalIgnoreCase))
        If prop Is Nothing Then
            prop = RightColumnProperties.FirstOrDefault(Function(p) p.InternalName.Equals(propName, StringComparison.OrdinalIgnoreCase))
        End If
        If prop IsNot Nothing Then
            prop.Value = value
        End If
    End Sub

    Private Async Function FillDesignerAndAuthority(doc As Document, overwrite As Boolean) As Task
        Dim userTag = SessionManager.CurrentUser?.UserTag
        If String.IsNullOrWhiteSpace(userTag) Then Return

        Dim propsToChange As New List(Of String)
        Dim designerValue = GetPropertyValueFromDoc(doc, "Designer")
        If String.IsNullOrWhiteSpace(designerValue) OrElse (overwrite AndAlso Not designerValue.Equals(userTag, StringComparison.OrdinalIgnoreCase)) Then
            propsToChange.Add("Designer")
        End If
        Dim authorityValue = GetPropertyValueFromDoc(doc, "Authority")
        If String.IsNullOrWhiteSpace(authorityValue) OrElse (overwrite AndAlso Not authorityValue.Equals(userTag, StringComparison.OrdinalIgnoreCase)) Then
            propsToChange.Add("Authority")
        End If
        If Not propsToChange.Any() Then Return
        Dim needsConfirmation As Boolean = overwrite AndAlso propsToChange.Any(Function(p) Not String.IsNullOrWhiteSpace(GetPropertyValueFromDoc(doc, p)))
        Dim shouldProceed As Boolean = True
        If needsConfirmation Then
            Dim propNames = String.Join(" e ", propsToChange)
            Dim msgBox = New MessageBoxWindow($"O(s) campo(s) '{propNames}' será(ão) substituído(s) pela sua TAG ('{userTag}')?", "Confirmar Substituição", ModernMessageBoxButtons.YesNo)
            msgBox.Owner = Window.GetWindow(Me)
            msgBox.ShowDialog()
            If msgBox.Result = MessageBoxResult.No Then
                shouldProceed = False
            End If
        End If
        If shouldProceed Then
            For Each propName In propsToChange
                SetInventorProperty(doc, propName, userTag)
                UpdateDisplayedProperty(propName, userTag)
            Next
        End If
    End Function

    Private Sub LoadConfiguration()
        Try
            If Not System.IO.File.Exists(_configFilePath) Then ExtractDefaultConfig()
            Dim json As String = System.IO.File.ReadAllText(_configFilePath)
            _config = JsonConvert.DeserializeObject(Of PropertyConfig)(json)
            If _config Is Nothing Then
                ShowModernMessageBox("O arquivo de configuração parece estar corrompido.", "Aviso", ModernMessageBoxButtons.OK)
                ExtractDefaultConfig()
                json = System.IO.File.ReadAllText(_configFilePath)
                _config = JsonConvert.DeserializeObject(Of PropertyConfig)(json)
            End If
        Catch ex As Exception
            ShowModernMessageBox("Erro fatal ao carregar a configuração: " & ex.Message, "Erro de Configuração", ModernMessageBoxButtons.OK)
            _config = New PropertyConfig()
        End Try
    End Sub

    Private Sub ExtractDefaultConfig()
        Dim resourceName As String = "CitrotecAddin.PropriedadesConfig.json"
        Try
            Using stream As Stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName)
                If stream Is Nothing Then
                    ShowModernMessageBox("ERRO: O recurso 'PropriedadesConfig.json' não foi encontrado na DLL.", "Erro Crítico", ModernMessageBoxButtons.OK)
                    Return
                End If
                Using reader As New StreamReader(stream)
                    Dim defaultConfigJson As String = reader.ReadToEnd()
                    System.IO.File.WriteAllText(_configFilePath, defaultConfigJson)
                End Using
            End Using
        Catch ex As Exception
            ShowModernMessageBox("Não foi possível extrair a configuração padrão: " & ex.Message, "Erro de Extração", ModernMessageBoxButtons.OK)
        End Try
    End Sub

    Private Sub SaveConfiguration()
        Try
            Dim json As String = JsonConvert.SerializeObject(_config, Formatting.Indented)
            System.IO.File.WriteAllText(_configFilePath, json)
        Catch ex As Exception
            ShowStatusMessage("Erro ao salvar configuração: " & ex.Message, isError:=True)
        End Try
    End Sub

    Private Sub TextBox_PreviewKeyDown(sender As Object, e As KeyEventArgs)
        If Keyboard.Modifiers = ModifierKeys.Control Then
            Dim txtBox = TryCast(sender, System.Windows.Controls.TextBox)
            If txtBox Is Nothing Then Return
            Select Case e.Key
                Case Key.C
                    If txtBox.SelectionLength > 0 Then Clipboard.SetText(txtBox.SelectedText)
                    e.Handled = True
                Case Key.V
                    If Clipboard.ContainsText() Then
                        Dim selectionStart = txtBox.SelectionStart
                        Dim selectionLength = txtBox.SelectionLength
                        Dim pasteText = Clipboard.GetText()
                        txtBox.Text = txtBox.Text.Remove(selectionStart, selectionLength).Insert(selectionStart, pasteText)
                        txtBox.CaretIndex = selectionStart + pasteText.Length
                    End If
                    e.Handled = True
                Case Key.X
                    If txtBox.SelectionLength > 0 Then
                        Clipboard.SetText(txtBox.SelectedText)
                        Dim selectionStart = txtBox.SelectionStart
                        Dim selectionLength = txtBox.SelectionLength
                        txtBox.Text = txtBox.Text.Remove(selectionStart, selectionLength)
                        txtBox.CaretIndex = selectionStart
                    End If
                    e.Handled = True
            End Select
        End If
    End Sub
    Private Sub SetInventorProperty(doc As Document, propName As String, value As Object)
        If doc Is Nothing Then Return
        Dim propSet As PropertySet
        Select Case propName.ToLower()
            Case "part number", "revision number", "designer", "authority", "project", "stock number", "creation time", "date checked", "engr date approved"
                propSet = GetPropertySet(doc, "Design Tracking Properties")
            Case "title", "author"
                propSet = GetPropertySet(doc, "Inventor Summary Information")
            Case "company"
                propSet = GetPropertySet(doc, "Document Summary Information")
            Case Else
                propSet = GetPropertySet(doc, "Inventor User Defined Properties")
        End Select
        Try
            Dim prop As Inventor.Property = Nothing
            Try
                prop = propSet.Item(propName)
            Catch
            End Try
            If prop IsNot Nothing Then
                If prop.Value?.ToString() <> value?.ToString() Then
                    prop.Value = value
                End If
            Else
                If propSet.Name = "Inventor User Defined Properties" Then
                    propSet.Add(value, propName)
                End If
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"Falha ao definir propriedade '{propName}': {ex.Message}")
        End Try
    End Sub
    Private Function GetPropertyValueFromDoc(doc As Document, propName As String) As String
        If doc Is Nothing Then Return ""
        Dim propSet As PropertySet
        Dim propValue As Object = ""
        Select Case propName.ToLower()
            Case "part number", "revision number", "designer", "authority", "project", "stock number", "creation time", "date checked", "engr date approved"
                propSet = GetPropertySet(doc, "Design Tracking Properties")
            Case "title", "author"
                propSet = GetPropertySet(doc, "Inventor Summary Information")
            Case "company"
                propSet = GetPropertySet(doc, "Document Summary Information")
            Case Else
                propSet = GetPropertySet(doc, "Inventor User Defined Properties")
        End Select
        Try
            propValue = propSet.Item(propName).Value
        Catch
            Return ""
        End Try
        If TypeOf propValue Is Date Then
            Dim dt = CDate(propValue)
            If dt.Year < 1900 Then Return ""
            Return dt.ToString("dd/MM/yyyy")
        End If
        Return propValue?.ToString()
    End Function
    Private Function GetDocTypeKey(docType As DocumentTypeEnum) As String
        Select Case docType
            Case DocumentTypeEnum.kPartDocumentObject : Return "Part"
            Case DocumentTypeEnum.kAssemblyDocumentObject : Return "Assembly"
            Case DocumentTypeEnum.kDrawingDocumentObject : Return "Drawing"
            Case DocumentTypeEnum.kPresentationDocumentObject : Return "Presentation"
            Case Else : Return "Unknown"
        End Select
    End Function
    Private Function GetDocTypeName(docType As DocumentTypeEnum) As String
        Select Case docType
            Case DocumentTypeEnum.kPartDocumentObject : Return "Peça"
            Case DocumentTypeEnum.kAssemblyDocumentObject : Return "Montagem"
            Case DocumentTypeEnum.kDrawingDocumentObject : Return "Desenho"
            Case DocumentTypeEnum.kPresentationDocumentObject : Return "Apresentação"
            Case Else : Return "Desconhecido"
        End Select
    End Function
    Private Function GetOrCreateProperty(doc As Document, propName As String, defaultValue As Object) As Inventor.Property
        Dim customPropSet As PropertySet = GetPropertySet(doc, "Inventor User Defined Properties")
        Try
            Return customPropSet.Item(propName)
        Catch
            Return customPropSet.Add(defaultValue, propName)
        End Try
    End Function
    Private Function GetPropertySet(doc As Document, setName As String) As PropertySet
        Try
            Return doc.PropertySets.Item(setName)
        Catch ex As Exception
            Return doc.PropertySets.Add(setName)
        End Try
    End Function
    Private Sub ClearAllFields()
        ApplyButton.IsEnabled = False
        AddPropertyButton.IsEnabled = False
        SyncSectraButton.Visibility = Visibility.Collapsed
        NoPropertiesMessage.Visibility = Visibility.Visible
        SectionHeaderLabel.Content = "AMBIENTE NÃO IDENTIFICADO"
        LeftColumnProperties.Clear()
        RightColumnProperties.Clear()
    End Sub
    Private Sub ShowModernMessageBox(message As String, title As String, buttons As ModernMessageBoxButtons)
        Dispatcher.Invoke(Sub()
                              Dim msgBox = New MessageBoxWindow(message, title, buttons)
                              msgBox.Owner = Window.GetWindow(Me)
                              msgBox.ShowDialog()
                          End Sub)
    End Sub
    Private Async Sub ShowStatusMessage(message As String, isError As Boolean)
        Await Dispatcher.InvokeAsync(Async Sub()
                                         If isError Then
                                             NotificationBorder.Background = New SolidColorBrush(System.Windows.Media.Color.FromRgb(211, 47, 47))
                                         Else
                                             NotificationBorder.Background = New SolidColorBrush(System.Windows.Media.Color.FromRgb(250, 166, 26))
                                         End If
                                         NotificationTextBlock.Text = message
                                         NotificationBorder.Visibility = Visibility.Visible
                                         Await Task.Delay(3000)
                                         NotificationBorder.Visibility = Visibility.Collapsed
                                     End Sub)
    End Sub
#End Region
End Class