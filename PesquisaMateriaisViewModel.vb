' Arquivo: PesquisaMateriaisViewModel.vb (Corrigido para Sintaxe e Erro RCW)
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices
Imports System.Threading.Tasks
Imports System.Windows
Imports System.Windows.Input ' Import necessário para ICommand
Imports Inventor ' Import para interagir com Inventor API
Imports System.Runtime.InteropServices ' Import para ReleaseComObject

Public Class PesquisaMateriaisViewModel
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Private Sub OnPropertyChanged(<CallerMemberName> Optional propertyName As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
    End Sub

    ' --- Propriedades para BINDING ---
    Private _searchCodigo As String = ""
    Public Property SearchCodigo As String
        Get
            Return _searchCodigo
        End Get
        Set(value As String)
            If _searchCodigo <> value Then
                _searchCodigo = value
                OnPropertyChanged()
            End If
        End Set
    End Property

    Private _searchOT As String = ""
    Public Property SearchOT As String
        Get
            Return _searchOT
        End Get
        Set(value As String)
            If _searchOT <> value Then
                _searchOT = value
                OnPropertyChanged()
            End If
        End Set
    End Property

    Private _searchDescricao As String = ""
    Public Property SearchDescricao As String
        Get
            Return _searchDescricao
        End Get
        Set(value As String)
            If _searchDescricao <> value Then
                _searchDescricao = value
                OnPropertyChanged()
            End If
        End Set
    End Property

    Private _isSearching As Boolean = False
    Public Property IsSearching As Boolean
        Get
            Return _isSearching
        End Get
        Set(value As Boolean)
            If _isSearching <> value Then
                _isSearching = value
                OnPropertyChanged()
                ' CORREÇÃO: Especifica o namespace System.Windows
                Dim dispatcher = System.Windows.Application.Current?.Dispatcher
                If dispatcher IsNot Nothing Then
                    ' Usar BeginInvoke ou InvokeAsync para garantir que a atualização ocorra na thread da UI
                    dispatcher.BeginInvoke(New Action(Sub()
                                                          SearchByCodeCommand.RaiseCanExecuteChanged()
                                                          SearchByFilterCommand.RaiseCanExecuteChanged()
                                                          ClearFiltersCommand.RaiseCanExecuteChanged()
                                                          ApplyMaterialCommand.RaiseCanExecuteChanged()
                                                      End Sub))
                Else
                    ' Fallback se o dispatcher não estiver disponível (menos comum em WPF)
                    SearchByCodeCommand.RaiseCanExecuteChanged()
                    SearchByFilterCommand.RaiseCanExecuteChanged()
                    ClearFiltersCommand.RaiseCanExecuteChanged()
                    ApplyMaterialCommand.RaiseCanExecuteChanged()
                End If
            End If
        End Set
    End Property

    Private _statusMessage As String = "Pronto para buscar."
    Public Property StatusMessage As String
        Get
            Return _statusMessage
        End Get
        Set(value As String)
            If _statusMessage <> value Then
                _statusMessage = value
                OnPropertyChanged()
            End If
        End Set
    End Property

    Private _lastSearchTerms As String = ""
    Public Property LastSearchTerms As String
        Get
            Return _lastSearchTerms
        End Get
        Set(value As String)
            If _lastSearchTerms <> value Then
                _lastSearchTerms = value
                OnPropertyChanged()
            End If
        End Set
    End Property

    Public ReadOnly Property SearchResults As New ObservableCollection(Of MaterialSearchResult)()

    ' --- Comandos (Botões) ---
    Public ReadOnly Property SearchByCodeCommand As New RelayCommand(AddressOf ExecuteSearchByCodeAsync, Function(param As Object) Not IsSearching) ' Renomeado para Async
    Public ReadOnly Property SearchByFilterCommand As New RelayCommand(AddressOf ExecuteSearchByFilterAsync, Function(param As Object) Not IsSearching) ' Renomeado para Async
    Public ReadOnly Property CopyCodeCommand As New RelayCommand(AddressOf ExecuteCopyCode)
    Public ReadOnly Property CopyDescriptionCommand As New RelayCommand(AddressOf ExecuteCopyDescription)
    Public ReadOnly Property ClearFiltersCommand As New RelayCommand(AddressOf ExecuteClearFilters, Function(param As Object) Not IsSearching)
    Public ReadOnly Property ApplyMaterialCommand As New RelayCommand(AddressOf ExecuteApplyMaterial, Function(param As Object) Not IsSearching)


    ' --- Lógica de Execução ---

    ' Renomeado para indicar que é Async
    Private Async Sub ExecuteSearchByCodeAsync(param As Object)
        If String.IsNullOrWhiteSpace(SearchCodigo) Then
            StatusMessage = "Digite um Código ou Desenho para buscar."
            Return
        End If
        Dim codigoParaBuscar = SearchCodigo.Trim()
        LastSearchTerms = $"Busca por Código/Desenho: '{codigoParaBuscar}'"
        IsSearching = True
        SearchResults.Clear()
        StatusMessage = "Buscando por código no Sectra..."
        Dim result As MaterialSearchResult = Nothing
        Try
            result = Await Task.Run(Function()
                                        Dim server As String = "172.16.0.120"
                                        Dim database As String = "INDUSTRIAL"
                                        Dim connectionString As String = $"Server={server};Database={database};User Id=ORCAMENTOS;Password=orc1248;Connection Timeout=5;"
                                        Dim materialValue As String = ""
                                        Using connection As New SqlConnection(connectionString)
                                            connection.Open()
                                            Dim query As String

                                            ' 1. Busca em MT_MATERIAL (com filtros CCKATIVO e '01')
                                            query = "SELECT DESCRICAO FROM MT_MATERIAL WHERE CODIGO = @CodPesq AND CCKATIVO = 'S' AND CODIGO LIKE '01%'"
                                            Using command As New SqlCommand(query, connection)
                                                command.Parameters.AddWithValue("@CodPesq", codigoParaBuscar)
                                                Dim queryResult = command.ExecuteScalar()
                                                If queryResult IsNot Nothing AndAlso Not DBNull.Value.Equals(queryResult) Then
                                                    materialValue = queryResult.ToString()
                                                End If
                                            End Using

                                            ' 2. Busca em PR_DESENHO se não achou em MT_MATERIAL (com filtro 88/89/85)
                                            If String.IsNullOrEmpty(materialValue) Then
                                                query = "SELECT DESCRICAO FROM PR_DESENHO " &
                                                        "WHERE DESENHO = @CodPesq " &
                                                        "AND DESENHO NOT LIKE '88%' " &
                                                        "AND DESENHO NOT LIKE '89%' " &
                                                        "AND DESENHO NOT LIKE '85%'"
                                                Using command As New SqlCommand(query, connection)
                                                    command.Parameters.AddWithValue("@CodPesq", codigoParaBuscar)
                                                    Dim queryResult = command.ExecuteScalar()
                                                    If queryResult IsNot Nothing AndAlso Not DBNull.Value.Equals(queryResult) Then
                                                        materialValue = queryResult.ToString()
                                                    End If
                                                End Using
                                            End If
                                        End Using
                                        If Not String.IsNullOrEmpty(materialValue) Then
                                            Return New MaterialSearchResult With {.Codigo = codigoParaBuscar, .Descricao = materialValue}
                                        End If
                                        Return Nothing
                                    End Function)
            If result IsNot Nothing Then
                SearchResults.Add(result)
                StatusMessage = "1 resultado encontrado."
            Else
                StatusMessage = "Nenhum material ativo/válido encontrado com este código (ou código não começa com '01', ou é um desenho com prefixo 88/89/85)."
            End If
        Catch ex As SqlException
            StatusMessage = "Erro de SQL: Verifique a conexão com o banco."
            LastSearchTerms = ""
        Catch ex As Exception
            StatusMessage = "Erro: " & ex.Message
            LastSearchTerms = ""
        Finally
            IsSearching = False
        End Try
    End Sub

    ' Renomeado para indicar que é Async
    Private Async Sub ExecuteSearchByFilterAsync(param As Object)
        Dim filtroOT = SearchOT?.Trim()
        Dim filtroDescCompleto = SearchDescricao
        If String.IsNullOrWhiteSpace(filtroOT) AndAlso String.IsNullOrWhiteSpace(filtroDescCompleto) Then
            StatusMessage = "Digite uma OT ou termos de Descrição para buscar."
            Return
        End If
        Dim filtrosDescricao As New List(Of String)
        If Not String.IsNullOrWhiteSpace(filtroDescCompleto) Then
            filtrosDescricao.AddRange(filtroDescCompleto.Split(";"c).Where(Function(s) Not String.IsNullOrWhiteSpace(s))) ' Corrigido IsNullOrEmpty para IsNullOrWhiteSpace
        End If
        Dim termsDisplay As New List(Of String)
        If Not String.IsNullOrWhiteSpace(filtroOT) Then termsDisplay.Add($"OT: '{filtroOT}'")
        If filtrosDescricao.Any() Then termsDisplay.Add($"Descrição: {String.Join(" - ", filtrosDescricao)}")
        LastSearchTerms = "Busca por Filtros: " & String.Join(" | ", termsDisplay)
        IsSearching = True
        SearchResults.Clear()
        StatusMessage = "Buscando por filtros no Sectra..."
        Dim results As List(Of MaterialSearchResult) = Nothing
        Try
            results = Await Task.Run(Function()
                                         Dim server As String = "172.16.0.120"
                                         Dim database As String = "INDUSTRIAL"
                                         Dim connectionString As String = $"Server={server};Database={database};User Id=ORCAMENTOS;Password=orc1248;Connection Timeout=5;"
                                         Dim materialList As New List(Of MaterialSearchResult)()
                                         Dim queryBase As String
                                         Dim whereClauses As New List(Of String)
                                         Dim orderByClause As String = " ORDER BY DESCRICAO"
                                         Using connection As New SqlConnection(connectionString)
                                             connection.Open()
                                             Using command As New SqlCommand()
                                                 command.Connection = connection
                                                 If String.IsNullOrWhiteSpace(filtroOT) Then
                                                     ' Busca em MT_MATERIAL (com filtros CCKATIVO e '01')
                                                     queryBase = "SELECT DISTINCT CODIGO, DESCRICAO FROM MT_MATERIAL"
                                                     whereClauses.Add("CCKATIVO = 'S' AND CODIGO LIKE '01%'")
                                                 Else
                                                     ' Busca em PR_DESENHO
                                                     queryBase = "SELECT DISTINCT DESENHO, DESCRICAO FROM PR_DESENHO"
                                                     whereClauses.Add("OBRA LIKE '%' + @oNROOT + '%' " &
                                                                      "AND DESENHO NOT LIKE '88%' " &
                                                                      "AND DESENHO NOT LIKE '89%' " &
                                                                      "AND DESENHO NOT LIKE '85%'")
                                                     command.Parameters.AddWithValue("@oNROOT", filtroOT)
                                                 End If
                                                 For i As Integer = 0 To filtrosDescricao.Count - 1 ' Corrigido loop para usar Integer
                                                     Dim paramName = $"@FiltroDescricao{i + 1}"
                                                     whereClauses.Add($"DESCRICAO LIKE '%' + {paramName} + '%'")
                                                     command.Parameters.AddWithValue(paramName, filtrosDescricao(i))
                                                 Next
                                                 command.CommandText = queryBase & " WHERE " & String.Join(" AND ", whereClauses) & orderByClause
                                                 Using reader As SqlDataReader = command.ExecuteReader()
                                                     While reader.Read()
                                                         materialList.Add(New MaterialSearchResult With {.Codigo = reader(0).ToString(), .Descricao = reader("DESCRICAO").ToString()})
                                                     End While
                                                 End Using
                                             End Using
                                         End Using
                                         Return materialList
                                     End Function)
            If results IsNot Nothing AndAlso results.Count > 0 Then
                For Each item In results
                    SearchResults.Add(item)
                Next
                StatusMessage = $"{results.Count} resultados encontrados."
            Else
                ' Mensagem atualizada para incluir a nova restrição
                StatusMessage = "Nenhum material ativo/válido encontrado com esses filtros (ou códigos não começam com '01', ou são desenhos com prefixo 88/89/85)."
            End If
        Catch ex As SqlException
            StatusMessage = "Erro de SQL: Verifique a conexão com o banco."
            System.Diagnostics.Debug.WriteLine("SQL Error: " & ex.Message)
            If ex.InnerException IsNot Nothing Then System.Diagnostics.Debug.WriteLine("Inner SQL Error: " & ex.InnerException.Message)
            LastSearchTerms = ""
        Catch ex As Exception
            StatusMessage = "Erro inesperado: " & ex.Message
            System.Diagnostics.Debug.WriteLine("General Error: " & ex.ToString())
            LastSearchTerms = ""
        Finally
            IsSearching = False
        End Try
    End Sub

    Private Sub ExecuteCopyCode(parameter As Object)
        Dim result = TryCast(parameter, MaterialSearchResult)
        If result IsNot Nothing AndAlso Not String.IsNullOrEmpty(result.Codigo) Then
            Try
                Clipboard.SetText(result.Codigo)
                Dim originalStatus = StatusMessage
                StatusMessage = $"Código '{result.Codigo}' copiado!"
                ' Usar ContinueWith com TaskScheduler é mais robusto para UI updates
                Task.Delay(2000).ContinueWith(Sub(t)
                                                  If StatusMessage = $"Código '{result.Codigo}' copiado!" Then
                                                      StatusMessage = originalStatus
                                                  End If
                                              End Sub, TaskScheduler.FromCurrentSynchronizationContext())
            Catch ex As Exception
                StatusMessage = "Falha ao copiar para o clipboard."
            End Try
        End If
    End Sub

    Private Sub ExecuteCopyDescription(parameter As Object)
        Dim result = TryCast(parameter, MaterialSearchResult)
        If result IsNot Nothing AndAlso Not String.IsNullOrEmpty(result.Descricao) Then
            Try
                Clipboard.SetText(result.Descricao)
                Dim originalStatus = StatusMessage
                StatusMessage = $"Descrição '{result.Descricao}' copiada!"
                Task.Delay(2000).ContinueWith(Sub(t)
                                                  If StatusMessage = $"Descrição '{result.Descricao}' copiada!" Then
                                                      StatusMessage = originalStatus
                                                  End If
                                              End Sub, TaskScheduler.FromCurrentSynchronizationContext())
            Catch ex As Exception
                StatusMessage = "Falha ao copiar descrição para o clipboard."
            End Try
        End If
    End Sub

    Private Sub ExecuteClearFilters(param As Object)
        SearchCodigo = ""
        SearchOT = ""
        SearchDescricao = ""
        SearchResults.Clear()
        LastSearchTerms = ""
        StatusMessage = "Filtros limpos. Pronto para buscar."
    End Sub

    Private Sub ExecuteApplyMaterial(parameter As Object)
        Dim result = TryCast(parameter, MaterialSearchResult)
        If result Is Nothing OrElse String.IsNullOrEmpty(result.Codigo) Then
            StatusMessage = "Material inválido selecionado."
            Return
        End If

        Dim invApp As Inventor.Application = Nothing
        Dim activeDoc As Document = Nothing
        Dim customPropSet As PropertySet = Nothing
        Dim originalStatus As String = StatusMessage ' Salva o status atual
        Dim transaction As Transaction = Nothing

        Try
            invApp = Globals.g_inventorApplication ' Assume que Globals.g_inventorApplication está acessível
            If invApp Is Nothing Then
                StatusMessage = "Erro: Instância do Inventor não encontrada."
                Return
            End If

            activeDoc = invApp.ActiveDocument

            If activeDoc Is Nothing Then
                StatusMessage = "Nenhum documento ativo no Inventor."
                Return
            End If
            If Not (activeDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject OrElse
                    activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                StatusMessage = "Abra uma Peça (.ipt) ou Montagem (.iam) para aplicar o material."
                Return
            End If

            transaction = invApp.TransactionManager.StartTransaction(activeDoc, "Aplicar Material (Citrotec Addin)")

            StatusMessage = $"Aplicando '{result.Codigo}'..."

            customPropSet = GetOrAddPropertySet(activeDoc, "Inventor User Defined Properties", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            If customPropSet Is Nothing Then
                Throw New Exception("Não foi possível obter ou criar o PropertySet 'Inventor User Defined Properties'.")
            End If

            Dim codigoApplied As Boolean = SetPropertyValue(customPropSet, "Código", result.Codigo)
            Dim descApplied As Boolean = SetPropertyValue(customPropSet, "Descrição", result.Descricao)

            transaction.End()

            If codigoApplied Or descApplied Then
                StatusMessage = $"Material '{result.Codigo}' aplicado com sucesso!"
                ' activeDoc.Update() ' Descomente se a atualização for necessária imediatamente
            Else
                StatusMessage = $"Material '{result.Codigo}' já estava aplicado."
            End If

            Task.Delay(3000).ContinueWith(Sub(t)
                                              If StatusMessage.StartsWith($"Material '{result.Codigo}'") Then
                                                  StatusMessage = originalStatus
                                              End If
                                          End Sub, TaskScheduler.FromCurrentSynchronizationContext())

        Catch ex As Exception
            StatusMessage = "Erro ao aplicar material: " & ex.Message
            If transaction IsNot Nothing AndAlso Not transaction.Aborted Then
                Try : transaction.Abort() : Catch : End Try
            End If
            Task.Delay(5000).ContinueWith(Sub(t)
                                              StatusMessage = originalStatus
                                          End Sub, TaskScheduler.FromCurrentSynchronizationContext())
        Finally
            ' Libera objetos COM locais
            ReleaseComObject(customPropSet)
            ReleaseComObject(activeDoc)
            ReleaseComObject(transaction)
            ' NÃO liberar invApp aqui, pois é obtido de Globals
        End Try
    End Sub

    ' Função auxiliar GetOrAddPropertySet
    Private Function GetOrAddPropertySet(doc As Document, setName As String, internalName As String) As PropertySet
        Dim propSet As PropertySet = Nothing
        Try
            propSet = doc.PropertySets.Item(setName)
        Catch exNotFound As Exception
            Try
                propSet = doc.PropertySets.Add(setName, internalName)
                System.Diagnostics.Debug.WriteLine($"   -> Criado PropertySet '{setName}'.")
            Catch exAdd As Exception
                System.Diagnostics.Debug.WriteLine($"   -> ERRO ao ADICIONAR PropertySet '{setName}': {exAdd.Message}")
                propSet = Nothing
            End Try
        Catch exOther As Exception
            System.Diagnostics.Debug.WriteLine($"   -> ERRO ao OBTER PropertySet '{setName}': {exOther.Message}")
            propSet = Nothing
        End Try
        Return propSet
    End Function


    Private Function SetPropertyValue(propSet As PropertySet, propName As String, value As Object) As Boolean
        If propSet Is Nothing Then Return False

        If value Is Nothing Then value = String.Empty

        Dim existingProp As Inventor.Property = Nothing
        Dim changed As Boolean = False
        Dim propWasCreated As Boolean = False ' Flag para saber se a propriedade foi criada nesta chamada

        Try
            ' Tenta obter a propriedade existente
            Try
                existingProp = propSet.Item(propName)
            Catch exNotFound As Exception
                ' Propriedade não existe, será criada
                existingProp = Nothing
            End Try

            If existingProp IsNot Nothing Then
                ' Propriedade existe, verifica se o valor é diferente
                Dim currentValue As Object = Nothing
                Try
                    currentValue = existingProp.Value
                Catch exRead As Exception
                    System.Diagnostics.Debug.WriteLine($"AVISO: Erro ao ler valor atual de '{propName}': {exRead.Message}")
                    currentValue = Nothing ' Assume que não sabemos o valor
                End Try

                ' Compara como string para simplificar (cuidado com tipos numéricos se precisar de comparação exata)
                If currentValue Is Nothing OrElse Not currentValue.ToString().Equals(value.ToString()) Then
                    existingProp.Value = value
                    changed = True
                End If
            Else
                ' Propriedade não existe, adiciona
                existingProp = propSet.Add(value, propName) ' existingProp agora referencia a nova propriedade
                changed = True
                propWasCreated = True
                System.Diagnostics.Debug.WriteLine($"   -> Criada iProperty '{propName}'.")
            End If

        Catch exSetOrAdd As Exception
            System.Diagnostics.Debug.WriteLine($"ERRO ao definir/adicionar propriedade '{propName}': {exSetOrAdd.Message}")
            changed = False ' Garante que não retorne True se deu erro
        Finally
            ' Libera o objeto COM da propriedade (seja existente ou recém-criada)
            ReleaseComObject(existingProp)
        End Try

        Return changed
    End Function


    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                ' Loop até que a contagem de referência seja 0
                While Marshal.ReleaseComObject(obj) > 0
                End While
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine($"INFO: Erro (geralmente ignorável) ao liberar objeto COM: {ex.Message}")
        Finally
            obj = Nothing ' Define a variável como Nothing após a liberação
        End Try
    End Sub

End Class

' Classe auxiliar RelayCommand
Public Class RelayCommand
    Implements ICommand
    Private ReadOnly _execute As Action(Of Object)
    Private ReadOnly _canExecute As Predicate(Of Object)

    Public Sub New(execute As Action(Of Object), Optional canExecute As Predicate(Of Object) = Nothing)
        If execute Is Nothing Then Throw New ArgumentNullException("execute") ' Garante que execute não seja nulo
        _execute = execute
        _canExecute = canExecute
    End Sub

    Public Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

    Public Function CanExecute(parameter As Object) As Boolean Implements ICommand.CanExecute
        Return If(_canExecute Is Nothing, True, _canExecute(parameter))
    End Function

    Public Sub Execute(parameter As Object) Implements ICommand.Execute
        _execute(parameter)
    End Sub

    Public Sub RaiseCanExecuteChanged()
        ' CORREÇÃO: Especifica o namespace System.Windows
        Dim dispatcher = System.Windows.Application.Current?.Dispatcher
        If dispatcher IsNot Nothing Then
            ' Usar BeginInvoke é geralmente seguro para RaiseCanExecuteChanged
            dispatcher.BeginInvoke(New Action(Sub()
                                                  RaiseEvent CanExecuteChanged(Me, EventArgs.Empty)
                                              End Sub))
        Else
            ' Executar diretamente se não houver dispatcher (cenário raro)
            RaiseEvent CanExecuteChanged(Me, EventArgs.Empty)
        End If
    End Sub
End Class