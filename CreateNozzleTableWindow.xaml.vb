' Arquivo: CreateNozzleTableWindow.xaml.vb (VERSÃO 5.2 - Corrigido SheetMetal e NaturalSort)
Imports Inventor
Imports System.Windows
Imports System.Text
Imports System.Linq
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Imports System.IO ' Para Path

' Estrutura auxiliar para armazenar dados do bocal
Public Structure NozzleInfo
    Public Property TAG As String
    Public Property QTD As String
    Public Property DN As String
    Public Property DESCRIÇÃO As String
    Public Property ELEVAÇÃO As String ' XML
    Public Property OD As String       ' XML
    Public Property ESP_BOCAL As String ' XML
    Public Property NPS_SCH As String   ' XML
    Public Property LARG_PAD As String  ' XML
    Public Property ESP_PAD As String   ' XML
End Structure


Public Class CreateNozzleTableWindow
    Private _logBuilder As New StringBuilder()
    Private _invApp As Inventor.Application

    ' ==================================================================
    ' ALTERAÇÃO 2: P/Invoke para Ordenação Natural (Natural Sort)
    ' Isso fará com que N1, N2, N10 seja ordenado corretamente.
    ' ==================================================================
    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function StrCmpLogicalW(ByVal psz1 As String, ByVal psz2 As String) As Integer
    End Function

    Public Sub New()
        InitializeComponent()
        _invApp = Globals.g_inventorApplication
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub btnExecute_Click(sender As Object, e As RoutedEventArgs)
        _logBuilder.Clear()
        LogTextBox.Clear()
        Log("Iniciando criação/atualização da Tabela de Bocais...")

        Dim oDoc As DrawingDocument = Nothing
        Dim oSheet As Inventor.Sheet = Nothing
        Dim oTransaction As Transaction = Nothing
        Dim nozzleDataList As New List(Of NozzleInfo)
        Dim existingTags As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) ' Ignora case para TAGs
        Dim includeXmlColumns As Boolean = chkIncludeXmlColumns.IsChecked.GetValueOrDefault(False)
        Dim ownerWindow As Window = FindOwnerWindow() ' Helper para definir owner de MessageBoxes

        ' -- Larguras padrão das colunas (em cm) --
        Dim tagWidth As Double = 1.5       ' 15 mm
        Dim qtdWidth As Double = 1.0       ' 10 mm
        Dim dnWidth As Double = 1.5        ' 15 mm
        Dim descWidth As Double = 6.0      ' 60 mm
        ' Larguras XML
        Dim elevacaoWidth As Double = 1.5  ' 15 mm
        Dim odWidth As Double = 1.5        ' 15 mm
        Dim espBocalWidth As Double = 1.5  ' 15 mm
        Dim npsSchWidth As Double = 2.0    ' 20 mm
        Dim largPadWidth As Double = 1.5   ' 15 mm
        Dim espPadWidth As Double = 1.5    ' 15 mm


        Try
            ' 1. Validação do Documento
            oDoc = TryCast(_invApp.ActiveDocument, DrawingDocument)
            If oDoc Is Nothing OrElse oDoc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                Log("ERRO: O documento ativo não é um desenho (.idw).")
                FinalizeLog() : Return
            End If
            Log($"Documento de desenho ativo: {oDoc.DisplayName}")
            oSheet = oDoc.ActiveSheet
            Log($"Folha ativa: {oSheet.Name}")

            ' 2. Verificar Tabela Existente (Sem alterações)
            Dim existingTable As CustomTable = Nothing
            Try
                For Each table As CustomTable In oSheet.CustomTables
                    If table.Title.Trim().Equals("TABELA DE BOCAIS", StringComparison.OrdinalIgnoreCase) Then
                        existingTable = table
                        Exit For
                    End If
                Next
                If existingTable IsNot Nothing Then
                    Log("Tabela 'TABELA DE BOCAIS' existente encontrada.")
                    Dim msgReplace = New MessageBoxWindow("Já existe uma TABELA DE BOCAIS na folha." & vbCrLf & "Deseja substituí-la?", "CITROTEC - Substituir Tabela", ModernMessageBoxButtons.YesNo)
                    If ownerWindow IsNot Nothing Then msgReplace.Owner = ownerWindow Else msgReplace.Owner = Me
                    msgReplace.ShowDialog()

                    If msgReplace.Result = MessageBoxResult.Yes Then
                        Log("Usuário optou por substituir. Tentando capturar larguras e excluir tabela...")
                        Try
                            tagWidth = existingTable.Columns.Item("TAG").Width
                            descWidth = existingTable.Columns.Item("DESCRIÇÃO").Width
                            If includeXmlColumns Then
                                Try : odWidth = existingTable.Columns.Item("OD").Width : Catch : End Try
                                Try : espBocalWidth = existingTable.Columns.Item("ESP. BOCAL").Width : Catch : End Try
                                Try : npsSchWidth = existingTable.Columns.Item("NPS/SCH").Width : Catch : End Try
                                Try : elevacaoWidth = existingTable.Columns.Item("ELEVAÇÃO").Width : Catch : End Try
                                Try : largPadWidth = existingTable.Columns.Item("LARG. PAD").Width : Catch : End Try
                                Try : espPadWidth = existingTable.Columns.Item("ESP. PAD").Width : Catch : End Try
                            Else
                                Try : qtdWidth = existingTable.Columns.Item("QTD").Width : Catch : End Try
                                Try : dnWidth = existingTable.Columns.Item("D.N").Width : Catch : End Try
                            End If
                            Log("Larguras capturadas da tabela existente.")
                        Catch exWidth As Exception
                            Log($"AVISO: Erro ao capturar larguras da tabela existente: {exWidth.Message}. Usando padrões.")
                        End Try
                        existingTable.Delete()
                        Log("Tabela existente excluída.")
                        existingTable = Nothing
                    Else
                        Log("Usuário optou por NÃO substituir. Operação cancelada.")
                        FinalizeLog() : Return
                    End If
                End If
            Catch exCheck As Exception
                Log($"ERRO ao verificar/excluir tabela existente: {exCheck.Message}")
            End Try

            ' 3. Iniciar Transação e Otimização
            oTransaction = _invApp.TransactionManager.StartTransaction(oDoc, "Criar Tabela de Bocais")
            _invApp.SilentOperation = True
            Log("Transação iniciada.")
            Log($"Modo XML selecionado: {includeXmlColumns}")

            ' 4. Coletar Dados de Componentes (Varredura da Montagem)
            '    (Usando a correção ItemByName da v5.1)
            Log("Iniciando varredura da montagem principal...")
            Dim refAssemblyDoc As AssemblyDocument = Nothing
            Dim mainAssemblyFileName As String = ""
            Dim firstViewRefDesc As DocumentDescriptor = Nothing

            If oSheet.DrawingViews.Count > 0 Then
                Dim firstView As DrawingView = Nothing
                Try
                    firstView = oSheet.DrawingViews.Item(1)
                    firstViewRefDesc = firstView.ReferencedDocumentDescriptor
                    If firstViewRefDesc IsNot Nothing AndAlso firstViewRefDesc.ReferencedDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        mainAssemblyFileName = firstViewRefDesc.FullDocumentName
                        Log($"Montagem referenciada encontrada via vista '{firstView.Name}': {System.IO.Path.GetFileName(mainAssemblyFileName)}")
                    Else
                        Log("A primeira vista não referencia uma montagem. Pulando varredura.")
                    End If
                Catch exView As Exception
                    Log($"Erro ao acessar a primeira vista: {exView.Message}. Pulando varredura.")
                Finally
                    ReleaseComObject(firstView)
                    ReleaseComObject(firstViewRefDesc)
                End Try
            Else
                Log("Nenhuma vista encontrada. Pulando varredura.")
            End If

            ' Fallback para PartsList
            If String.IsNullOrEmpty(mainAssemblyFileName) AndAlso oSheet.PartsLists.Count > 0 Then
                Dim oPL As PartsList = Nothing
                Dim plDesc As DocumentDescriptor = Nothing
                Try
                    oPL = oSheet.PartsLists.Item(1)
                    plDesc = oPL.ReferencedDocumentDescriptor
                    If plDesc IsNot Nothing AndAlso plDesc.ReferencedDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        mainAssemblyFileName = plDesc.FullDocumentName
                        Log($"Montagem referenciada encontrada via PartsList '{oPL.Title}': {System.IO.Path.GetFileName(mainAssemblyFileName)}")
                    Else
                        Log("A PartsList não referencia uma montagem.")
                    End If
                Catch exPL As Exception
                    Log($"Erro ao acessar a PartsList: {exPL.Message}.")
                Finally
                    ReleaseComObject(plDesc)
                    ReleaseComObject(oPL)
                End Try
            End If


            If Not String.IsNullOrEmpty(mainAssemblyFileName) Then
                Dim wasOpened As Boolean = False
                Try
                    refAssemblyDoc = TryCast(_invApp.Documents.ItemByName(mainAssemblyFileName), AssemblyDocument)

                    If refAssemblyDoc Is Nothing Then
                        refAssemblyDoc = TryCast(_invApp.Documents.Open(mainAssemblyFileName, False), AssemblyDocument)
                        If refAssemblyDoc IsNot Nothing Then wasOpened = True
                        Log($"Montagem '{refAssemblyDoc?.DisplayName}' aberta invisivelmente.")
                    Else
                        Log($"Montagem '{refAssemblyDoc.DisplayName}' já estava aberta.")
                    End If

                    If refAssemblyDoc IsNot Nothing Then
                        Log("Iniciando varredura recursiva de componentes...")
                        ProcessPhantomReferenceOccurrences(refAssemblyDoc.ComponentDefinition.Occurrences, nozzleDataList, existingTags, includeXmlColumns)
                    Else
                        Log("ERRO: Falha ao obter/abrir montagem.")
                    End If

                Catch exOpenAssy As Exception
                    Log($"ERRO ao abrir/processar montagem '{System.IO.Path.GetFileName(mainAssemblyFileName)}': {exOpenAssy.Message}")
                    Log(exOpenAssy.StackTrace)
                Finally
                    If wasOpened AndAlso refAssemblyDoc IsNot Nothing Then
                        Try : refAssemblyDoc.Close(True) : Log($"Montagem fechada.") : Catch exClose As Exception : Log($"AVISO: Erro ao fechar montagem: {exClose.Message}") : End Try
                    End If
                    ReleaseComObject(refAssemblyDoc)
                End Try
            Else
                Log("ERRO: Nenhuma montagem (.iam) de referência encontrada na folha (via Vista ou PartsList).")
            End If
            Log($"Coleta via varredura de montagem concluída. Total de {nozzleDataList.Count} bocais únicos encontrados.")

            ' 5. Verificar Dados
            If nozzleDataList.Count = 0 Then
                Log("Nenhum bocal válido encontrado na montagem.")
                oTransaction.Abort()
                _invApp.SilentOperation = False
                Dim msgNoData = New MessageBoxWindow("Nenhum bocal válido (com 'TAG_BOCAL') foi encontrado na montagem.", "Nenhum Dado", ModernMessageBoxButtons.OK)
                If ownerWindow IsNot Nothing Then msgNoData.Owner = ownerWindow Else msgNoData.Owner = Me
                msgNoData.ShowDialog()
                FinalizeLog() : Return
            End If

            ' ==================================================================
            ' ALTERAÇÃO 2: Ordenar usando a função de Ordenação Natural
            ' ==================================================================
            ' 6. Ordenar
            Try
                nozzleDataList.Sort(Function(a, b) StrCmpLogicalW(a.TAG, b.TAG))
                Log("Dados ordenados (ordem natural).")
            Catch exSort As Exception
                Log($"AVISO: Erro ao ordenar (revertendo para ordem simples): {exSort.Message}.")
                ' Fallback para ordenação simples em caso de erro
                nozzleDataList.Sort(Function(a, b) String.Compare(a.TAG, b.TAG, StringComparison.OrdinalIgnoreCase))
            End Try

            ' 7. Criar Tabela (Sem alterações)
            Dim oTable As CustomTable = Nothing
            Try
                Log("Criando a tabela...")
                Dim columnTitles As String()
                If includeXmlColumns Then
                    columnTitles = {"TAG", "DESCRIÇÃO", "OD", "ESP. BOCAL", "NPS/SCH", "ELEVAÇÃO", "LARG. PAD", "ESP. PAD"}
                Else
                    columnTitles = {"TAG", "QTD", "D.N", "DESCRIÇÃO"}
                End If
                Dim startPoint As Point2d = _invApp.TransientGeometry.CreatePoint2d(1, oSheet.Height - 1)

                oTable = oSheet.CustomTables.Add("TABELA DE BOCAIS", startPoint, columnTitles.Length, nozzleDataList.Count, columnTitles)
                Log($"Tabela criada com {columnTitles.Length} cols / {nozzleDataList.Count} linhas.")

                Log("Preenchendo dados...")
                For i As Integer = 0 To nozzleDataList.Count - 1
                    Dim rowData As NozzleInfo = nozzleDataList(i)
                    Dim currentRow = oTable.Rows.Item(i + 1) ' Obtém a linha (base 1)

                    currentRow.Item("TAG").Value = rowData.TAG
                    currentRow.Item("DESCRIÇÃO").Value = rowData.DESCRIÇÃO

                    If includeXmlColumns Then
                        currentRow.Item("OD").Value = rowData.OD
                        currentRow.Item("ESP. BOCAL").Value = rowData.ESP_BOCAL
                        currentRow.Item("NPS/SCH").Value = rowData.NPS_SCH
                        currentRow.Item("ELEVAÇÃO").Value = rowData.ELEVAÇÃO
                        currentRow.Item("LARG. PAD").Value = rowData.LARG_PAD
                        currentRow.Item("ESP. PAD").Value = rowData.ESP_PAD
                    Else
                        currentRow.Item("QTD").Value = rowData.QTD
                        currentRow.Item("D.N").Value = rowData.DN
                    End If
                    ReleaseComObject(currentRow)
                Next
                Log("Dados preenchidos.")

                Log("Aplicando larguras...")
                Try
                    oTable.Columns.Item("TAG").Width = tagWidth
                    oTable.Columns.Item("DESCRIÇÃO").Width = descWidth
                    If includeXmlColumns Then
                        oTable.Columns.Item("OD").Width = odWidth
                        oTable.Columns.Item("ESP. BOCAL").Width = espBocalWidth
                        oTable.Columns.Item("NPS/SCH").Width = npsSchWidth
                        oTable.Columns.Item("ELEVAÇÃO").Width = elevacaoWidth
                        oTable.Columns.Item("LARG. PAD").Width = largPadWidth
                        oTable.Columns.Item("ESP. PAD").Width = espPadWidth
                    Else
                        oTable.Columns.Item("QTD").Width = qtdWidth
                        oTable.Columns.Item("D.N").Width = dnWidth
                    End If
                    Log("Larguras aplicadas.")
                Catch exApplyWidth As Exception
                    Log($"AVISO: Erro ao aplicar larguras: {exApplyWidth.Message}")
                End Try

                Try
                    oTable.Style = oDoc.StylesManager.TableStyles.Item("TABELA_BOCAIS_STYLE")
                    Log("Estilo 'TABELA_BOCAIS_STYLE' aplicado.")
                Catch exStyle As Exception
                    Log($"AVISO: Estilo 'TABELA_BOCAIS_STYLE' não encontrado: {exStyle.Message}")
                End Try

            Catch exCreateTable As Exception
                Log($"ERRO CRÍTICO ao criar/preencher tabela: {exCreateTable.Message}")
                Throw
            Finally
                ReleaseComObject(oTable)
            End Try

            ' 8. Finalizar (Sem alterações)
            oDoc.Update()
            oTransaction.End()
            _invApp.SilentOperation = False
            Log("Transação finalizada.")
            Log("-------------------------------------")
            Log("Tabela de Bocais criada/atualizada com sucesso.")

            Dim msgSuccessFinal = New MessageBoxWindow("Tabela de Bocais criada/atualizada com sucesso!", "Operação Concluída", ModernMessageBoxButtons.OK)
            If ownerWindow IsNot Nothing Then msgSuccessFinal.Owner = ownerWindow Else msgSuccessFinal.Owner = Me
            msgSuccessFinal.ShowDialog()

        Catch ex As Exception
            Log("======================================")
            Log("ERRO GERAL INESPERADO: " & ex.Message)
            Log(ex.StackTrace)
            Log("======================================")
            If oTransaction IsNot Nothing AndAlso oTransaction.Aborted = False Then
                Try : oTransaction.Abort() : Log("Transação abortada.") : Catch exAbort As Exception : Log($"Erro ao abortar: {exAbort.Message}") : End Try
            End If
            _invApp.SilentOperation = False

            Dim ownerWindowError As Window = FindOwnerWindow()
            Dim msgEx = New MessageBoxWindow($"Ocorreu um erro: {vbCrLf}{ex.Message}", "Erro", ModernMessageBoxButtons.OK)
            If ownerWindowError IsNot Nothing Then msgEx.Owner = ownerWindowError Else msgEx.Owner = Me
            msgEx.ShowDialog()

        Finally
            _invApp.SilentOperation = False
            FinalizeLog()
            ReleaseComObject(oTransaction)
        End Try
    End Sub

    ' Função Recursiva para processar ocorrências
    Private Sub ProcessPhantomReferenceOccurrences(oOccurrences As ComponentOccurrences,
                                                  ByRef nozzleDataList As List(Of NozzleInfo),
                                                  ByRef existingTags As HashSet(Of String),
                                                  includeXmlColumns As Boolean)

        If oOccurrences Is Nothing OrElse oOccurrences.Count = 0 Then Return
        Dim oOcc As ComponentOccurrence = Nothing
        Dim compDef As ComponentDefinition = Nothing
        Dim doc As Document = Nothing
        Dim propSetCustom As PropertySet = Nothing
        Dim propSetInstance As PropertySet = Nothing

        For Each oOcc In oOccurrences
            If oOcc.Suppressed Then Continue For
            compDef = Nothing : doc = Nothing : propSetCustom = Nothing : propSetInstance = Nothing
            Try
                compDef = oOcc.Definition
                If compDef Is Nothing Then Continue For

                ' Varre sub-ocorrências PRIMEIRO (busca recursiva)
                If oOcc.SubOccurrences IsNot Nothing AndAlso oOcc.SubOccurrences.Count > 0 Then
                    ProcessPhantomReferenceOccurrences(oOcc.SubOccurrences, nozzleDataList, existingTags, includeXmlColumns)
                End If

                ' ==================================================================
                ' ALTERAÇÃO 1: Incluir Peças Sheet Metal (kSheetMetalComponentDefinitionObject)
                ' ==================================================================
                If compDef.Type <> ObjectTypeEnum.kPartComponentDefinitionObject AndAlso
                   compDef.Type <> ObjectTypeEnum.kSheetMetalComponentDefinitionObject Then

                    ' Se não for uma PEÇA (Regular ou SheetMetal), pula para a próxima
                    Continue For
                End If

                doc = compDef.Document
                If doc Is Nothing Then Continue For

                Dim docName As String = "N/A"
                Try : docName = System.IO.Path.GetFileName(doc.FullFileName) : Catch : End Try
                Log($" -> Verificando: '{oOcc.Name}' ({docName}) [Tipo: {compDef.Type.ToString()}]")

                Dim tagBocal As String = "N/A"
                Dim data As New NozzleInfo()

                ' 1. Tenta ler iProperties de Instância (se habilitado)
                Try
                    If oOcc.OccurrencePropertySetsEnabled Then
                        propSetInstance = oOcc.OccurrencePropertySets.Item(1) ' "User Defined" de Instância
                        tagBocal = GetPropertyValue(propSetInstance, "TAG_BOCAL", "N/A")
                        Log($"    -> Tentativa Leitura Instance Prop TAG_BOCAL: '{tagBocal}'")
                        If tagBocal <> "N/A" AndAlso Not String.IsNullOrEmpty(tagBocal) Then
                            data.QTD = GetPropertyValue(propSetInstance, "QTD")
                            data.DN = GetPropertyValue(propSetInstance, "D.N")
                            data.DESCRIÇÃO = GetPropertyValue(propSetInstance, "DESCRICAO_BOCAL")?.ToUpper()
                            If includeXmlColumns Then
                                data.ELEVAÇÃO = GetPropertyValue(propSetInstance, "Elevação")
                                data.OD = GetPropertyValue(propSetInstance, "OD")
                                data.ESP_BOCAL = GetPropertyValue(propSetInstance, "ESP_BOCAL")
                                data.NPS_SCH = GetPropertyValue(propSetInstance, "NPS/SCH")
                                data.LARG_PAD = GetPropertyValue(propSetInstance, "PAD_LARG")
                                data.ESP_PAD = GetPropertyValue(propSetInstance, "PAD_ESP")
                            End If
                            Log($"    -> Lidos de Instance Props.")
                        Else
                            Log($"    -> Instance Props sem TAG_BOCAL válida.")
                            tagBocal = "N/A" ' Reseta se for ""
                        End If
                    Else
                        Log($"    -> OccurrencePropertySetsEnabled = False.")
                    End If
                Catch exInst As Exception : Log($"    -> Erro Instance Props: {exInst.Message}") : tagBocal = "N/A" : End Try

                ' 2. Se não achou na Instância, tenta ler iProperties Customizadas (da Peça)
                If tagBocal = "N/A" Then
                    Try
                        propSetCustom = doc.PropertySets.Item("Inventor User Defined Properties")
                        tagBocal = GetPropertyValue(propSetCustom, "TAG_BOCAL", "N/A")
                        Log($"    -> Tentativa Leitura Custom Prop TAG_BOCAL: '{tagBocal}'")
                        If tagBocal <> "N/A" AndAlso Not String.IsNullOrEmpty(tagBocal) Then
                            data.QTD = GetPropertyValue(propSetCustom, "QTD")
                            data.DN = GetPropertyValue(propSetCustom, "D.N")
                            data.DESCRIÇÃO = GetPropertyValue(propSetCustom, "DESCRICAO_BOCAL")?.ToUpper()
                            If includeXmlColumns Then
                                data.ELEVAÇÃO = GetPropertyValue(propSetCustom, "Elevação")
                                data.OD = GetPropertyValue(propSetCustom, "OD")
                                data.ESP_BOCAL = GetPropertyValue(propSetCustom, "ESP_BOCAL")
                                data.NPS_SCH = GetPropertyValue(propSetCustom, "NPS/SCH")
                                data.LARG_PAD = GetPropertyValue(propSetCustom, "PAD_LARG")
                                data.ESP_PAD = GetPropertyValue(propSetCustom, "PAD_ESP")
                            End If
                            Log($"    -> Lidos de Custom Props.")
                        Else
                            Log($"    -> Sem TAG_BOCAL válida em Custom Props.")
                            tagBocal = "N/A" ' Reseta se for ""
                        End If
                    Catch exCust As ArgumentException
                        Log($"    -> 'Inventor User Defined Properties' não existe em '{docName}'.")
                        tagBocal = "N/A"
                    Catch exCust As Exception : Log($"    -> Erro Custom Props: {exCust.Message}") : tagBocal = "N/A" : End Try
                End If

                ' 3. Adiciona na lista se for válido e único
                If tagBocal <> "N/A" AndAlso tagBocal <> "[SEM ITEM]" Then
                    If existingTags.Add(tagBocal) Then
                        data.TAG = tagBocal
                        nozzleDataList.Add(data)
                        Log($" -> ADICIONADO Bocal TAG='{tagBocal}'.")
                    Else
                        Log($" -> Bocal TAG='{tagBocal}' já adicionado. Pulando.")
                    End If
                Else
                    Log($" -> Peça '{oOcc.Name}' não tem TAG_BOCAL válida. Ignorando.")
                End If

            Catch exOcc As Exception
                Dim occName As String = If(oOcc IsNot Nothing, oOcc.Name, "N/A")
                Log($" ## ERRO ao processar ocorrência '{occName}': {exOcc.Message}")
            Finally
                ReleaseComObject(propSetInstance)
                ReleaseComObject(propSetCustom)
                ReleaseComObject(doc)
                ReleaseComObject(compDef)
                ReleaseComObject(oOcc) ' Libera a ocorrência atual
            End Try
        Next
    End Sub

    ' --- Funções Auxiliares (Sem alterações) ---

    Private Function GetPropertyValue(propSet As PropertySet, propName As String, Optional defaultValue As String = "") As String
        If propSet Is Nothing Then Return defaultValue
        Dim prop As Inventor.Property = Nothing
        Dim valueString As String = defaultValue
        Try
            prop = propSet.Item(propName)
            If prop.Value IsNot Nothing Then
                valueString = prop.Value.ToString()
            End If
            Return valueString
        Catch ex As Exception
            Return defaultValue
        Finally
            ReleaseComObject(prop)
        End Try
    End Function

    Private Function FindOwnerWindow() As Window
        Dim owner As Window = Nothing
        Try : owner = System.Windows.Application.Current.Windows.OfType(Of JanelaPrincipal).FirstOrDefault() : Catch : End Try
        If owner Is Nothing Then owner = Me
        Return owner
    End Function

    Private Sub Log(message As String)
        _logBuilder.AppendLine($"{DateTime.Now:HH:mm:ss} - {message}")
        Me.Dispatcher.InvokeAsync(Sub()
                                      If LogTextBox IsNot Nothing Then
                                          LogTextBox.Text = _logBuilder.ToString()
                                          LogTextBox.ScrollToEnd()
                                      End If
                                  End Sub)
    End Sub

    Private Sub FinalizeLog()
        Me.Dispatcher.InvokeAsync(Sub()
                                      If LogTextBox IsNot Nothing Then
                                          LogTextBox.Text = _logBuilder.ToString()
                                          LogTextBox.ScrollToEnd()
                                      End If
                                  End Sub)
    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try : If obj IsNot Nothing Then : While Marshal.ReleaseComObject(obj) > 0 : End While : End If : Catch : Finally : obj = Nothing : End Try
    End Sub

End Class