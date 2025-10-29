' Arquivo: PrepareNozzlePartWindow.xaml.vb (VERSÃO 5 - Modo Dual Peça/Montagem)
Imports Inventor
Imports System.Windows
Imports System.Text
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.IO

Public Class PrepareNozzlePartWindow
    Private _logBuilder As New StringBuilder()
    Private _invApp As Inventor.Application
    Private _activePartDoc As PartDocument = Nothing ' Armazena o documento de peça
    Private _docType As DocumentTypeEnum

    ' Nomes das propriedades a serem gerenciadas
    Private ReadOnly _propNames As New List(Of String) From {"TAG_BOCAL", "DESCRICAO_BOCAL", "D.N", "QTD"}

    Public Sub New()
        InitializeComponent()
        _invApp = Globals.g_inventorApplication
    End Sub

    ' ==================================================================
    ' LÓGICA DE INICIALIZAÇÃO (MODO DUAL)
    ' ==================================================================

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        _logBuilder.Clear()
        LogTextBox.Clear()

        Dim activeDoc As Document = _invApp.ActiveDocument
        If activeDoc Is Nothing Then
            Log("ERRO: Nenhum documento ativo.") : FinalizeLog() : Return
        End If

        _docType = activeDoc.DocumentType

        If _docType = DocumentTypeEnum.kPartDocumentObject Then
            ' --- MODO PEÇA ---
            Log("Modo Peça (.ipt) ativado.")
            PartInputGrid.Visibility = Visibility.Visible
            PartButtonPanel.Visibility = Visibility.Visible
            AssemblyButtonPanel.Visibility = Visibility.Collapsed
            LoadInitialValues() ' Roda a lógica original de carregamento de peça

        ElseIf _docType = DocumentTypeEnum.kAssemblyDocumentObject Then
            ' --- MODO MONTAGEM ---
            Log("Modo Montagem (.iam) ativado.")
            Log("Pronto para LIMPAR propriedades de bocais de todas as peças.")

            TitleTextBlock.Text = "Limpar Propriedades da Montagem"
            DescriptionTextBlock.Text = "Este comando varrerá a montagem ativa e limpará as iProperties de bocal (TAG_BOCAL, D.N, etc.) de TODAS as peças (incluindo sub-conjuntos, phantom, etc.)"

            PartInputGrid.Visibility = Visibility.Collapsed
            PartButtonPanel.Visibility = Visibility.Collapsed
            AssemblyButtonPanel.Visibility = Visibility.Visible

            FinalizeLog()
        Else
            ' --- MODO INVÁLIDO ---
            Log($"ERRO: Documento ativo não é Peça ou Montagem ({activeDoc.DisplayName}).")
            PartInputGrid.Visibility = Visibility.Collapsed
            PartButtonPanel.Visibility = Visibility.Collapsed
            AssemblyButtonPanel.Visibility = Visibility.Collapsed
            Log("Abra uma Peça (.ipt) ou Montagem (.iam).")
            FinalizeLog()
        End If
    End Sub

    ' ==================================================================
    ' LÓGICA DO MODO MONTAGEM (NOVO)
    ' ==================================================================

    ' Evento do novo botão "Limpar Propriedades da Montagem"
    Private Sub btnCleanseAssembly_Click(sender As Object, e As RoutedEventArgs)
        Log("Iniciando limpeza de propriedades na montagem...")

        Dim oAssyDoc As AssemblyDocument = TryCast(_invApp.ActiveDocument, AssemblyDocument)
        If oAssyDoc Is Nothing Then
            Log("ERRO: Documento ativo não é uma montagem. Operação cancelada.")
            FinalizeLog() : Return
        End If

        Dim transaction As Transaction = Nothing
        ' HashSet para logar apenas uma vez cada peça modificada
        Dim cleanedParts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        btnCleanseAssembly.IsEnabled = False

        Try
            transaction = _invApp.TransactionManager.StartTransaction(oAssyDoc, "Limpar Propriedades de Bocais")
            Log("Transação iniciada.")
            _invApp.SilentOperation = True

            ' Inicia a varredura recursiva
            ProcessOccurrencesForClearing(oAssyDoc.ComponentDefinition.Occurrences, cleanedParts)

            _invApp.SilentOperation = False
            transaction.End()
            Log("Transação concluída.")

            Log("-------------------------------------")
            If cleanedParts.Count = 0 Then
                Log("Limpeza concluída. Nenhuma peça com propriedades de bocal foi encontrada/modificada.")
                Dim msgNoChange = New MessageBoxWindow("Nenhuma peça com propriedades de bocal foi encontrada.", "Limpeza Concluída", ModernMessageBoxButtons.OK)
                msgNoChange.Owner = Me
                msgNoChange.ShowDialog()
            Else
                Log($"Limpeza concluída. {cleanedParts.Count} peças foram limpas:")
                ' Ordena a lista de peças para o log
                For Each partFile As String In cleanedParts.OrderBy(Function(f) f)
                    Log($" -> {IO.Path.GetFileName(partFile)}")
                Next
                Dim msgSuccess = New MessageBoxWindow($"{cleanedParts.Count} peças tiveram suas propriedades de bocal limpas.", "Limpeza Concluída", ModernMessageBoxButtons.OK)
                msgSuccess.Owner = Me
                msgSuccess.ShowDialog()
            End If

        Catch ex As Exception
            Log("======================================")
            Log("ERRO GERAL NA LIMPEZA: " & ex.Message)
            Log(ex.StackTrace)
            Log("======================================")
            If transaction IsNot Nothing AndAlso transaction.Aborted = False Then
                Try : transaction.Abort() : Log("Transação abortada devido a erro.") : Catch : End Try
            End If
            _invApp.SilentOperation = False
        Finally
            btnCleanseAssembly.IsEnabled = True
            FinalizeLog()
            ReleaseComObject(transaction)
        End Try
    End Sub

    ' Nova Função Recursiva para limpar propriedades
    Private Sub ProcessOccurrencesForClearing(oOccurrences As ComponentOccurrences, ByRef cleanedParts As HashSet(Of String))
        If oOccurrences Is Nothing OrElse oOccurrences.Count = 0 Then Return

        Dim oOcc As ComponentOccurrence = Nothing
        Dim compDef As ComponentDefinition = Nothing
        Dim doc As Document = Nothing
        Dim propSetCustom As PropertySet = Nothing
        Dim prop As Inventor.Property = Nothing

        For Each oOcc In oOccurrences
            If oOcc.Suppressed Then Continue For
            compDef = Nothing : doc = Nothing : propSetCustom = Nothing : prop = Nothing

            Try
                compDef = oOcc.Definition
                If compDef Is Nothing Then Continue For

                ' 1. Recurse (Varre sub-montagens primeiro)
                If oOcc.SubOccurrences IsNot Nothing AndAlso oOcc.SubOccurrences.Count > 0 Then
                    ProcessOccurrencesForClearing(oOcc.SubOccurrences, cleanedParts)
                End If

                ' 2. Processa a ocorrência atual (se for peça ou sheet metal)
                If compDef.Type <> ObjectTypeEnum.kPartComponentDefinitionObject AndAlso
                   compDef.Type <> ObjectTypeEnum.kSheetMetalComponentDefinitionObject Then
                    Continue For
                End If

                doc = compDef.Document
                If doc Is Nothing Then Continue For

                ' 3. Tenta obter o PropertySet
                Try
                    propSetCustom = doc.PropertySets.Item("Inventor User Defined Properties")
                Catch exNoSet As Exception
                    ' Peça não tem o property set, então não tem o que limpar. Pula.
                    Continue For
                End Try

                ' 4. Itera pelas propriedades de bocal e as limpa
                Dim changesMade As Boolean = False
                For Each propName As String In _propNames ' Usa a lista global
                    prop = Nothing
                    Try
                        ' Tenta obter a propriedade
                        prop = propSetCustom.Item(propName)
                        ' Se ela existe e seu valor não é vazio, limpa
                        If prop.Value IsNot Nothing AndAlso prop.Value.ToString() <> "" Then
                            prop.Value = ""
                            changesMade = True
                        End If
                    Catch exNoProp As Exception
                        ' Propriedade não existe, nada a limpar. Continua para a próxima.
                    Finally
                        ReleaseComObject(prop)
                    End Try
                Next

                ' 5. Se alguma mudança foi feita, adiciona ao log
                If changesMade Then
                    cleanedParts.Add(doc.FullFileName)
                End If

            Catch exOcc As Exception
                ' Ignora erros em ocorrências individuais para não parar o processo
            Finally
                ReleaseComObject(propSetCustom)
                ReleaseComObject(doc)
                ReleaseComObject(compDef)
                ReleaseComObject(oOcc)
            End Try
        Next
    End Sub

    ' Handler do botão "Fechar" do modo montagem
    Private Sub btnCloseAssembly_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub


    ' ==================================================================
    ' LÓGICA DO MODO PEÇA (ORIGINAL)
    ' ==================================================================

    ' Carrega os valores atuais das iProperties nos TextBoxes
    Private Sub LoadInitialValues()
        ' (Função original - levemente modificada para se encaixar no novo Window_Loaded)
        Log("Carregando valores atuais das propriedades da Peça...")
        btnApply.IsEnabled = False ' Desabilita o botão até carregar

        Dim customPropSet As PropertySet = Nothing

        Try
            ' 1. Validação do Documento (Já validado como Peça no Window_Loaded)
            _activePartDoc = TryCast(_invApp.ActiveDocument, PartDocument)
            If _activePartDoc Is Nothing Then
                Log("ERRO: O documento ativo não é uma Peça (.ipt).")
                FinalizeLog() : Return
            End If
            Log($"Documento de peça ativo: {_activePartDoc.DisplayName}")

            ' 2. Obter PropertySet (sem criar ainda)
            Try
                customPropSet = _activePartDoc.PropertySets.Item("Inventor User Defined Properties")
            Catch ex As Exception
                Log("AVISO: PropertySet 'Inventor User Defined Properties' não encontrado. As propriedades serão criadas ao aplicar.")
                txtTagBocal.Text = ""
                txtDescricaoBocal.Text = ""
                txtDN.Text = ""
                txtQtd.Text = "1"
                btnApply.IsEnabled = True
                FinalizeLog()
                Return
            End Try

            ' 3. Ler valores existentes
            Log("Lendo valores existentes...")
            txtTagBocal.Text = GetPropertyValue(customPropSet, "TAG_BOCAL", "")
            txtDescricaoBocal.Text = GetPropertyValue(customPropSet, "DESCRICAO_BOCAL", "")
            txtDN.Text = GetPropertyValue(customPropSet, "D.N", "")
            txtQtd.Text = GetPropertyValue(customPropSet, "QTD", "1")

            Log("Valores carregados nos campos.")
            btnApply.IsEnabled = True

        Catch ex As Exception
            Log("======================================")
            Log("ERRO INESPERADO AO CARREGAR VALORES: " & ex.Message)
            Log(ex.StackTrace)
            Log("======================================")
            Try : btnApply.IsEnabled = False : Catch : End Try
        Finally
            FinalizeLog()
            ReleaseComObject(customPropSet)
        End Try
    End Sub

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub btnClear_Click(sender As Object, e As RoutedEventArgs)
        Log("Campos do formulário foram limpos.")
        txtTagBocal.Text = ""
        txtDescricaoBocal.Text = ""
        txtDN.Text = ""
        txtQtd.Text = "1" ' Manter o padrão de 1 para QTD
        FinalizeLog()
    End Sub

    ' Lógica do botão "Aplicar" (Modo Peça)
    Private Sub btnApply_Click(sender As Object, e As RoutedEventArgs)
        Log("Iniciando aplicação dos valores na Peça...")

        ' RE-OBTER documento ativo
        _activePartDoc = TryCast(_invApp.ActiveDocument, PartDocument)
        If _activePartDoc Is Nothing OrElse _activePartDoc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
            Log("ERRO: Nenhuma peça ativa válida. Recarregue a peça e tente novamente.")
            FinalizeLog()
            MessageBox.Show("Nenhuma peça (.ipt) ativa. Abra/ative a peça e tente novamente.", "Erro", MessageBoxButton.OK, MessageBoxImage.Warning)
            Return
        End If

        Dim customPropSet As PropertySet = Nothing
        Dim transaction As Transaction = Nothing
        Dim changesMade As Integer = 0
        Dim transactionStarted As Boolean = False

        Try
            ' 1. Iniciar Transação
            transaction = _invApp.TransactionManager.StartTransaction(_activePartDoc, "Aplicar Propriedades Bocal")
            transactionStarted = True
            Log("Transação iniciada.")

            ' 2. Obter/Criar PropertySet
            customPropSet = GetOrAddPropertySet(_activePartDoc, "Inventor User Defined Properties", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            If customPropSet Is Nothing Then
                Throw New Exception("Não foi possível obter ou criar o conjunto de propriedades 'Inventor User Defined Properties'. O documento pode ser somente leitura.")
            End If

            ' 3. Mapear e Aplicar Valores
            Dim valueMap As New Dictionary(Of System.Windows.Controls.TextBox, String) From {
                {txtTagBocal, "TAG_BOCAL"},
                {txtDescricaoBocal, "DESCRICAO_BOCAL"},
                {txtDN, "D.N"},
                {txtQtd, "QTD"}
            }

            For Each kvp As KeyValuePair(Of System.Windows.Controls.TextBox, String) In valueMap
                Dim textBox As System.Windows.Controls.TextBox = kvp.Key
                Dim propName As String = kvp.Value
                Dim newValue As String = If(textBox IsNot Nothing, textBox.Text, "")

                Dim prop As Inventor.Property = Nothing
                Try
                    prop = GetOrAddProperty(customPropSet, propName, "")
                    If prop Is Nothing Then
                        Log($" -> ERRO: Falha ao obter/criar propriedade '{propName}'. Pulando.")
                        Continue For
                    End If

                    Dim currentValue As String = ""
                    Try
                        If prop.Value IsNot Nothing Then
                            currentValue = prop.Value.ToString()
                        Else
                            currentValue = ""
                        End If
                    Catch exRead As Exception
                        Log($" -> AVISO: Não foi possível ler o valor atual de '{propName}': {exRead.Message}")
                        currentValue = ""
                    End Try

                    If currentValue <> newValue Then
                        prop.Value = newValue
                        Log($" -> Propriedade '{propName}' atualizada para '{newValue}'.")
                        changesMade += 1
                    Else
                        Log($" -> Propriedade '{propName}' já possuía o valor '{newValue}'. Nenhuma alteração.")
                    End If

                Catch exSet As Exception
                    Log($" -> ERRO ao definir valor para '{propName}': {exSet.Message}")
                Finally
                    ReleaseComObject(prop)
                End Try
            Next

            ' 4. Encerrar transação
            If transactionStarted AndAlso transaction IsNot Nothing Then
                transaction.End()
                Log("Transação concluída.")
            End If

            Log("-------------------------------------")
            If changesMade > 0 Then
                Log($"{changesMade} propriedades foram atualizadas/criadas.")
                Try
                    _activePartDoc.Update()
                Catch exUp As Exception
                    Log($"AVISO: Erro ao atualizar documento: {exUp.Message}")
                End Try
                Try
                    ' *** CORREÇÃO DO ERRO 3 (BC30451) ***
                    _activePartDoc.Save()
                Catch exSave As Exception
                    Log($"AVISO: Erro ao salvar documento: {exSave.Message}")
                End Try

                Dim msgSuccess = New MessageBoxWindow($"Valores aplicados com sucesso! {changesMade} propriedades foram atualizadas/criadas.", "Valores Aplicados", ModernMessageBoxButtons.OK)
                msgSuccess.Owner = Me
                msgSuccess.ShowDialog()
            Else
                Log("Nenhuma alteração de valor detectada. Nenhuma propriedade foi modificada.")
                Dim msgNoChange = New MessageBoxWindow("Nenhum valor foi alterado.", "Sem Alterações", ModernMessageBoxButtons.OK)
                msgNoChange.Owner = Me
                msgNoChange.ShowDialog()
            End If

        Catch ex As Exception
            Log("======================================")
            Log("ERRO INESPERADO AO APLICAR VALORES: " & ex.Message)
            Log(ex.StackTrace)
            Log("======================================")
            If transactionStarted AndAlso transaction IsNot Nothing Then
                Try
                    transaction.Abort()
                    Log("Transação abortada devido a erro.")
                Catch exAbort As Exception
                    Log($"Erro secundário ao tentar abortar transação: {exAbort.Message}")
                End Try
            End If

            Dim msgEx = New MessageBoxWindow($"Ocorreu um erro ao aplicar os valores: {vbCrLf}{ex.Message}", "Erro na Operação", ModernMessageBoxButtons.OK)
            msgEx.Owner = Me
            msgEx.ShowDialog()

        Finally
            FinalizeLog()
            ReleaseComObject(customPropSet)
            ReleaseComObject(transaction)
        End Try
    End Sub

    ' ==================================================================
    ' FUNÇÕES AUXILIARES (LOG, COM, IPROPERTY)
    ' ==================================================================

    Private Function GetOrAddPropertySet(doc As Document, setName As String, internalName As String) As PropertySet
        Dim propSet As PropertySet = Nothing
        Try
            propSet = doc.PropertySets.Item(setName)
        Catch exNotFound As Exception
            Try
                propSet = doc.PropertySets.Add(setName, internalName)
                Log($"   -> Criado PropertySet '{setName}'.")
            Catch exAdd As Exception
                Log($"   -> ERRO ao ADICIONAR PropertySet '{setName}': {exAdd.Message}")
                propSet = Nothing
            End Try
        Catch exOther As Exception
            Log($"   -> ERRO ao OBTER PropertySet '{setName}': {exOther.Message}")
            propSet = Nothing
        End Try
        Return propSet
    End Function

    Private Function GetOrAddProperty(propSet As PropertySet, propName As String, defaultValue As Object) As Inventor.Property
        If propSet Is Nothing Then Return Nothing
        Dim prop As Inventor.Property = Nothing
        Try
            prop = propSet.Item(propName)
        Catch exNotFound As Exception
            Try
                prop = propSet.Add(defaultValue, propName)
                Log($"   -> Criada iProperty '{propName}'.")
            Catch exAdd As Exception
                Log($"   -> ERRO ao ADICIONAR iProperty '{propName}': {exAdd.Message}")
                prop = Nothing
            End Try
        Catch exOther As Exception
            Log($"   -> ERRO ao OBTER iProperty '{propName}': {exOther.Message}")
            prop = Nothing
        End Try
        Return prop
    End Function

    Private Function GetPropertyValue(propSet As PropertySet, propName As String, Optional defaultValue As String = "") As String
        If propSet Is Nothing Then Return defaultValue
        Dim prop As Inventor.Property = Nothing
        Try
            prop = propSet.Item(propName)
            If prop.Value IsNot Nothing Then
                Return prop.Value.ToString()
            Else
                Return defaultValue
            End If
        Catch ex As Exception
            Return defaultValue
        Finally
            ReleaseComObject(prop)
        End Try
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
        Try
            If obj IsNot Nothing Then
                While Marshal.ReleaseComObject(obj) > 0
                End While
            End If
        Catch ex As Exception
            Debug.WriteLine($"INFO: Erro (geralmente ignorável) ao liberar objeto COM: {ex.Message}")
        Finally
            obj = Nothing
        End Try
    End Sub

End Class