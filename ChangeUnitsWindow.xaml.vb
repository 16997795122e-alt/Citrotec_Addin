' Arquivo: ChangeUnitsWindow.xaml.vb
Imports Inventor
Imports System.Windows
Imports System.Text
Imports System.Linq
Imports System.Runtime.InteropServices ' Para ReleaseComObject

Public Class ChangeUnitsWindow
    Private _logBuilder As New StringBuilder()
    Private _invApp As Inventor.Application

    Public Sub New()
        InitializeComponent()
        _invApp = Globals.g_inventorApplication ' Pega a instância global do Inventor
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
        Log("Iniciando alteração de unidades...")

        Dim activeDoc As Document = _invApp.ActiveDocument
        Dim oTransaction As Transaction = Nothing
        Dim isMetric As Boolean = rbMetric.IsChecked.GetValueOrDefault(True)

        Try
            ' 1. Validar Documento Ativo
            If activeDoc Is Nothing OrElse Not (activeDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject OrElse activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject) Then
                Log("ERRO: Este comando só pode ser executado em um documento de Peça (.ipt) ou Montagem (.iam) ativo.")
                FinalizeLog()
                Return
            End If
            Log($"Documento ativo: {activeDoc.DisplayName}")

            ' 2. Definir Unidades Alvo
            Dim targetLengthUnits As UnitsTypeEnum
            Dim targetMassUnits As UnitsTypeEnum

            If isMetric Then
                targetLengthUnits = UnitsTypeEnum.kMillimeterLengthUnits
                targetMassUnits = UnitsTypeEnum.kKilogramMassUnits
                Log("Sistema selecionado: Métrico (mm/kg)")
            Else
                targetLengthUnits = UnitsTypeEnum.kInchLengthUnits
                targetMassUnits = UnitsTypeEnum.kLbMassMassUnits
                Log("Sistema selecionado: Polegadas (in/lb)")
            End If

            ' 3. Iniciar Transação
            oTransaction = _invApp.TransactionManager.StartTransaction(activeDoc, "Alterar Unidades Documento")
            _invApp.SilentOperation = True ' Desliga atualizações de tela para performance
            Log("Transação iniciada.")

            ' 4. Alterar Unidades do Documento Ativo
            Log($"Alterando unidades do documento ativo...")
            If activeDoc.UnitsOfMeasure.LengthUnits <> targetLengthUnits OrElse activeDoc.UnitsOfMeasure.MassUnits <> targetMassUnits Then
                activeDoc.UnitsOfMeasure.LengthUnits = targetLengthUnits
                activeDoc.UnitsOfMeasure.MassUnits = targetMassUnits
                activeDoc.Update()
                Log(" -> Unidades do documento ativo alteradas.")
            Else
                Log(" -> Unidades do documento ativo já estavam corretas.")
            End If


            ' 5. Alterar Unidades dos Documentos Referenciados
            Log("Verificando documentos referenciados...")
            Dim refDoc As Document = Nothing
            Dim changedRefCount As Integer = 0
            Dim processedDocs As New HashSet(Of String) ' Para evitar processar o mesmo doc múltiplas vezes (caso raro)

            ' Usar uma cópia da coleção para iterar pode ser mais seguro se a estrutura mudar
            Dim allRefs As New List(Of Document)
            Try
                For Each d As Document In activeDoc.AllReferencedDocuments
                    allRefs.Add(d)
                Next
            Catch ex As Exception
                Log($"AVISO: Erro ao listar todos os documentos referenciados: {ex.Message}")
            End Try


            For Each refDoc In allRefs
                If refDoc Is Nothing Then Continue For
                Dim docPath As String = ""
                Try
                    docPath = refDoc.FullFileName ' Obter caminho para checar duplicados
                    If processedDocs.Contains(docPath) Then Continue For

                    ' Verificar se é uma peça ou montagem
                    If refDoc.DocumentType = DocumentTypeEnum.kPartDocumentObject OrElse refDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        ' Verifica se as unidades já estão corretas
                        If refDoc.UnitsOfMeasure.LengthUnits <> targetLengthUnits OrElse refDoc.UnitsOfMeasure.MassUnits <> targetMassUnits Then
                            Log($" -> Alterando unidades de: {System.IO.Path.GetFileName(docPath)}")
                            refDoc.UnitsOfMeasure.LengthUnits = targetLengthUnits
                            refDoc.UnitsOfMeasure.MassUnits = targetMassUnits
                            refDoc.Update()
                            If refDoc.RequiresUpdate Then
                                ' Tentar salvar apenas se estiver aberto para escrita
                                If Not refDoc.IsReadOnly Then
                                    Try
                                        refDoc.Save()
                                        Log("    -> Documento salvo.")
                                    Catch exSave As Exception
                                        Log($"    ## AVISO: Falha ao salvar '{System.IO.Path.GetFileName(docPath)}': {exSave.Message}")
                                    End Try
                                Else
                                    Log("    -> AVISO: Documento somente leitura, não foi salvo.")
                                End If
                            End If
                            changedRefCount += 1
                            'Else ' Logar isso pode poluir muito o log em montagens grandes
                            '     Log($" -> Unidades já corretas em: {System.IO.Path.GetFileName(docPath)}. Pulando.")
                        End If
                        processedDocs.Add(docPath) ' Marca como processado
                    End If
                Catch exRef As Exception
                    Log($" ## ERRO ao processar '{System.IO.Path.GetFileName(docPath)}': {exRef.Message}")
                Finally
                    ' ReleaseComObject(refDoc) ' Evitar liberar dentro do loop For Each
                End Try
            Next
            Log($"{changedRefCount} documentos referenciados tiveram suas unidades alteradas.")


            ' 6. Finalizar Transação e Operação
            oTransaction.End()
            Log("Transação finalizada.")

            ' Atualiza o documento ativo uma última vez para refletir mudanças nos componentes
            activeDoc.Update()

            Log("-------------------------------------")
            Log($"Operação concluída. {changedRefCount + 1} documentos processados.")

            ' 7. Mensagem de Sucesso (opcional, já está no log)
            'Dim msgSuccess = New MessageBoxWindow($"Unidades alteradas com sucesso...", "Operação Concluída", ModernMessageBoxButtons.OK)
            'msgSuccess.Owner = Me
            'msgSuccess.ShowDialog()


        Catch ex As Exception
            Log("======================================")
            Log("ERRO GERAL INESPERADO: " & ex.Message)
            Log(ex.StackTrace)
            Log("======================================")
            If oTransaction IsNot Nothing Then
                Try
                    oTransaction.Abort()
                    Log("Transação abortada devido a erro.")
                Catch exAbort As Exception
                    Log($"Erro ao abortar transação: {exAbort.Message}")
                End Try
            End If

            ' Mensagem de Erro (usando a caixa de mensagem customizada)
            Dim msgEx = New MessageBoxWindow($"Ocorreu um erro ao alterar as unidades: {vbCrLf}{ex.Message}", "Erro na Operação", ModernMessageBoxButtons.OK)
            msgEx.Owner = Me
            msgEx.ShowDialog()

        Finally
            _invApp.SilentOperation = False ' Garante que a atualização da tela seja reativada
            FinalizeLog() ' Garante que o log final seja exibido
        End Try
    End Sub

    ' --- Funções Auxiliares de Log ---
    Private Sub Log(message As String)
        _logBuilder.AppendLine($"{DateTime.Now:HH:mm:ss} - {message}")
        ' Atualiza o TextBox na thread da UI de forma segura
        Me.Dispatcher.Invoke(Sub()
                                 If LogTextBox IsNot Nothing Then
                                     LogTextBox.Text = _logBuilder.ToString()
                                     LogTextBox.ScrollToEnd() ' Garante que o log role para o final
                                 End If
                             End Sub)
    End Sub

    Private Sub FinalizeLog()
        ' Apenas garante que a última atualização do log seja exibida
        Me.Dispatcher.Invoke(Sub()
                                 If LogTextBox IsNot Nothing Then
                                     LogTextBox.Text = _logBuilder.ToString()
                                     LogTextBox.ScrollToEnd()
                                 End If
                             End Sub)
    End Sub

    ' Opcional: Função auxiliar para liberar objetos COM (se necessário no futuro)
    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                While Marshal.ReleaseComObject(obj) > 0
                End While
            End If
        Catch ex As Exception
            ' Logar erro de liberação se necessário, mas não parar a execução
            Debug.WriteLine($"Erro ao liberar objeto COM: {ex.Message}")
        Finally
            obj = Nothing
        End Try
    End Sub

End Class