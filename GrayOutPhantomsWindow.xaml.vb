' Arquivo: GrayOutPhantomsWindow.xaml.vb
Imports Inventor
Imports System.Windows
Imports System.Text
Imports System.Linq
Imports System.Runtime.InteropServices

Public Class GrayOutPhantomsWindow
    Private _logBuilder As New StringBuilder()
    Private _invApp As Inventor.Application
    Private ReadOnly _targetColor As Inventor.Color ' Cor cinza (192, 192, 192)

    Public Sub New()
        InitializeComponent()
        _invApp = Globals.g_inventorApplication
        ' Define a cor cinza uma vez no construtor
        _targetColor = _invApp.TransientObjects.CreateColor(192, 192, 192)
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
        Log("Iniciando realce de Phantoms/Referência...")

        Dim oDoc As DrawingDocument = Nothing
        Dim oTransaction As Transaction = Nothing
        Dim pecasProcessadas As Integer = 0
        Dim pecasAlteradas As Integer = 0

        Try
            ' 1. Validação do Documento
            oDoc = TryCast(_invApp.ActiveDocument, DrawingDocument)
            If oDoc Is Nothing OrElse oDoc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                Log("ERRO: O documento ativo não é um desenho (.idw).")
                FinalizeLog()
                Return
            End If
            Log($"Documento de desenho ativo: {oDoc.DisplayName}")

            ' 2. Iniciar Transação e Otimização
            oTransaction = _invApp.TransactionManager.StartTransaction(oDoc, "Realçar Phantoms/Referência")
            _invApp.SilentOperation = True
            Log("Transação iniciada. Ocultando atualizações de tela.")

            ' 3. Iterar sobre Folhas e Vistas
            Dim oSheet As Inventor.Sheet = Nothing
            Dim oView As DrawingView = Nothing

            For Each oSheet In oDoc.Sheets
                Log($"Processando Folha: '{oSheet.Name}'")
                For Each oView In oSheet.DrawingViews
                    Log($" -> Verificando Vista: '{oView.Name}'")
                    Dim refDoc As Document = Nothing
                    Dim oAssyDef As AssemblyComponentDefinition = Nothing

                    Try
                        ' Pula vistas sem referência ou que não referenciam montagens
                        If oView.ReferencedDocumentDescriptor Is Nothing Then
                            Log("    -> Vista sem documento referenciado. Pulando.")
                            Continue For
                        End If
                        refDoc = oView.ReferencedDocumentDescriptor.ReferencedDocument
                        If refDoc Is Nothing Then
                            Log("    -> Não foi possível obter o documento referenciado. Pulando.")
                            Continue For
                        End If
                        If refDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                            Log("    -> Vista não referencia uma montagem. Pulando.")
                            Continue For
                        End If

                        oAssyDef = TryCast(refDoc.ComponentDefinition, AssemblyComponentDefinition)
                        If oAssyDef Is Nothing Then
                            Log("    -> Não foi possível obter a definição da montagem. Pulando.")
                            Continue For
                        End If

                        ' Chama a função recursiva para processar ocorrências
                        Log($"    -> Processando ocorrências da montagem '{System.IO.Path.GetFileName(refDoc.FullFileName)}'...")
                        ProcessarOcorrencias(oView, oAssyDef.Occurrences, _targetColor, pecasProcessadas, pecasAlteradas)

                    Catch refEx As Exception
                        Log($"    ## ERRO ao acessar referência da vista '{oView.Name}': {refEx.Message}")
                    Finally
                        ' Liberar objetos COM obtidos dentro do loop da vista, se houver
                        ReleaseComObject(oAssyDef)
                        ReleaseComObject(refDoc)
                    End Try
                Next ' Fim do loop de Vistas
            Next ' Fim do loop de Folhas

            ' 4. Finalizar
            Log("Finalizando processamento...")
            oDoc.Update() ' Aplica as mudanças de cor visualmente
            oTransaction.End()
            _invApp.SilentOperation = False
            Log("Transação concluída. Atualizações de tela reativadas.")
            Log("-------------------------------------")

            ' 5. Exibir Resultado
            If pecasProcessadas = 0 Then
                Log("Nenhuma peça com estrutura BOM 'Phantom' ou 'Reference' foi encontrada nas vistas processadas.")
                Dim msgInfo = New MessageBoxWindow("Nenhuma peça Phantom/Reference encontrada nas vistas que referenciam montagens.", "Aviso", ModernMessageBoxButtons.OK)
                msgInfo.Owner = Me
                msgInfo.ShowDialog()
            Else
                Log($"{pecasAlteradas} de {pecasProcessadas} peças encontradas tiveram suas curvas alteradas para cinza.")
                Dim msgSuccess = New MessageBoxWindow($"{pecasAlteradas} de {pecasProcessadas} peças Phantom/Reference encontradas foram realçadas com sucesso.", "Operação Concluída", ModernMessageBoxButtons.OK)
                msgSuccess.Owner = Me
                msgSuccess.ShowDialog()
            End If


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
            _invApp.SilentOperation = False

            Dim msgEx = New MessageBoxWindow($"Ocorreu um erro durante o processo: {vbCrLf}{ex.Message}", "Erro na Operação", ModernMessageBoxButtons.OK)
            msgEx.Owner = Me
            msgEx.ShowDialog()

        Finally
            _invApp.SilentOperation = False ' Garante reativação
            FinalizeLog()
            ' Liberar objetos principais se necessário (geralmente não precisa para oDoc, sheets, views)
            ReleaseComObject(oTransaction) ' Liberar objeto Transaction
        End Try
    End Sub

    ' Função Recursiva para processar ocorrências
    Private Sub ProcessarOcorrencias(oView As DrawingView, oOccurrences As ComponentOccurrences, oColor As Inventor.Color, ByRef total As Integer, ByRef alteradas As Integer)
        If oOccurrences Is Nothing OrElse oOccurrences.Count = 0 Then Return

        Dim oOcc As ComponentOccurrence = Nothing
        Dim oPartDef As PartComponentDefinition = Nothing
        Dim oCurvas As DrawingCurvesEnumerator = Nothing
        Dim curva As DrawingCurve = Nothing

        For Each oOcc In oOccurrences
            ' Pula ocorrências suprimidas na montagem
            If oOcc.Suppressed Then Continue For

            Try
                ' Tenta obter a definição como Peça
                oPartDef = TryCast(oOcc.Definition, PartComponentDefinition)

                ' Se for uma peça, verifica a estrutura da BOM
                If oPartDef IsNot Nothing Then
                    If oPartDef.BOMStructure = BOMStructureEnum.kPhantomBOMStructure OrElse
                       oPartDef.BOMStructure = BOMStructureEnum.kReferenceBOMStructure Then

                        total += 1 ' Incrementa contador de peças encontradas
                        Dim curvasAlteradasNestaPeca As Boolean = False
                        Log($"    -> Encontrada: '{oOcc.Name}' (Phantom/Reference)")

                        Try
                            ' Obtém as curvas específicas desta ocorrência DENTRO da vista atual
                            oCurvas = oView.DrawingCurves(oOcc)

                            ' Itera sobre as curvas e aplica a cor
                            If oCurvas IsNot Nothing AndAlso oCurvas.Count > 0 Then
                                For Each curva In oCurvas
                                    Try
                                        ' Aplica a cor cinza definida
                                        curva.Color = oColor
                                        curvasAlteradasNestaPeca = True ' Marca que pelo menos uma curva foi alterada
                                    Catch exCurva As Exception
                                        Log($"       ## ERRO ao colorir curva para '{oOcc.Name}': {exCurva.Message}")
                                    Finally
                                        ReleaseComObject(curva) ' Libera objeto curva dentro do loop interno
                                    End Try
                                Next
                            Else
                                Log($"       -> Nenhuma curva encontrada para '{oOcc.Name}' nesta vista.")
                            End If

                            ' Incrementa o contador de peças alteradas SOMENTE se alguma curva foi modificada
                            If curvasAlteradasNestaPeca Then
                                alteradas += 1
                                Log($"       -> Curvas de '{oOcc.Name}' alteradas para cinza.")
                            End If

                        Catch exCurvas As Exception
                            Log($"    ## ERRO ao obter/processar curvas para '{oOcc.Name}': {exCurvas.Message}")
                        Finally
                            ReleaseComObject(oCurvas) ' Libera o enumerador de curvas
                        End Try
                    End If ' Fim do If BOMStructure
                End If ' Fim do If oPartDef IsNot Nothing

                ' Processa subocorrências recursivamente, SE a ocorrência atual NÃO for suprimida na vista
                If Not oOcc.IsSuppressed(oView) AndAlso oOcc.SubOccurrences IsNot Nothing AndAlso oOcc.SubOccurrences.Count > 0 Then
                    ProcessarOcorrencias(oView, oOcc.SubOccurrences, oColor, total, alteradas)
                End If

            Catch exOcc As Exception
                Log($" ## ERRO ao processar ocorrência '{oOcc?.Name}': {exOcc.Message}")
            Finally
                ' Libera objetos obtidos dentro deste loop
                ReleaseComObject(oPartDef)
                ReleaseComObject(oOcc) ' Libera a ocorrência atual
            End Try
        Next ' Fim do For Each oOcc
    End Sub


    ' --- Funções Auxiliares de Log ---
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