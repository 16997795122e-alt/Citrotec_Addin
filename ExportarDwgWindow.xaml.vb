' Arquivo: ExportarDwgWindow.xaml.vb
Imports Inventor
Imports System.Windows
Imports System.Text
Imports System.IO
Imports System.Diagnostics
Imports System.Collections.Generic

Public Class ExportarDwgWindow
    Private _logBuilder As New StringBuilder()
    Private _invApp As Inventor.Application

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub btnExecute_Click(sender As Object, e As RoutedEventArgs)
        _logBuilder.Clear()
        LogTextBox.Clear()
        Log("Iniciando exportação DWG...")

        _invApp = g_inventorApplication

        Dim oDrawDoc As DrawingDocument = Nothing
        Dim oTagProjetista As String = ""
        Dim tagFinal As String = ""
        Dim savePath As String = ""

        Try
            ' 1. Validação Inicial
            oDrawDoc = TryCast(_invApp.ActiveDocument, DrawingDocument)
            If oDrawDoc Is Nothing OrElse oDrawDoc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                Log("ERRO: O documento ativo não é um desenho (.idw).")
                FinalizeLog()
                Return
            End If
            Log("Documento de desenho ativo: " & System.IO.Path.GetFileName(oDrawDoc.FullFileName))

            ' 2. Obter e Preparar a TAG do Usuário do Add-in
            If SessionManager.IsUserLoggedIn AndAlso Not String.IsNullOrWhiteSpace(SessionManager.CurrentUser.UserTag) Then
                oTagProjetista = SessionManager.CurrentUser.UserTag
                Log("TAG do usuário logado: " & oTagProjetista)
                tagFinal = oTagProjetista.Replace(".", "") ' Remove pontos
                Log("TAG formatada para nome de arquivo: " & tagFinal)
            Else
                Log("AVISO: Usuário não logado ou sem TAG definida. Usando TAG padrão 'EAS'.")
                oTagProjetista = "E.A.S."
                tagFinal = "EAS"
            End If

            ' 3. Confirmação com o Usuário
            Dim confirmMsg As String = $"O arquivo será salvo com a TAG DESENHISTA: {tagFinal}{vbCrLf}Deseja continuar?"
            Dim msgBoxConfirm = New MessageBoxWindow(confirmMsg, "CITROTEC - CONFIRMAÇÃO", ModernMessageBoxButtons.YesNo)
            msgBoxConfirm.Owner = Me
            msgBoxConfirm.ShowDialog()

            If msgBoxConfirm.Result = MessageBoxResult.No Then
                Dim inputBox = New InputBoxWindow("Digite a nova sigla do Projetista sem pontos (Exemplo: EAS)", "NOVA TAG - CITROTEC")
                inputBox.InputText = tagFinal
                inputBox.Owner = Me
                If inputBox.ShowDialog().GetValueOrDefault(False) Then
                    tagFinal = inputBox.InputText.Trim().Replace(".", "")
                    If String.IsNullOrEmpty(tagFinal) Then
                        Log("ERRO: Nenhuma TAG fornecida. O processo foi cancelado.")
                        FinalizeLog()
                        Return
                    End If
                    Log("TAG alterada pelo usuário para: " & tagFinal)
                Else
                    Log("Processo cancelado pelo usuário na solicitação de nova TAG.")
                    FinalizeLog()
                    Return
                End If
            ElseIf msgBoxConfirm.Result <> MessageBoxResult.Yes Then
                Log("Processo cancelado pelo usuário na confirmação.")
                FinalizeLog()
                Return
            End If

            ' 4. Ler iProperties e Construir Caminho de Salvamento
            Dim N_Desenho As String = GetPropertyValueFromDoc(oDrawDoc, "NRODESENHO")
            Dim N_Ot As String = GetPropertyValueFromDoc(oDrawDoc, "NROOT")
            Dim N_RevisionStr As String = GetPropertyValueFromDoc(oDrawDoc, "Revision Number", "Inventor Summary Information")
            Dim N_Revision As Integer = 0

            If String.IsNullOrWhiteSpace(N_Desenho) OrElse String.IsNullOrWhiteSpace(N_Ot) Then
                Log("ERRO: As iProperties 'NRODESENHO' ou 'NROOT' não foram encontradas ou estão vazias.")
                FinalizeLog()
                Return
            End If
            If String.IsNullOrWhiteSpace(N_RevisionStr) Then
                N_RevisionStr = "0"
                Log("AVISO: 'Revision Number' não definido. Usando valor padrão '0'.")
            End If
            Integer.TryParse(N_RevisionStr, N_Revision)

            Log($"Propriedades lidas: NRODESENHO='{N_Desenho}', NROOT='{N_Ot}', Revision='{N_Revision}'")

            Dim revisionFormatted As String = If(N_Revision < 10, $"-0{N_Revision}", $"-{N_Revision}")
            Dim fileName As String = $"{N_Ot} - {tagFinal} - {N_Desenho}{revisionFormatted}.dwg"
            Dim saveFolder As String = "\\SERVIDOR20\eng\Transferência Desenhos\3_Desenhos para serem aprovados"
            savePath = System.IO.Path.Combine(saveFolder, fileName)
            Log("Caminho de salvamento definido: " & savePath)

            ' Garante que a pasta de destino exista
            If Not Directory.Exists(saveFolder) Then
                Log($"ERRO: A pasta de destino '{saveFolder}' não foi encontrada ou não está acessível.")
                FinalizeLog()
                Return
            End If

            ' 5. Configurar Opções de Exportação DWG
            Dim DXFAddIn As TranslatorAddIn = Nothing
            Try
                DXFAddIn = CType(_invApp.ApplicationAddIns.ItemById("{C24E3AC4-122E-11D5-8E91-0010B541CD80}"), TranslatorAddIn)
            Catch ex As Exception
                Log("ERRO: Não foi possível encontrar o AddIn de exportação DWG/DXF. Verifique se ele está carregado no Inventor.")
                Log(ex.Message)
                FinalizeLog()
                Return
            End Try

            Dim oContext As TranslationContext = _invApp.TransientObjects.CreateTranslationContext
            oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
            Dim oOptions As NameValueMap = _invApp.TransientObjects.CreateNameValueMap
            Dim oDataMedium As DataMedium = _invApp.TransientObjects.CreateDataMedium
            oDataMedium.FileName = savePath

            If DXFAddIn.HasSaveCopyAsOptions(oDrawDoc, oContext, oOptions) Then
                Dim strIniFile As String = "\\SERVIDOR20\eng\Engenharia\iLogic\Config_Export_DWG.ini"
                If System.IO.File.Exists(strIniFile) Then
                    oOptions.Value("Export_Acad_IniFile") = strIniFile
                    Log("Arquivo de configuração .ini encontrado e aplicado: " & strIniFile)
                Else
                    Log("AVISO: Arquivo .ini de configuração não encontrado em " & strIniFile & ". Usando opções padrão.")
                End If
            End If

            ' 6. Executar Exportação
            Log("Iniciando SaveCopyAs para DWG...")
            DXFAddIn.SaveCopyAs(oDrawDoc, oContext, oOptions, oDataMedium)
            Log("Arquivo DWG salvo com sucesso!")

            ' 7. Executar Lógica de Atualização da Montagem
            Log("Iniciando atualização das propriedades da montagem...")
            UpdateAssemblyProperties(oDrawDoc)
            Log("Atualização das propriedades da montagem concluída.")

            ' 8. Perguntar se deseja abrir a pasta
            Dim msgBoxOpen = New MessageBoxWindow("Gostaria de abrir o local do arquivo salvo?", "CITROTEC - EXPORT DWG", ModernMessageBoxButtons.YesNo)
            msgBoxOpen.Owner = Me
            msgBoxOpen.ShowDialog()
            If msgBoxOpen.Result = MessageBoxResult.Yes Then
                Try
                    Process.Start("explorer.exe", saveFolder)
                Catch ex As Exception
                    Log("Erro ao tentar abrir a pasta: " & ex.Message)
                End Try
            End If

            Log("-------------------------------------")
            Log("Exportação DWG finalizada com sucesso.")

        Catch ex As Exception
            Log("======================================")
            Log("ERRO GERAL INESPERADO DURANTE A EXPORTAÇÃO: " & ex.Message)
            Log(ex.StackTrace)
            Log("======================================")
        Finally
            FinalizeLog()
        End Try
    End Sub

    Private Sub UpdateAssemblyProperties(oDrawDoc As DrawingDocument)
        If oDrawDoc Is Nothing Then Return

        Dim oSheet As Sheet = oDrawDoc.ActiveSheet
        If oSheet.DrawingViews.Count = 0 Then
            Log("   -> Nenhuma vista encontrada no desenho para referenciar montagem. Pulando atualização.")
            Return
        End If

        Dim oView As DrawingView = oSheet.DrawingViews.Item(1)
        Dim refDocDesc As DocumentDescriptor = oView.ReferencedDocumentDescriptor

        If refDocDesc Is Nothing OrElse refDocDesc.ReferencedDocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            Log("   -> A primeira vista não referencia uma montagem (.iam). Pulando atualização.")
            Return
        End If

        Dim oAsmDoc As AssemblyDocument = Nothing
        Dim mustCloseAssembly As Boolean = False
        Dim assemblyFullFileName As String = refDocDesc.FullDocumentName

        Try
            Try
                oAsmDoc = TryCast(_invApp.Documents.Item(assemblyFullFileName), AssemblyDocument)
            Catch exNotFound As Exception
                oAsmDoc = Nothing
            End Try

            If oAsmDoc Is Nothing Then
                Log("   -> Abrindo montagem referenciada: " & System.IO.Path.GetFileName(assemblyFullFileName))
                oAsmDoc = TryCast(_invApp.Documents.Open(assemblyFullFileName, False), AssemblyDocument)
                mustCloseAssembly = True
            Else
                Log("   -> Montagem referenciada já está aberta.")
            End If

            If oAsmDoc Is Nothing Then Throw New Exception("Falha ao obter o documento da montagem.")

            Dim oAsmProps As PropertySet = oAsmDoc.PropertySets.Item("Inventor User Defined Properties")
            Dim oIdwProps As PropertySet = oDrawDoc.PropertySets.Item("Inventor User Defined Properties")

            Dim propList As New List(Of String) From {"NRODESENHO", "NROOT", "TITULO2"}
            Dim propertiesUpdated As Boolean = False

            For Each propName As String In propList
                Try
                    Dim idwPropValue As String = oIdwProps(propName).Value.ToString()
                    If UpdateOrAddPropertyIfChanged(oAsmProps, propName, idwPropValue) Then
                        Log($"      -> Propriedade '{propName}' atualizada na montagem.")
                        propertiesUpdated = True
                    End If
                Catch exProp As Exception
                    Log($"      -> AVISO: Propriedade '{propName}' não encontrada no desenho ou erro ao copiar: {exProp.Message}")
                End Try
            Next

            If propertiesUpdated Then
                Log("   -> Salvando alterações na montagem...")
                oAsmDoc.Update()
                oAsmDoc.Save()
                Log("   -> Montagem salva.")
            Else
                Log("   -> Nenhuma propriedade precisou ser atualizada na montagem.")
            End If

        Catch ex As Exception
            Log("   -> ERRO durante a atualização da montagem: " & ex.Message)
        Finally
            If mustCloseAssembly AndAlso oAsmDoc IsNot Nothing Then
                Log("   -> Fechando montagem.")
                oAsmDoc.Close(True)
            End If
        End Try
    End Sub

    Private Function UpdateOrAddPropertyIfChanged(props As PropertySet, propName As String, propValue As String) As Boolean
        Try
            Dim existingProp As Inventor.Property = props.Item(propName)
            If Not existingProp.Value.ToString().Equals(propValue) Then
                existingProp.Value = propValue
                Return True
            Else
                Return False
            End If
        Catch exProp As Exception
            props.Add(propValue, propName)
            Return True
        End Try
    End Function

    Private Function GetPropertyValueFromDoc(doc As Document, propName As String, Optional propSetName As String = "Inventor User Defined Properties") As String
        Dim standardSets As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"Revision Number", "Inventor Summary Information"},
            {"Project", "Design Tracking Properties"}
        }

        Dim targetSetName As String = If(standardSets.ContainsKey(propName), standardSets(propName), propSetName)

        Try
            Dim propSet As PropertySet = doc.PropertySets.Item(targetSetName)
            Dim prop As Inventor.Property = propSet.Item(propName)
            Dim propValue As String = If(prop.Value IsNot Nothing, prop.Value.ToString(), "")
            If propName = "Revision Number" AndAlso String.IsNullOrWhiteSpace(propValue) Then
                Log("   -> AVISO: Propriedade 'Revision Number' está vazia. Usando valor padrão '0'.")
                Return "0"
            End If
            Return propValue
        Catch ex As Exception
            If propName = "Revision Number" Then
                Log("   -> AVISO: Não foi possível ler a propriedade 'Revision Number' do conjunto 'Inventor Summary Information'. Usando valor padrão '0'.")
                Return "0"
            End If
            Log($"   -> AVISO: Não foi possível ler a propriedade '{propName}' do conjunto '{targetSetName}'. {ex.Message}")
            Return ""
        End Try
    End Function

    Private Sub Log(message As String)
        _logBuilder.AppendLine($"{DateTime.Now:HH:mm:ss} - {message}")
        Me.Dispatcher.Invoke(Sub() LogTextBox.Text = _logBuilder.ToString())
    End Sub

    Private Sub FinalizeLog()
        Me.Dispatcher.Invoke(Sub()
                                 LogTextBox.Text = _logBuilder.ToString()
                                 LogTextBox.ScrollToEnd()
                             End Sub)
    End Sub
End Class