Imports Inventor
Imports System.Windows
Imports System.Text
Imports System.Collections.Generic
Imports System.Linq

Public Class CriarPlanificacoesWindow
    Private _logBuilder As New StringBuilder()
    Private _invApp As Inventor.Application
    Private _drawDoc As DrawingDocument

    '===========================================================
    '#Region "Lógica de Gerenciamento do Parâmetro"
    '===========================================================

    Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
        _invApp = g_inventorApplication
        _drawDoc = TryCast(_invApp.ActiveDocument, DrawingDocument)
        LoadProcessedItemsToListBox()
    End Sub

    Private Sub LoadProcessedItemsToListBox()
        ProcessedItemsListBox.ItemsSource = Nothing
        If _drawDoc Is Nothing Then Return

        Dim param As UserParameter = GetOrCreateFlatParameter()
        If param IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(param.Value) Then
            Dim itemList As New List(Of String)(param.Value.ToString().Split(vbLf))
            itemList.RemoveAll(Function(s) String.IsNullOrWhiteSpace(s))
            ' Tenta converter para inteiro para ordenar numericamente, caso contrário ordena como texto
            ProcessedItemsListBox.ItemsSource = itemList.OrderBy(Function(s)
                                                                     Dim num As Integer
                                                                     If Integer.TryParse(s, num) Then Return num
                                                                     Return -1
                                                                 End Function).ToList()
        End If
    End Sub

    Private Sub btnDeleteSelected_Click(sender As Object, e As RoutedEventArgs)
        If ProcessedItemsListBox.SelectedItem Is Nothing Then
            MessageBox.Show("Por favor, selecione um item da lista para remover.", "Nenhum Item Selecionado", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        Dim itemToRemove As String = ProcessedItemsListBox.SelectedItem.ToString()

        Log($"Procurando vista para o item '{itemToRemove}' para exclusão...")
        Dim sheet As Sheet = _drawDoc.ActiveSheet
        Dim viewDeleted As Boolean = False
        For i As Integer = sheet.DrawingViews.Count To 1 Step -1
            Dim view As DrawingView = sheet.DrawingViews.Item(i)
            Try
                Dim attrSet As AttributeSet = view.AttributeSets.Item("CitrotecAddinData")
                If attrSet.Item("ListaItem").Value.ToString() = itemToRemove Then
                    If view.Position.X > sheet.Width Then
                        Log($"   -> Vista '{view.Name}' encontrada fora da folha. Excluindo...")
                        view.Delete()
                        viewDeleted = True
                        Exit For
                    Else
                        Log($"   -> Vista '{view.Name}' encontrada, mas está dentro da folha. A exclusão foi ignorada.")
                    End If
                End If
            Catch ex As Exception
            End Try
        Next
        If Not viewDeleted Then Log("   -> Nenhuma vista correspondente encontrada fora da folha para exclusão.")

        Dim param As UserParameter = GetOrCreateFlatParameter()
        Dim itemList As New List(Of String)(param.Value.ToString().Split(vbLf))
        itemList.RemoveAll(Function(s) String.IsNullOrWhiteSpace(s) Or s.Equals(itemToRemove, StringComparison.OrdinalIgnoreCase))

        param.Value = String.Join(vbLf, itemList)
        Log($"Item '{itemToRemove}' removido do parâmetro 'oComponentsFlat'.")
        FinalizeLog()

        LoadProcessedItemsToListBox()
    End Sub

    Private Sub btnClearAll_Click(sender As Object, e As RoutedEventArgs)
        Dim result = MessageBox.Show("Tem certeza que deseja limpar a lista E EXCLUIR TODAS as vistas planificadas criadas por este comando fora da folha?",
                                     "Confirmar Limpeza Total", MessageBoxButton.YesNo, MessageBoxImage.Warning)
        If result = MessageBoxResult.No Then Return

        Log("Iniciando limpeza total...")
        Dim sheet As Sheet = _drawDoc.ActiveSheet
        Dim deletedCount As Integer = 0
        For i As Integer = sheet.DrawingViews.Count To 1 Step -1
            Dim view As DrawingView = sheet.DrawingViews.Item(i)
            Try
                Dim attrSet As AttributeSet = view.AttributeSets.Item("CitrotecAddinData")
                If attrSet IsNot Nothing Then
                    If view.Position.X > sheet.Width Then
                        Log($"   -> Excluindo vista '{view.Name}'...")
                        view.Delete()
                        deletedCount += 1
                    End If
                End If
            Catch ex As Exception
                ' Ignora vistas que não têm a nossa etiqueta
            End Try
        Next
        Log($"{deletedCount} vistas foram excluídas com sucesso.")

        Dim param As UserParameter = GetOrCreateFlatParameter()
        param.Value = ""
        Log("Parâmetro 'oComponentsFlat' foi limpo.")
        FinalizeLog()

        LoadProcessedItemsToListBox()
    End Sub

    Private Function GetOrCreateFlatParameter() As UserParameter
        If _drawDoc Is Nothing Then Return Nothing
        Try
            Return _drawDoc.Parameters.UserParameters.Item("oComponentsFlat")
        Catch
            Return _drawDoc.Parameters.UserParameters.AddByValue("oComponentsFlat", "", UnitsTypeEnum.kTextUnits)
        End Try
    End Function

    '===========================================================
    '#End Region
    '===========================================================

    '===========================================================
    '#Region "Lógica Principal de Execução"
    '===========================================================

    Private Sub btnExecute_Click(sender As Object, e As RoutedEventArgs)
        _logBuilder.Clear()
        LogTextBox.Clear()
        Log("Iniciando rotina para criar planificações...")

        Dim refDoc As Document = Nothing

        Try
            If _drawDoc Is Nothing OrElse _drawDoc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                Log("ERRO: O documento ativo não é um desenho (.idw).")
                FinalizeLog()
                Return
            End If

            Dim oSheet As Sheet = _drawDoc.ActiveSheet
            Log("Folha ativa: " & oSheet.Name)

            ' Validação da folha e vista base
            Dim sheetSizeMap As New Dictionary(Of DrawingSheetSizeEnum, Tuple(Of Double, Double)) From {
                {DrawingSheetSizeEnum.kA1DrawingSheetSize, Tuple.Create(84.1, 59.4)},
                {DrawingSheetSizeEnum.kA2DrawingSheetSize, Tuple.Create(59.4, 42.0)},
                {DrawingSheetSizeEnum.kA3DrawingSheetSize, Tuple.Create(42.0, 29.7)},
                {DrawingSheetSizeEnum.kA4DrawingSheetSize, Tuple.Create(29.7, 21.0)}
            }

            If Not sheetSizeMap.ContainsKey(oSheet.Size) Then
                Log("ERRO: Formato da folha não suportado por esta rotina.")
                FinalizeLog()
                Return
            End If

            Dim sheetDims As Tuple(Of Double, Double) = sheetSizeMap(oSheet.Size)
            Dim sheetWidth As Double = sheetDims.Item1
            Log($"Dimensões da folha ({oSheet.Size}): {sheetWidth}cm x ...")

            If oSheet.PartsLists.Count = 0 Then
                Log("ERRO: Nenhuma Parts List encontrada na folha.")
                FinalizeLog()
                Return
            End If

            Dim partsList As PartsList = oSheet.PartsLists.Item(1)
            refDoc = partsList.ReferencedDocumentDescriptor.ReferencedDocument

            If refDoc.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
                Log("ERRO: A Parts List não referencia uma montagem (.iam).")
                FinalizeLog()
                Return
            End If
            Log("Montagem de referência via Parts List: " & System.IO.Path.GetFileName(refDoc.FullFileName))

            ' Obter escala base de uma vista existente na folha
            Dim baseScale As Double = 1.0
            If oSheet.DrawingViews.Count > 0 Then
                baseScale = oSheet.DrawingViews.Item(1).Scale
                Log($"Escala base obtida da primeira vista: {baseScale}")
            Else
                Log("Nenhuma vista encontrada na folha. Usando escala padrão 1:1.")
            End If

            Dim oComponentsFlatParam As UserParameter = GetOrCreateFlatParameter()
            Dim itemList As New List(Of String)(oComponentsFlatParam.Value.ToString().Split(vbLf))
            itemList.RemoveAll(Function(s) String.IsNullOrWhiteSpace(s))
            Log($"Encontrados {itemList.Count} itens já processados no parâmetro 'oComponentsFlat'.")

            ' Lógica principal: leitura da Parts List e criação de vistas
            Dim excludeKeywords As String() = {"ARR_PRESS", "ARR_LISA", "Controle"}
            Dim validParts As New List(Of Tuple(Of Integer, PartDocument, String, String))

            CollectValidPartsFromPartsList(partsList, validParts, excludeKeywords, itemList)

            Log($"Parts List analisada. Encontradas {validParts.Count} novas peças de chapa metálica para detalhar.")

            If validParts.Count = 0 Then
                Log("Nenhuma peça nova para processar. Rotina finalizada.")
                FinalizeLog()
                Return
            End If

            validParts.Sort(Function(a, b) a.Item1.CompareTo(b.Item1))

            Dim createdViews As New List(Of DrawingView)
            CreateOrderedViews(validParts, oSheet, baseScale, createdViews, itemList)
            Log($"Processo de criação de vistas concluído. {createdViews.Count} novas vistas foram criadas.")

            If createdViews.Count > 0 Then
                ApplyLabelsToViews(createdViews)
                Log("Rótulos aplicados às novas vistas.")
            End If

            oComponentsFlatParam.Value = String.Join(vbLf, itemList)
            Log($"Parâmetro 'oComponentsFlat' atualizado com {itemList.Count} itens.")
            LoadProcessedItemsToListBox()

            _drawDoc.Update()
            Log("Documento atualizado.")
            Log("-------------------------------------")
            Log("Execução finalizada com sucesso.")
        Catch ex As Exception
            Log("======================================")
            Log("ERRO GERAL INESPERADO: " & ex.Message)
            Log(ex.StackTrace)
            Log("======================================")
        Finally
            FinalizeLog()
        End Try
    End Sub

    Private Sub CollectValidPartsFromPartsList(partsList As PartsList, ByRef validParts As List(Of Tuple(Of Integer, PartDocument, String, String)), excludeKeywords As String(), itemList As List(Of String))
        Try
            Dim rowCount As Integer = partsList.PartsListRows.Count
            Log($"Processando {rowCount} linhas na Parts List.")

            For i As Integer = 1 To rowCount
                Dim row As PartsListRow = partsList.PartsListRows.Item(i)
                If row.ReferencedFiles.Count = 0 Then
                    Log($"Aviso: Linha {i} sem arquivos referenciados. Pulando.")
                    Continue For
                End If

                Dim doc As Document = row.ReferencedFiles.Item(1).ReferencedDocument
                If doc.DocumentType <> DocumentTypeEnum.kPartDocumentObject Then
                    Log($"Aviso: Linha {i} não referencia uma peça. Pulando.")
                    Continue For
                End If

                If doc.SubType <> "{9C464203-9BAE-11D3-8BAD-0060B0CE6BB4}" Then
                    Log($"Aviso: Linha {i} não é uma peça de chapa metálica. Pulando.")
                    Continue For
                End If

                Dim partDoc As PartDocument = CType(doc, PartDocument)
                Dim smDef As SheetMetalComponentDefinition = TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
                If smDef Is Nothing OrElse Not smDef.HasFlatPattern Then
                    Log($"Aviso: Linha {i} ({partDoc.DisplayName}) não possui planificação válida. Pulando.")
                    Continue For
                End If

                If excludeKeywords.Any(Function(k) partDoc.FullFileName.ToUpper().Contains(k)) Then
                    Log($"Aviso: Linha {i} ({partDoc.DisplayName}) contém palavra-chave de exclusão. Pulando.")
                    Continue For
                End If

                Dim itemNumberStr As String = ""
                Dim itemQtyStr As String = ""

                Try
                    itemNumberStr = row.Item("ITEM").Value.ToString()
                    itemQtyStr = row.Item("ITEM_QTY").Value.ToString()
                Catch ex As Exception
                    Log($"Aviso: Coluna 'ITEM' ou 'ITEM_QTY' não encontrada na linha {i}. Erro: {ex.Message}. Pulando.")
                    Continue For
                End Try

                If String.IsNullOrEmpty(itemNumberStr) Then
                    Log($"Aviso: Item number vazio na linha {i}. Pulando.")
                    Continue For
                End If

                If itemList.Contains(itemNumberStr) Then
                    Log($"Aviso: Item {itemNumberStr} na linha {i} já processado. Pulando.")
                    Continue For
                End If

                Dim itemNumber As Integer
                If Integer.TryParse(itemNumberStr, itemNumber) Then
                    validParts.Add(Tuple.Create(itemNumber, partDoc, itemNumberStr, itemQtyStr))
                    Log($"Peça válida encontrada na linha {i}: {partDoc.DisplayName} (Item {itemNumberStr}, Qtd: {itemQtyStr})")
                Else
                    Log($"Aviso: Item number '{itemNumberStr}' na linha {i} não é um inteiro válido. Pulando.")
                End If
            Next
        Catch ex As Exception
            Log($"ERRO ao processar Parts List: {ex.Message}")
        End Try
    End Sub

    Private Sub CreateOrderedViews(validParts As List(Of Tuple(Of Integer, PartDocument, String, String)), oSheet As Sheet, baseScale As Double, ByRef createdViews As List(Of DrawingView), ByRef itemList As List(Of String))
        Dim options As NameValueMap = _invApp.TransientObjects.CreateNameValueMap()
        options.Add("SheetMetalFoldedModel", False)

        Dim sheetWidth As Double = oSheet.Width
        Dim startX As Double = sheetWidth + 5
        Dim currentX As Double = startX
        Dim currentY As Double = 5
        Dim maxViewHeight As Double = 10
        Dim offsetX As Double = 5
        Dim offsetY As Double = 5

        For Each partData In validParts
            If currentY + maxViewHeight > oSheet.Height Then
                currentX += offsetX
                currentY = 5
            End If

            Dim partDoc As PartDocument = partData.Item2
            Dim itemNumberStr As String = partData.Item3
            Dim itemQtyStr As String = partData.Item4
            Dim oPoint As Point2d = _invApp.TransientGeometry.CreatePoint2d(currentX, currentY)

            Try
                Log($"Criando vista para: {partDoc.DisplayName} (Item {itemNumberStr})...")
                Dim flatView As DrawingView = oSheet.DrawingViews.AddBaseView(partDoc, oPoint, baseScale, ViewOrientationTypeEnum.kFlatBacksideViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle, , , options)

                If flatView IsNot Nothing Then
                    createdViews.Add(flatView)

                    Dim partModel As Document = flatView.ReferencedDocumentDescriptor.ReferencedDocument
                    Dim userProps As PropertySet = partModel.PropertySets.Item("User Defined Properties")
                    SetOrAddiProperty(userProps, "ListaItem", itemNumberStr)
                    SetOrAddiProperty(userProps, "ItemQty", itemQtyStr)

                    Dim attrSet As AttributeSet
                    Try
                        attrSet = flatView.AttributeSets.Item("CitrotecAddinData")
                    Catch
                        attrSet = flatView.AttributeSets.Add("CitrotecAddinData")
                    End Try
                    SetOrAddAttribute(attrSet, "ListaItem", itemNumberStr)

                    itemList.Add(itemNumberStr)

                    If flatView.Height > maxViewHeight Then maxViewHeight = flatView.Height + 5
                    currentY += flatView.Height + offsetY

                    If flatView.IsFlatPatternView Then
                        AddBendNotes(flatView, oSheet)
                    End If
                End If
            Catch ex As Exception
                Log($"ERRO ao criar vista para {partDoc.DisplayName}: {ex.Message}")
            End Try
        Next
    End Sub

    Private Sub ApplyLabelsToViews(createdViews As List(Of DrawingView))
        Dim oPartNumber As String = "<StyleOverride FontSize='0,4' Underline='True'>ITEM </StyleOverride>"
        Dim oListItem As String = "<StyleOverride FontSize='0,4' Underline='True'><Property Document='model' PropertySet='User Defined Properties' Property='ListaItem' FormatID='{D5CDD505-2E9C-101B-9397-08002B2CF9AE}' PropertyID='26'>ListaItem</Property></StyleOverride>"
        Dim oStringScale As String = "<Br/><StyleOverride FontSize='0,25'>( <DrawingViewScale/> )</StyleOverride>"
        Dim oQty As String = "<Br/><StyleOverride FontSize='0,25'>Fabricar: <Property Document='model' PropertySet='User Defined Properties' Property='ItemQty' FormatID='{D5CDD505-2E9C-101B-9397-08002B2CF9AE}' PropertyID='27'>ItemQty</Property>x</StyleOverride>"
        Dim Planificacao As String = "<Br/><StyleOverride FontSize='0,25'>PLANIFICAÇÃO</StyleOverride>"

        For Each oView In createdViews
            Try
                oView.ShowLabel = True
                oView.Label.FormattedText = If(oView.IsFlatPatternView, oPartNumber & oListItem & Planificacao & oQty & oStringScale, oPartNumber & oListItem & oQty & oStringScale)
            Catch ex As Exception
                Log($"ERRO ao aplicar rótulo na vista {oView.Name}: {ex.Message}")
            End Try
        Next
    End Sub

    Private Sub SetOrAddiProperty(propSet As PropertySet, propName As String, propValue As Object)
        Try
            propSet.Item(propName).Value = propValue
        Catch
            propSet.Add(propValue, propName)
        End Try
    End Sub

    Private Sub SetOrAddAttribute(attrSet As AttributeSet, attrName As String, attrValue As String)
        Try
            attrSet.Item(attrName).Value = attrValue
        Catch
            attrSet.Add(attrName, ValueTypeEnum.kStringType, attrValue)
        End Try
    End Sub

    Private Sub AddBendNotes(flatView As DrawingView, oSheet As Sheet)
        Try
            For Each oCurve As DrawingCurve In flatView.DrawingCurves
                If oCurve.EdgeType = DrawingEdgeTypeEnum.kBendUpEdge OrElse oCurve.EdgeType = DrawingEdgeTypeEnum.kBendDownEdge Then
                    oSheet.DrawingNotes.BendNotes.Add(oCurve)
                End If
            Next
            Log($"   -> Notas de dobra adicionadas para a vista {flatView.Name}")
        Catch ex As Exception
            Log($"   -> ERRO ao adicionar notas de dobra: {ex.Message}")
        End Try
    End Sub

    Private Sub Log(message As String)
        _logBuilder.AppendLine($"{DateTime.Now:HH:mm:ss} - {message}")
    End Sub

    Private Sub FinalizeLog()
        LogTextBox.Text = _logBuilder.ToString()
        LogTextBox.ScrollToEnd()
    End Sub

    '===========================================================
    '#End Region
    '===========================================================

    '===========================================================
    '#Region "Eventos de UI (arrastar, fechar, etc.)"
    '===========================================================
    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    '===========================================================
    '#End Region
    '===========================================================
End Class