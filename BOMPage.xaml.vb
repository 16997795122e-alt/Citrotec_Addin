' Arquivo: BOMPage.xaml.vb (CORREÇÃO FINAL E DEFINITIVA)
Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports Inventor
Imports System.IO

Public Class BOMPage
    Private _activeDoc As Document
    Private _sourceAssemblyDoc As AssemblyDocument
    Public Property BOMEntries As New ObservableCollection(Of BOMEntry)

    Private ReadOnly logFilePath As String = "C:\Temp\CitrotecAddin_Log.txt"

    Public Sub New()
        InitializeComponent()
        Me.DataContext = Me
        If System.IO.File.Exists(logFilePath) Then System.IO.File.Delete(logFilePath)
        Log("================ INICIANDO NOVA SESSÃO ================")
    End Sub

    Private Sub Log(message As String)
        Try
            Dim logDirectory = System.IO.Path.GetDirectoryName(logFilePath)
            If Not System.IO.Directory.Exists(logDirectory) Then
                System.IO.Directory.CreateDirectory(logDirectory)
            End If
            System.IO.File.AppendAllText(logFilePath, $"{DateTime.Now:dd/MM/yyyy HH:mm:ss.fff} - {message}{vbCrLf}")
        Catch ex As Exception
        End Try
    End Sub

    Public Sub UpdateActiveDocument(ByVal doc As Document)
        Log("UpdateActiveDocument chamado.")
        _activeDoc = doc
        _sourceAssemblyDoc = Nothing
        BOMEntries.Clear()
        If _activeDoc Is Nothing Then
            StatusTextBlock.Text = "Nenhum documento ativo."
            SetButtonsState(canLoad:=False, canSave:=False, canFormat:=False)
        ElseIf _activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            StatusTextBlock.Text = "Pronto para carregar BOM da montagem: " & System.IO.Path.GetFileName(_activeDoc.FullFileName)
            SetButtonsState(canLoad:=True, canSave:=False, canFormat:=True)
            FormatBOMButton.Content = "Formatar BOM Montagem"
        ElseIf _activeDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            StatusTextBlock.Text = "Pronto para carregar BOM do desenho: " & System.IO.Path.GetFileName(_activeDoc.FullFileName)
            SetButtonsState(canLoad:=True, canSave:=False, canFormat:=True)
            FormatBOMButton.Content = "Formatar BOM Desenho"
        Else
            StatusTextBlock.Text = "Abra uma montagem (.iam) ou desenho (.idw) para começar."
            SetButtonsState(canLoad:=False, canSave:=False, canFormat:=False)
        End If
    End Sub

    Private Sub LoadBOMButton_Click(sender As Object, e As RoutedEventArgs)
        Log("Botão 'Carregar BOM' clicado.")
        BOMEntries.Clear()
        Try
            If _activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                _sourceAssemblyDoc = CType(_activeDoc, AssemblyDocument)
            ElseIf _activeDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                Dim drawingDoc As DrawingDocument = CType(_activeDoc, DrawingDocument)
                If drawingDoc.ActiveSheet.PartsLists.Count = 0 Then
                    MessageBox.Show("Não foi encontrada nenhuma Lista de Peças (Parts List) na folha ativa do desenho.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning)
                    Return
                End If
                Dim partsList As PartsList = drawingDoc.ActiveSheet.PartsLists.Item(1)
                _sourceAssemblyDoc = CType(partsList.ReferencedDocumentDescriptor.ReferencedDocument, AssemblyDocument)
            Else
                Return
            End If
            StatusTextBlock.Text = "Carregando BOM de " & System.IO.Path.GetFileName(_sourceAssemblyDoc.FullFileName)
            Dim bom As BOM = _sourceAssemblyDoc.ComponentDefinition.BOM
            bom.StructuredViewEnabled = True
            Dim bomView As BOMView = bom.BOMViews.Item("Structured")
            For Each row As BOMRow In bomView.BOMRows
                If row.ComponentDefinitions.Count = 0 OrElse TypeOf row.ComponentDefinitions.Item(1) Is VirtualComponentDefinition Then Continue For
                Dim primaryDef As ComponentDefinition = row.ComponentDefinitions.Item(1)
                Dim doc As Document = primaryDef.Document
                Dim customProps = GetPropertySet(doc, "User Defined Properties")
                Dim designProps = GetPropertySet(doc, "Design Tracking Properties")
                Dim itemNumberAsInt As Integer = 0
                Integer.TryParse(row.ItemNumber, itemNumberAsInt)
                Dim bomQty As Integer = 0
                Integer.TryParse(row.TotalQuantity, bomQty)
                Dim customQtyString As String = GetPropertyValue(customProps, "QTD")
                Dim displayQty As Integer
                If String.IsNullOrEmpty(customQtyString) Then
                    displayQty = bomQty
                    SetPropertyValue(customProps, "QTD", displayQty.ToString())
                Else
                    Integer.TryParse(customQtyString, displayQty)
                End If
                Dim unitQtyValue As String = ""
                Try
                    Dim uom As UnitsOfMeasure = doc.UnitsOfMeasure
                    Dim qType As BOMQuantityTypeEnum
                    Dim qQty As Object
                    primaryDef.BOMQuantity.GetBaseQuantity(qType, qQty)
                    If qType = BOMQuantityTypeEnum.kParameterBOMQuantity AndAlso qQty IsNot Nothing Then
                        Dim qParam As Parameter = CType(qQty, Parameter)
                        unitQtyValue = Math.Round(uom.ConvertUnits(qParam.ModelValue, qParam.Units, "mm"), 2).ToString() & " mm"
                    ElseIf qType = BOMQuantityTypeEnum.kEachBOMQuantity Then
                        unitQtyValue = "Each"
                    End If
                Catch ex As Exception
                    unitQtyValue = ""
                End Try
                Dim newEntry As New BOMEntry With {
                    .OriginalRow = row,
                    .Item = itemNumberAsInt,
                    .PartNumber = GetPropertyValue(designProps, "Part Number"),
                    .UnitQty = unitQtyValue,
                    .Qty = row.TotalQuantity.ToString(),
                    .Qtd = displayQty,
                    .Codigo = GetPropertyValue(customProps, "Código"),
                    .Descricao = GetPropertyValue(customProps, "Descrição"),
                    .DIMENSAO1 = GetPropertyValue(customProps, "DIMENSAO1"),
                    .DIMENSAO2 = GetPropertyValue(customProps, "DIMENSAO2"),
                    .COMPLEMENTO = GetPropertyValue(customProps, "COMPLEMENTO"),
                    .Peso = "",
                    .QD = ""
                }
                BOMEntries.Add(newEntry)
            Next
            Dim sortedEntries = BOMEntries.OrderBy(Function(entry) entry.Item).ToList()
            BOMEntries.Clear()
            For Each entry In sortedEntries
                BOMEntries.Add(entry)
            Next
            StatusTextBlock.Text = BOMEntries.Count & " itens carregados. Pronto para editar."
            SetButtonsState(canLoad:=True, canSave:=True, canFormat:=True)
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro ao carregar a BOM: " & ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
            SetButtonsState(canLoad:=True, canSave:=False, canFormat:=True)
        End Try
    End Sub

    Private Sub SaveChangesButton_Click(sender As Object, e As RoutedEventArgs)
        Log("Botão 'Salvar Alterações' clicado.")
        If Not BOMEntries.Any() Then
            MessageBox.Show("Não há dados na BOM para salvar.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Information)
            Return
        End If

        Try
            For Each entry As BOMEntry In BOMEntries
                Dim primaryDef As ComponentDefinition = entry.OriginalRow.ComponentDefinitions.Item(1)
                Dim doc As Document = primaryDef.Document

                SetPropertyValue(GetPropertySet(doc, "User Defined Properties"), "QTD", entry.Qtd.ToString())
                SetPropertyValue(GetPropertySet(doc, "Design Tracking Properties"), "Part Number", entry.PartNumber)
                SetPropertyValue(GetPropertySet(doc, "User Defined Properties"), "Código", entry.Codigo)
                SetPropertyValue(GetPropertySet(doc, "User Defined Properties"), "Descrição", entry.Descricao)
                SetPropertyValue(GetPropertySet(doc, "User Defined Properties"), "DIMENSAO1", entry.DIMENSAO1)
                SetPropertyValue(GetPropertySet(doc, "User Defined Properties"), "DIMENSAO2", entry.DIMENSAO2)
                SetPropertyValue(GetPropertySet(doc, "User Defined Properties"), "COMPLEMENTO", entry.COMPLEMENTO)
            Next

            Log("Loop de salvamento concluído com sucesso. Atualizando documentos.")
            _sourceAssemblyDoc.Update()
            If _activeDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                _activeDoc.Update()
            End If
            MessageBox.Show(BOMEntries.Count & " itens processados e salvos com sucesso!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)

        Catch ex As Exception
            Log("ERRO: " & ex.ToString()) ' Log completo da exceção
            MessageBox.Show("Ocorreu um erro ao salvar as alterações: " & ex.Message & vbCrLf & "(0x" & Hex(ex.HResult) & ")", "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

    Private Sub FormatBOMButton_Click(sender As Object, e As RoutedEventArgs)
        Dim layoutFilePath As String = "\\SERVIDOR20\eng\Engenharia\iLogic\REGRAS\Lista Idw\LayoutLista.xml"
        If Not System.IO.File.Exists(layoutFilePath) Then
            MessageBox.Show("Arquivo de layout não encontrado em: " & layoutFilePath, "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
            Return
        End If
        Try
            If _activeDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
                Dim drawingDoc As DrawingDocument = CType(_activeDoc, DrawingDocument)
                If drawingDoc.ActiveSheet.PartsLists.Count = 0 Then
                    MessageBox.Show("Não foi encontrada nenhuma Lista de Peças na folha ativa para formatar.", "Aviso", MessageBoxButton.OK, MessageBoxImage.Warning)
                    Return
                End If
                Dim partsList As PartsList = drawingDoc.ActiveSheet.PartsLists.Item(1)
                partsList.Style = drawingDoc.StylesManager.PartsListStyles.Item("CITROTEC-EDIT")
                Dim refDoc As AssemblyDocument = CType(partsList.ReferencedDocumentDescriptor.ReferencedDocument, AssemblyDocument)
                refDoc.ComponentDefinition.BOM.ImportBOMCustomization(layoutFilePath)
                refDoc.Update()
                drawingDoc.Update()
                MessageBox.Show("Estilo da Lista de Peças alterado para 'CITROTEC-EDIT' e BOM da montagem formatada.", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
            ElseIf _sourceAssemblyDoc IsNot Nothing Then
                _sourceAssemblyDoc.ComponentDefinition.BOM.ImportBOMCustomization(layoutFilePath)
                _sourceAssemblyDoc.Update()
                MessageBox.Show("Formatação da BOM aplicada com sucesso na montagem!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
            ElseIf _activeDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                Dim asmDoc As AssemblyDocument = CType(_activeDoc, AssemblyDocument)
                asmDoc.ComponentDefinition.BOM.ImportBOMCustomization(layoutFilePath)
                asmDoc.Update()
                MessageBox.Show("Formatação da BOM aplicada com sucesso na montagem!", "Sucesso", MessageBoxButton.OK, MessageBoxImage.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro ao formatar: " & ex.Message, "Erro", MessageBoxButton.OK, MessageBoxImage.Error)
        End Try
    End Sub

#Region "Funções Auxiliares"
    Private Sub SetButtonsState(canLoad As Boolean, canSave As Boolean, canFormat As Boolean)
        LoadBOMButton.IsEnabled = canLoad
        SaveChangesButton.IsEnabled = canSave
        FormatBOMButton.IsEnabled = canFormat
    End Sub
    Private Function GetPropertySet(doc As Document, setName As String) As PropertySet
        Try
            Return doc.PropertySets.Item(setName)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Private Function GetPropertyValue(propSet As PropertySet, propName As String) As String
        If propSet Is Nothing Then Return ""
        Try
            Return propSet.Item(propName).Value?.ToString()
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Private Sub SetPropertyValue(propSet As PropertySet, propName As String, value As Object)
        If propSet Is Nothing Then Return

        ' ====================== CORREÇÃO DEFINITIVA AQUI ======================
        ' Se o valor a ser salvo for nulo (Nothing), substitui por uma string vazia.
        If value Is Nothing Then
            value = String.Empty
        End If
        ' ====================================================================

        Dim existingProp As Inventor.Property = Nothing
        Try
            existingProp = propSet.Item(propName)
        Catch
        End Try

        If existingProp IsNot Nothing Then
            If existingProp.Value Is Nothing OrElse Not existingProp.Value.ToString().Equals(value.ToString()) Then
                existingProp.Value = value
            End If
        Else
            propSet.Add(value, propName)
        End If
    End Sub

#End Region

End Class