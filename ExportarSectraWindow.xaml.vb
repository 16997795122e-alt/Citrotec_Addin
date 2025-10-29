' Arquivo: ExportarSectraWindow.xaml.vb
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Inventor
Imports CitrotecAddin
Imports Excel = Microsoft.Office.Interop.Excel

Namespace CitrotecAddin

    Partial Public Class ExportarSectraWindow
        Inherits System.Windows.Window

        Private _logBuilder As New StringBuilder()
        Private _invApp As Inventor.Application

        Public Sub New()
            InitializeComponent()
            ' O código Try...Catch para carregar o vídeo foi removido daqui
        End Sub

        Private Sub Border_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
            Me.DragMove()
        End Sub

        Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
            Me.Close()
        End Sub

        'Private Sub CommandVideo_MediaEnded(sender As Object, e As RoutedEventArgs)
        ' Reinicia a posição do vídeo para o início e toca novamente
        ' Dim mediaElement = TryCast(sender, MediaElement) ' VARIÁVEL RENOMEADA

        'erifica a variável correta (mediaElement)
        'If mediaElement IsNot Nothing Then
        '     mediaElement.Position = TimeSpan.Zero ' Usa a variável correta
        '    mediaElement.Play() ' Usa a variável correta
        'End If
        ' End Sub

        Private Sub btnExecute_Click(sender As Object, e As RoutedEventArgs)
            _logBuilder.Clear()
            LogTextBox.Clear()
            Log("Iniciando exportação SECTRA para Excel (via COM)...")

            _invApp = Globals.g_inventorApplication
            Dim oDrawDoc As DrawingDocument = Nothing
            Dim excelFilePath As String = ""
            Dim saveFolder As String = "C:\Sctra\acad\ListaMat"

            Dim excelApp As Excel.Application = Nothing
            Dim excelWorkbook As Excel.Workbook = Nothing
            Dim createdSheetNames As New List(Of String)

            Try
                ' 1. Validação Inicial
                oDrawDoc = TryCast(_invApp.ActiveDocument, DrawingDocument)
                If oDrawDoc Is Nothing OrElse oDrawDoc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                    Log("ERRO: O documento ativo não é um desenho (.idw).")
                    FinalizeLog()
                    Return
                End If
                Log("Documento de desenho ativo: " & System.IO.Path.GetFileName(oDrawDoc.FullFileName))

                ' 2. Determinar nome do arquivo e verificar existência
                Dim partNumber As String = GetPropertyValueFromDoc(oDrawDoc, "Part Number", "Design Tracking Properties")?.ToString()
                If String.IsNullOrWhiteSpace(partNumber) Then
                    Log("ERRO: iProperty 'Part Number' (Número da Peça) não encontrada ou vazia no desenho.")
                    FinalizeLog()
                    Return
                End If
                Log("Part Number do desenho: " & partNumber)

                If Not Directory.Exists(saveFolder) Then
                    Try
                        Directory.CreateDirectory(saveFolder)
                        Log($"Pasta de destino '{saveFolder}' não existia e foi criada.")
                    Catch exDir As Exception
                        Log($"ERRO: Falha ao criar a pasta de destino '{saveFolder}'. {exDir.Message}")
                        FinalizeLog()
                        Return
                    End Try
                End If

                excelFilePath = System.IO.Path.Combine(saveFolder, $"{partNumber}.xls")
                Log("Caminho do arquivo Excel de destino: " & excelFilePath)

                If System.IO.File.Exists(excelFilePath) Then
                    Dim msgBoxConfirm = New MessageBoxWindow($"O Arquivo já existe:{vbCrLf}{excelFilePath}{vbCrLf}{vbCrLf}Deseja Substituir?", "Arquivo Existente", ModernMessageBoxButtons.YesNo)
                    msgBoxConfirm.Owner = Me
                    msgBoxConfirm.ShowDialog()
                    If msgBoxConfirm.Result = MessageBoxResult.Yes Then
                        Try
                            System.IO.File.Delete(excelFilePath)
                            Log("Arquivo existente excluído.")
                        Catch exDel As Exception
                            Log($"ERRO: Não foi possível excluir o arquivo existente. {exDel.Message}")
                            FinalizeLog()
                            Return
                        End Try
                    Else
                        Log("Exportação cancelada pelo usuário (arquivo existente).")
                        FinalizeLog()
                        Return
                    End If
                End If

                ' 3. Iniciar Excel e Criar Workbook
                Log("Iniciando instância do Excel...")
                Try
                    excelApp = New Excel.Application()
                    excelApp.Visible = False
                    excelApp.DisplayAlerts = False
                Catch exExcel As Exception
                    Throw New Exception("Não foi possível iniciar o Microsoft Excel. Verifique se ele está instalado corretamente. Detalhes: " & exExcel.Message)
                End Try

                Log("Criando novo arquivo Excel...")
                excelWorkbook = excelApp.Workbooks.Add()

                Dim defaultSheetNames As New List(Of String)
                Dim wsTemp As Excel.Worksheet = Nothing
                For i As Integer = 1 To excelWorkbook.Worksheets.Count
                    Try
                        wsTemp = CType(excelWorkbook.Worksheets(i), Excel.Worksheet)
                        defaultSheetNames.Add(wsTemp.Name)
                    Finally
                        ReleaseComObject(wsTemp)
                    End Try
                Next
                Log($"Planilha(s) padrão inicial(is): {String.Join(", ", defaultSheetNames)}")

                ' 4. Exportar Nossas Planilhas
                Log("Adicionando planilhas de dados...")
                Dim plName = ExportPartsListCOM(oDrawDoc, excelWorkbook)
                If Not String.IsNullOrEmpty(plName) Then createdSheetNames.Add(plName)

                Dim revName = ExportRevisionTableCOM(oDrawDoc, excelWorkbook)
                If Not String.IsNullOrEmpty(revName) Then createdSheetNames.Add(revName)

                Dim propName = ExportDrawingPropertiesCOM(oDrawDoc, excelWorkbook)
                If Not String.IsNullOrEmpty(propName) Then createdSheetNames.Add(propName)

                ' Deleta planilhas extras (lógica segura)
                Log("Removendo planilhas padrão não utilizadas...")
                excelApp.DisplayAlerts = False
                Dim sheetToDelete As Excel.Worksheet = Nothing
                For i As Integer = excelWorkbook.Worksheets.Count To 1 Step -1
                    Try
                        sheetToDelete = CType(excelWorkbook.Worksheets(i), Excel.Worksheet)
                        If defaultSheetNames.Contains(sheetToDelete.Name, StringComparer.OrdinalIgnoreCase) AndAlso
                           Not createdSheetNames.Contains(sheetToDelete.Name, StringComparer.OrdinalIgnoreCase) Then
                            If excelWorkbook.Worksheets.Count > 1 Then
                                Log($"   -> Removendo planilha padrão '{sheetToDelete.Name}'...")
                                sheetToDelete.Delete()
                            Else
                                Log($"   -> Ignorando exclusão da última planilha restante '{sheetToDelete.Name}'.")
                            End If
                        End If
                    Catch exDel As Exception
                        Log($"   -> Erro ao tentar processar/deletar planilha #{i} ('{sheetToDelete?.Name}'): {exDel.Message}")
                    Finally
                        ReleaseComObject(sheetToDelete)
                        sheetToDelete = Nothing
                    End Try
                Next
                excelApp.DisplayAlerts = True

                ' 7. Salvar e Fechar Excel
                If excelWorkbook.Worksheets.Count = 0 Then
                    Log("ERRO: Nenhuma planilha foi criada/mantida no arquivo Excel. Salvamento cancelado.")
                    Throw New Exception("Nenhuma planilha utilizável criada.")
                End If

                Log("Salvando arquivo Excel: " & excelFilePath)
                excelWorkbook.SaveAs(excelFilePath, Excel.XlFileFormat.xlExcel8)
                Log("Arquivo Excel salvo com sucesso!")

                excelWorkbook.Close(False)
                excelApp.Quit()

                ' 8. Abrir Pasta
                Dim msgBoxOpen = New MessageBoxWindow($"Informações exportadas com sucesso para:{vbCrLf}{excelFilePath}{vbCrLf}{vbCrLf}Gostaria de abrir a pasta?", "CITROTEC - Exportação Concluída", ModernMessageBoxButtons.YesNo)
                msgBoxOpen.Owner = Me
                msgBoxOpen.ShowDialog()
                If msgBoxOpen.Result = MessageBoxResult.Yes Then
                    Try
                        Process.Start("explorer.exe", saveFolder)
                    Catch exOpen As Exception
                        Log("Erro ao tentar abrir a pasta: " & exOpen.Message)
                    End Try
                End If

                Log("-------------------------------------")
                Log("Exportação SECTRA (COM) finalizada com sucesso.")

            Catch ex As Exception
                Log("======================================")
                Log("ERRO GERAL INESPERADO DURANTE A EXPORTAÇÃO: " & ex.Message)
                If ex.InnerException IsNot Nothing Then Log("   Inner Exception: " & ex.InnerException.Message)
                Log(ex.StackTrace)
                Log("======================================")
                Try
                    If excelWorkbook IsNot Nothing Then excelWorkbook.Close(False)
                    If excelApp IsNot Nothing Then excelApp.Quit()
                Catch exClose As Exception
                    Log("Erro ao tentar fechar o Excel após falha: " & exClose.Message)
                End Try
            Finally
                ReleaseComObject(excelWorkbook)
                If excelApp IsNot Nothing Then
                    ReleaseComObject(excelApp)
                    GC.Collect()
                    GC.WaitForPendingFinalizers()
                End If
                FinalizeLog()
            End Try
        End Sub ' Fim do btnExecute_Click

        ' --- Sub-rotina para exportar PartsList usando COM ---
        Private Function ExportPartsListCOM(oDrawDoc As DrawingDocument, excelWorkbook As Excel.Workbook) As String
            Dim ws As Excel.Worksheet = Nothing
            Dim oPartslist As PartsList = Nothing
            Dim finalSheetName As String = ""
            Try
                Dim processedAssemblies As New List(Of String)
                Dim baseSheetName As String = "LISTA DE MATERIAIS"
                Dim sheetName As String = baseSheetName
                Dim sheetIndex As Integer = 1

                For Each oSheet As Sheet In oDrawDoc.Sheets
                    For Each pl As PartsList In oSheet.PartsLists
                        If pl.ReferencedDocumentDescriptor IsNot Nothing Then
                            Dim refDocName = pl.ReferencedDocumentDescriptor.FullDocumentName
                            If Not processedAssemblies.Contains(refDocName) Then
                                oPartslist = pl
                                processedAssemblies.Add(refDocName)
                                While WorksheetExists(excelWorkbook, sheetName)
                                    sheetName = $"{baseSheetName} BOM {sheetIndex}"
                                    sheetIndex += 1
                                End While
                                finalSheetName = sheetName
                                Exit For
                            End If
                        End If
                    Next
                    If oPartslist IsNot Nothing Then Exit For
                Next

                If oPartslist Is Nothing Then
                    Log("   -> Nenhuma Lista de Materiais encontrada ou aplicável no desenho.")
                    Return ""
                End If

                Log($"   -> Exportando PartsList da folha '{oPartslist.Parent.Name}' para aba '{finalSheetName}'...")
                ws = CType(excelWorkbook.Worksheets.Add(After:=excelWorkbook.Sheets(excelWorkbook.Sheets.Count)), Excel.Worksheet)
                ws.Name = finalSheetName

                ws.Cells(1, 1).Value = "POSIÇÃO"
                ws.Cells(1, 2).Value = "LEGENDA (D/M)"
                ws.Cells(1, 3).Value = "DESENHO/MATERIAL"
                ws.Cells(1, 4).Value = "QUANTIDADE"
                ws.Cells(1, 5).Value = "COMPRIMENTO"
                ws.Cells(1, 6).Value = "LARGURA"
                ws.Cells(1, 7).Value = "PESO LÍQUIDO"
                ws.Cells(1, 8).Value = "PESO BRUTO"

                ws.Columns("A:A").NumberFormat = "@"
                ws.Columns("C:H").NumberFormat = "@"

                Dim columnsToRead As String() = {"ITEM", "Código", "Qtd", "DIMENSAO1", "DIMENSAO2", "PESO_UNIT_LIQUIDO", "PESO_UNIT_BRUTO"}
                Dim targetColumns As Integer() = {1, 3, 4, 5, 6, 7, 8}

                Dim currentRow As Integer = 2
                For Each row As PartsListRow In oPartslist.PartsListRows
                    For i As Integer = 0 To columnsToRead.Length - 1
                        Try
                            Dim cellValue As String = row.Item(columnsToRead(i)).Value
                            ws.Cells(currentRow, targetColumns(i)).Value = cellValue
                        Catch exCol As Exception
                        End Try
                    Next
                    ws.Cells(currentRow, 2).Formula = $"=IF(LEN(C{currentRow})>=8, ""M"", ""D"")"
                    currentRow += 1
                Next

                ws.UsedRange.Columns.AutoFit()
                Log($"   -> {currentRow - 2} linhas exportadas.")
                Return finalSheetName

            Catch ex As Exception
                Log($"   -> ERRO ao exportar Lista de Materiais: {ex.Message}")
                Return ""
            Finally
                ReleaseComObject(ws)
            End Try
        End Function ' Fim do ExportPartsListCOM

        ' --- Sub-rotina para exportar RevisionTable usando COM ---
        Private Function ExportRevisionTableCOM(oDrawDoc As DrawingDocument, excelWorkbook As Excel.Workbook) As String
            Dim ws As Excel.Worksheet = Nothing
            Dim oRevTable As RevisionTable = Nothing
            Dim finalSheetName As String = "REVISAO"
            Try
                Try
                    oRevTable = oDrawDoc.ActiveSheet.RevisionTables.Item(1)
                Catch exGetTable As Exception
                    Log("   -> Nenhuma Tabela de Revisão encontrada na folha ativa.")
                    Return ""
                End Try

                Log($"   -> Exportando Tabela de Revisão para aba '{finalSheetName}'...")
                If WorksheetExists(excelWorkbook, finalSheetName) Then
                    Dim wsToDelete As Excel.Worksheet = CType(excelWorkbook.Worksheets(finalSheetName), Excel.Worksheet)
                    wsToDelete.Delete()
                    ReleaseComObject(wsToDelete)
                End If

                ws = CType(excelWorkbook.Worksheets.Add(After:=excelWorkbook.Sheets(excelWorkbook.Sheets.Count)), Excel.Worksheet)
                ws.Name = finalSheetName

                For col As Integer = 1 To oRevTable.RevisionTableColumns.Count
                    ws.Cells(1, col).Value = oRevTable.RevisionTableColumns.Item(col).Title
                Next

                For row As Integer = 1 To oRevTable.RevisionTableRows.Count
                    For col As Integer = 1 To oRevTable.RevisionTableColumns.Count
                        ws.Cells(row + 1, col).Value = oRevTable.RevisionTableRows.Item(row).Item(col).Text
                    Next
                Next

                ws.UsedRange.Columns.AutoFit()
                Log($"   -> {oRevTable.RevisionTableRows.Count} revisões exportadas.")
                Return finalSheetName

            Catch ex As Exception
                Log($"   -> ERRO ao exportar Tabela de Revisão: {ex.Message}")
                Return ""
            Finally
                ReleaseComObject(ws)
            End Try
        End Function ' Fim do ExportRevisionTableCOM

        ' --- Sub-rotina para exportar Propriedades usando COM ---
        ' ***** VERSÃO FINAL CORRIGIDA - REVISÃO E DATAS *****
        Private Function ExportDrawingPropertiesCOM(oDrawDoc As DrawingDocument, excelWorkbook As Excel.Workbook) As String
            Dim ws As Excel.Worksheet = Nothing
            Dim finalSheetName As String = "PROPRIEDADES"
            Try
                Log($"   -> Exportando Propriedades para aba '{finalSheetName}'...")
                If WorksheetExists(excelWorkbook, finalSheetName) Then
                    Dim wsToDelete As Excel.Worksheet = CType(excelWorkbook.Worksheets(finalSheetName), Excel.Worksheet)
                    wsToDelete.Delete()
                    ReleaseComObject(wsToDelete)
                End If

                ws = CType(excelWorkbook.Worksheets.Add(After:=excelWorkbook.Sheets(excelWorkbook.Sheets.Count)), Excel.Worksheet)
                ws.Name = finalSheetName

                ' Escreve Rótulos (Coluna A)
                ws.Cells(1, 1).Value = "OT:"
                ws.Cells(2, 1).Value = "CLIENTE:"
                ws.Cells(3, 1).Value = "NUMERO DESENHO:"
                ws.Cells(4, 1).Value = "TITULO 1:"
                ws.Cells(5, 1).Value = "TITULO 2:"
                ws.Cells(6, 1).Value = "TITULO 3:"
                ws.Cells(7, 1).Value = "REVISÃO:"
                ws.Cells(8, 1).Value = "ESCALA:"
                ws.Cells(9, 1).Value = "FOLHAS:"
                ws.Cells(10, 1).Value = "DESENHADO POR:"
                ws.Cells(11, 1).Value = "DATA DESENHADO:"
                ws.Cells(12, 1).Value = "VERIFICADO POR:"
                ws.Cells(13, 1).Value = "DATA VERIFICADO:"
                ws.Cells(14, 1).Value = "APROVADO POR:"
                ws.Cells(15, 1).Value = "DATA APROVADO:"

                ' Escreve Valores (Coluna B)
                ws.Cells(1, 2).Value = GetPropertyValueFromDoc(oDrawDoc, "NROOT", "Inventor User Defined Properties")?.ToString()
                ws.Cells(2, 2).Value = GetPropertyValueFromDoc(oDrawDoc, "Company", "Inventor Document Summary Information")?.ToString()
                ws.Cells(3, 2).Value = GetPropertyValueFromDoc(oDrawDoc, "NRODESENHO", "Inventor User Defined Properties")?.ToString()
                ws.Cells(4, 2).Value = GetPropertyValueFromDoc(oDrawDoc, "TITULO1", "Inventor User Defined Properties")?.ToString()
                ws.Cells(5, 2).Value = GetPropertyValueFromDoc(oDrawDoc, "TITULO2", "Inventor User Defined Properties")?.ToString()
                ws.Cells(6, 2).Value = GetPropertyValueFromDoc(oDrawDoc, "TITULO3", "Inventor User Defined Properties")?.ToString()
                ' CORREÇÃO: Busca Revision Number do Summary Information
                ws.Cells(7, 2).Value = GetPropertyValueFromDoc(oDrawDoc, "Revision Number", "Inventor Summary Information")?.ToString()
                ws.Cells(8, 2).Value = GetPropertyValueFromDoc(oDrawDoc, "ESCALA", "Inventor User Defined Properties")?.ToString()
                ws.Cells(9, 2).Value = oDrawDoc.Sheets.Count.ToString()
                ws.Cells(10, 2).Value = GetPropertyValueFromDoc(oDrawDoc, "Authority", "Design Tracking Properties")?.ToString()

                ' Trata Datas para formato correto
                Dim creationDateObj = GetPropertyValueFromDoc(oDrawDoc, "Creation Time", "Design Tracking Properties")
                If TypeOf creationDateObj Is Date Then ws.Cells(11, 2).Value = CDate(creationDateObj) Else ws.Cells(11, 2).Value = ""
                ws.Cells(11, 2).NumberFormat = "dd/MM/yyyy"

                ws.Cells(12, 2).Value = "R.C.A." ' Fixo

                Dim checkedDateObj = GetPropertyValueFromDoc(oDrawDoc, "Date Checked", "Design Tracking Properties")
                If TypeOf checkedDateObj Is Date Then ws.Cells(13, 2).Value = CDate(checkedDateObj) Else ws.Cells(13, 2).Value = ""
                ws.Cells(13, 2).NumberFormat = "dd/MM/yyyy"

                ws.Cells(14, 2).Value = "P.H.S." ' Fixo

                Dim approvedDateObj = GetPropertyValueFromDoc(oDrawDoc, "Engr Date Approved", "Design Tracking Properties")
                If TypeOf approvedDateObj Is Date Then ws.Cells(15, 2).Value = CDate(approvedDateObj) Else ws.Cells(15, 2).Value = ""
                ws.Cells(15, 2).NumberFormat = "dd/MM/yyyy"

                ' Aplica formato texto APÓS escrever os valores
                ws.Cells(1, 2).NumberFormat = "@" ' OT
                ws.Cells(2, 2).NumberFormat = "@" ' CLIENTE
                ws.Cells(3, 2).NumberFormat = "@" ' NUMERO DESENHO
                ws.Cells(7, 2).NumberFormat = "@" ' REVISÃO
                ws.Cells(9, 2).NumberFormat = "@" ' FOLHAS

                ws.UsedRange.Columns.AutoFit()
                Log($"   -> Propriedades exportadas.")
                Return finalSheetName

            Catch ex As Exception
                Log($"   -> ERRO ao exportar Propriedades: {ex.Message} (Linha: {New StackTrace(ex, True).GetFrame(0).GetFileLineNumber()})")
                Return ""
            Finally
                ReleaseComObject(ws)
            End Try
        End Function ' Fim do ExportDrawingPropertiesCOM

        ' --- Função Auxiliar GetPropertyValueFromDoc ---
        ' ***** CORREÇÃO: Ajuste no mapeamento de Revision Number *****
        Private Function GetPropertyValueFromDoc(doc As Document, propName As String, Optional defaultPropSetName As String = "Inventor User Defined Properties") As Object
            Dim standardSets As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"Part Number", "Design Tracking Properties"},
                {"Revision Number", "Inventor Summary Information"}, ' CORRIGIDO para Summary
                {"Project", "Design Tracking Properties"},
                {"Authority", "Design Tracking Properties"},
                {"Designer", "Design Tracking Properties"},
                {"Engineer", "Design Tracking Properties"},
                {"Creation Time", "Design Tracking Properties"},
                {"Checked By", "Design Tracking Properties"},
                {"Date Checked", "Design Tracking Properties"},
                {"Engr Approved By", "Design Tracking Properties"},
                {"Engr Date Approved", "Design Tracking Properties"},
                {"Mfg Approved By", "Design Tracking Properties"},
                {"Mfg Date Approved", "Design Tracking Properties"},
                {"Company", "Inventor Document Summary Information"},
                {"Category", "Inventor Document Summary Information"},
                {"Manager", "Inventor Document Summary Information"},
                {"Author", "Inventor Summary Information"},
                {"Title", "Inventor Summary Information"},
                {"Subject", "Inventor Summary Information"},
                {"Keywords", "Inventor Summary Information"},
                {"Comments", "Inventor Summary Information"},
                {"Last Saved By", "Inventor Summary Information"}
            }

            Dim targetSetName As String = defaultPropSetName
            If standardSets.ContainsKey(propName) Then
                targetSetName = standardSets(propName)
            ElseIf defaultPropSetName = "Inventor User Defined Properties" Then
                targetSetName = "Inventor User Defined Properties"
            End If

            Try
                Dim propSet As PropertySet = doc.PropertySets.Item(targetSetName)
                Dim prop As Inventor.Property = propSet.Item(propName)

                If prop IsNot Nothing AndAlso prop.Value IsNot Nothing Then
                    If TypeOf prop.Value Is Date Then
                        Dim dtValue As Date = CDate(prop.Value)
                        If dtValue.Year > 1601 Then
                            Return dtValue ' Retorna o objeto Date
                        Else
                            Return ""
                        End If
                    End If
                    Return prop.Value ' Retorna o valor original (pode ser string, numero, etc.)
                Else
                    Return ""
                End If
            Catch ex As Exception
                Return ""
            End Try
        End Function ' Fim do GetPropertyValueFromDoc

        ' --- Função Auxiliar WorksheetExists ---
        Private Function WorksheetExists(workbook As Excel.Workbook, sheetName As String) As Boolean
            Dim ws As Excel.Worksheet = Nothing
            Dim exists As Boolean = False
            Try
                ws = TryCast(workbook.Sheets(sheetName), Excel.Worksheet)
                exists = (ws IsNot Nothing)
            Catch ex As Exception
                exists = False
            Finally
                ReleaseComObject(ws)
            End Try
            Return exists
        End Function ' Fim do WorksheetExists

        ' --- Função Auxiliar ReleaseComObject ---
        Private Sub ReleaseComObject(ByVal obj As Object)
            Try
                If obj IsNot Nothing Then
                    While Marshal.ReleaseComObject(obj) > 0
                    End While
                End If
            Catch ex As Exception
            Finally
                obj = Nothing
            End Try
        End Sub ' Fim do ReleaseComObject

        ' --- Funções Log e FinalizeLog ---
        Private Sub Log(message As String)
            _logBuilder.AppendLine($"{DateTime.Now:HH:mm:ss} - {message}")
            Me.Dispatcher.Invoke(Sub() LogTextBox.Text = _logBuilder.ToString())
        End Sub ' Fim do Log

        Private Sub FinalizeLog()
            Me.Dispatcher.Invoke(Sub()
                                     LogTextBox.Text = _logBuilder.ToString()
                                     LogTextBox.ScrollToEnd()
                                 End Sub)
        End Sub ' Fim do FinalizeLog

    End Class ' Fim da Classe ExportarSectraWindow

End Namespace ' Fim do Namespace CitrotecAddin