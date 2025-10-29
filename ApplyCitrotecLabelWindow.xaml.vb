' Arquivo: ApplyCitrotecLabelWindow.xaml.vb (VERSÃO CORRIGIDA FINAL - Compatível e Estável)
Imports Inventor
Imports System.Windows
Imports System.Text
Imports System.Linq
Imports System.Collections.Generic
Imports System.Runtime.InteropServices

Public Class ApplyCitrotecLabelWindow
    Private _logBuilder As New StringBuilder()
    Private _invApp As Inventor.Application

    Public Sub New()
        InitializeComponent()
        _invApp = Globals.g_inventorApplication
    End Sub

    ' ====================== EVENTOS DE UI ======================
    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    ' ====================== EXECUTAR OPERAÇÃO ======================
    Private Sub btnExecute_Click(sender As Object, e As RoutedEventArgs)
        _logBuilder.Clear()
        LogTextBox.Clear()
        Log("Iniciando aplicação de legenda Citrotec...")

        Dim oDoc As DrawingDocument = Nothing
        Dim oTransaction As Transaction = Nothing
        Dim selectedViews As New List(Of DrawingView)
        Dim labelTypeChoice As String = ""

        Try
            oDoc = TryCast(_invApp.ActiveDocument, DrawingDocument)
            If oDoc Is Nothing OrElse oDoc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                Log("ERRO: O documento ativo não é um desenho (.idw).")
                FinalizeLog()
                Return
            End If
            Log($"Documento de desenho ativo: {oDoc.DisplayName}")

            ' Verifica vistas selecionadas
            Dim oSSet As SelectSet = oDoc.SelectSet
            For i = 1 To oSSet.Count
                Dim oView As DrawingView = TryCast(oSSet.Item(i), DrawingView)
                If oView IsNot Nothing Then selectedViews.Add(oView)
            Next
            oSSet.Clear()

            If selectedViews.Count = 0 Then
                Log("Nenhuma vista pré-selecionada. Selecione manualmente...")
                Me.Hide()
                Do
                    Dim pickObj As Object = Nothing
                    Try
                        pickObj = _invApp.CommandManager.Pick(SelectionFilterEnum.kDrawingViewFilter, "Selecione vistas (ESC para sair)")
                        If pickObj Is Nothing Then Exit Do
                        Dim v As DrawingView = TryCast(pickObj, DrawingView)
                        If v IsNot Nothing AndAlso Not selectedViews.Contains(v) Then
                            selectedViews.Add(v)
                            Log($" -> Vista selecionada: {v.Name}")
                        End If
                    Catch ex As COMException When ex.ErrorCode = &H80004005
                        Exit Do
                    Catch ex As Exception
                        Log($"AVISO: Erro durante seleção: {ex.Message}")
                        Exit Do
                    Finally
                        ReleaseComObject(pickObj)
                    End Try
                Loop
                Me.Show()
                Me.Activate()
            End If

            If selectedViews.Count = 0 Then
                Log("Nenhuma vista foi selecionada. Cancelado.")
                FinalizeLog()
                Return
            End If

            ' Tipo de legenda
            If LabelTypeComboBox.SelectedItem Is Nothing Then
                Log("ERRO: Nenhum tipo de legenda selecionado.")
                FinalizeLog()
                Return
            End If
            labelTypeChoice = LabelTypeComboBox.SelectedItem.ToString()
            Log($"Tipo de legenda selecionado: '{labelTypeChoice}'")

            ' Transação
            oTransaction = _invApp.TransactionManager.StartTransaction(oDoc, "Aplicar Legenda Citrotec")
            _invApp.SilentOperation = True
            Log("Transação iniciada.")

            Dim successCount As Integer = 0
            Dim failCount As Integer = 0

            For Each v In selectedViews
                Try
                    Log($"-- Processando vista: '{v.Name}' --")

                    Dim refDesc = v.ReferencedDocumentDescriptor
                    Dim modelDoc As Document = If(refDesc IsNot Nothing, refDesc.ReferencedDocument, Nothing)
                    If modelDoc Is Nothing Then
                        Log("   -> Vista sem modelo referenciado. Pulando.")
                        failCount += 1
                        Continue For
                    End If

                    If ApplyLabelToView(v, modelDoc, labelTypeChoice) Then
                        successCount += 1
                        Log("   -> Legenda aplicada com sucesso.")
                    Else
                        failCount += 1
                        Log("   -> Falha ao aplicar legenda.")
                    End If

                Catch ex As Exception
                    Log($"   -> ERRO: {ex.Message}")
                    failCount += 1
                End Try
            Next

            oDoc.Update()
            oTransaction.End()
            _invApp.SilentOperation = False

            Log("Transação concluída.")
            Log("-------------------------------------")

            If failCount > 0 Then
                Log($"Finalizado com {successCount} sucessos e {failCount} falhas.")
                MessageBox.Show($"Legenda aplicada a {successCount} vistas, com {failCount} falhas.", "Operação Concluída com Erros", MessageBoxButton.OK, MessageBoxImage.Warning)
            Else
                MessageBox.Show($"Legenda '{labelTypeChoice}' aplicada com sucesso a todas as vistas.", "Operação Concluída", MessageBoxButton.OK, MessageBoxImage.Information)
            End If

        Catch ex As Exception
            Log($"ERRO GERAL: {ex.Message}")
            If oTransaction IsNot Nothing Then
                Try : oTransaction.Abort() : Catch : End Try
            End If
            _invApp.SilentOperation = False
        Finally
            FinalizeLog()
            ReleaseComObject(oTransaction)
        End Try
    End Sub

    ' ====================== LÓGICA PRINCIPAL ======================
    Private Function ApplyLabelToView(oView As DrawingView, modelDoc As Document, labelTypeChoice As String) As Boolean
        Try
            ' Assegura PropertySet
            Dim userSet As PropertySet = GetOrAddPropertySet(modelDoc, "User Defined Properties", "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
            If userSet Is Nothing Then Return False

            Dim propListaItem = GetOrAddProperty(userSet, "ListaItem", "[SEM ITEM]")
            Dim propItemQty = GetOrAddProperty(userSet, "ItemQty", 0)
            Dim propItemQtyDiv2 = GetOrAddProperty(userSet, "ItemQtyDiv2", 0)

            Dim qtyValue As Double = 0
            Try : qtyValue = CDbl(propItemQty.Value) : Catch : qtyValue = 0 : End Try
            Dim qtyDiv2 As Double = Math.Truncate(qtyValue / 2)
            propItemQtyDiv2.Value = qtyDiv2

            ' IDs fixos conforme padrão iLogic
            Dim propertyID_ListaItem As String = "26"
            Dim propertyID_ItemQty As String = "27"
            Dim propertyID_ItemQtyDiv2 As String = propItemQtyDiv2.PropId.ToString()

            ' Formatos válidos (baseados no código funcional)
            Dim oPartNumber As String = "<StyleOverride FontSize='0,4' Underline='True'>ITEM </StyleOverride>"
            Dim oListItem As String = $"<StyleOverride FontSize='0,4' Underline='True'><Property Document='model' PropertySet='User Defined Properties' Property='ListaItem' FormatID='{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}' PropertyID='{propertyID_ListaItem}'>ListaItem</Property></StyleOverride>"
            Dim oStringScale As String = "<Br/><StyleOverride FontSize='0,25'>( <DrawingViewScale/> )</StyleOverride>"
            Dim oQty As String = $"<Br/><StyleOverride FontSize='0,25'>Fabricar: <Property Document='model' PropertySet='User Defined Properties' Property='ItemQty' FormatID='{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}' PropertyID='{propertyID_ItemQty}'>ItemQty</Property>x</StyleOverride>"
            Dim oQtyDiv2 As String = $"<Br/><StyleOverride FontSize='0,25'>Esquerda: <Property Document='model' PropertySet='User Defined Properties' Property='ItemQtyDiv2' FormatID='{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}' PropertyID='{propertyID_ItemQtyDiv2}'>ItemQtyDiv2</Property>x</StyleOverride>" &
                                     $"<Br/><StyleOverride FontSize='0,25'>Direita: <Property Document='model' PropertySet='User Defined Properties' Property='ItemQtyDiv2' FormatID='{{D5CDD505-2E9C-101B-9397-08002B2CF9AE}}' PropertyID='{propertyID_ItemQtyDiv2}'>ItemQtyDiv2</Property>x</StyleOverride>"
            Dim Planificacao As String = "<Br/><StyleOverride FontSize='0,25'>PLANIFICAÇÃO</StyleOverride>"

            ' Detecta se é vista de planificação
            Dim isFlat As Boolean = False
            Try : isFlat = oView.IsFlatPatternView : Catch : End Try

            Dim legendaTexto As String = ""

            Select Case labelTypeChoice
                Case "Com quantidade"
                    legendaTexto = If(isFlat, oPartNumber & oListItem & Planificacao & oQty & oStringScale,
                                              oPartNumber & oListItem & oQty & oStringScale)
                Case "Sem quantidade"
                    legendaTexto = If(isFlat, oPartNumber & oListItem & Planificacao & oStringScale,
                                              oPartNumber & oListItem & oStringScale)
                Case "Apenas (ITEM)"
                    legendaTexto = oPartNumber & oListItem & oStringScale
                Case "Esquerda/Direita"
                    If qtyValue Mod 2 <> 0 Or qtyValue = 0 Then
                        MessageBox.Show($"A quantidade ({qtyValue}) é ímpar ou inválida. Corrija para aplicar Esquerda/Direita.", "Quantidade Inválida", MessageBoxButton.OK, MessageBoxImage.Warning)
                        Return False
                    End If
                    If isFlat Then
                        MessageBox.Show("A opção Esquerda/Direita não é aplicável a vistas planificadas.", "Inválido", MessageBoxButton.OK, MessageBoxImage.Warning)
                        Return False
                    End If
                    legendaTexto = oPartNumber & oListItem & oQtyDiv2 & oStringScale
                Case Else
                    Return False
            End Select

            oView.ShowLabel = True
            oView.Label.FormattedText = legendaTexto
            Return True

        Catch ex As Exception
            Log($"   -> ERRO CRÍTICO em ApplyLabelToView: {ex.Message}")
            Return False
        End Try
    End Function

    ' ====================== FUNÇÕES AUXILIARES ======================
    Private Function GetOrAddPropertySet(doc As Document, setName As String, internalName As String) As PropertySet
        Try
            Return doc.PropertySets.Item(setName)
        Catch
            Return doc.PropertySets.Add(setName, internalName)
        End Try
    End Function

    Private Function GetOrAddProperty(propSet As PropertySet, propName As String, defaultValue As Object) As Inventor.Property
        Try
            Return propSet.Item(propName)
        Catch
            Return propSet.Add(defaultValue, propName)
        End Try
    End Function

    Private Sub Log(msg As String)
        _logBuilder.AppendLine($"{DateTime.Now:HH:mm:ss} - {msg}")
        Me.Dispatcher.InvokeAsync(Sub()
                                      LogTextBox.Text = _logBuilder.ToString()
                                      LogTextBox.ScrollToEnd()
                                  End Sub)
    End Sub

    Private Sub FinalizeLog()
        Me.Dispatcher.InvokeAsync(Sub()
                                      LogTextBox.Text = _logBuilder.ToString()
                                      LogTextBox.ScrollToEnd()
                                  End Sub)
    End Sub

    Private Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                While Marshal.ReleaseComObject(obj) > 0
                End While
            End If
        Catch
        Finally
            obj = Nothing
        End Try
    End Sub
End Class
