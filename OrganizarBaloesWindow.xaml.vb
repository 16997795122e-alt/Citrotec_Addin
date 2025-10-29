' Arquivo: OrganizarBaloesWindow.xaml.vb
Imports Inventor
Imports System.Windows

Public Class OrganizarBaloesWindow

    Private Sub Border_MouseLeftButtonDown(sender As Object, e As Input.MouseButtonEventArgs)
        Me.DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub btnExecute_Click(sender As Object, e As RoutedEventArgs)
        LogTextBox.Clear()
        Log("Iniciando execução do comando 'Organizar Balões'...")

        Try
            Dim oDrawDoc As DrawingDocument = TryCast(g_inventorApplication.ActiveDocument, DrawingDocument)

            If oDrawDoc Is Nothing OrElse oDrawDoc.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then
                Log("ERRO: Nenhum documento de desenho (.idw) está ativo.")
                Return
            End If

            Log("Documento de desenho ativo: " & System.IO.Path.GetFileName(oDrawDoc.FullFileName))

            Dim oSheet As Sheet = oDrawDoc.ActiveSheet
            Log("Folha ativa: " & oSheet.Name)

            Dim balloons As Balloons = oSheet.Balloons
            Log($"Encontrados {balloons.Count} balões na folha.")

            Dim baloesOrdenados As Integer = 0
            For Each bal As Balloon In balloons
                If bal.BalloonValueSets.Count > 1 Then
                    ' Limpa seleção
                    oDrawDoc.SelectSet.Clear()

                    ' Seleciona o balloon
                    oDrawDoc.SelectSet.Select(bal)
                    Log($"Balão {bal.Style.Name} selecionado para ordenação.")

                    ' Executa o comando de ordenação
                    Dim sortCmd As ControlDefinition = g_inventorApplication.CommandManager.ControlDefinitions.Item("DLxBalloonSymSortCmd")
                    sortCmd.Execute()
                    Log("Comando de ordenação executado.")
                    baloesOrdenados += 1
                End If
            Next

            oDrawDoc.SelectSet.Clear()
            Log("Seleção limpa.")
            Log("-------------------------------------------------")
            Log($"Execução finalizada. {baloesOrdenados} balões foram organizados com sucesso.")

        Catch ex As Exception
            Log("======================================")
            Log("ERRO INESPERADO: " & ex.Message)
            Log(ex.StackTrace)
            Log("======================================")
        End Try
    End Sub

    Private Sub Log(message As String)
        LogTextBox.AppendText($"{DateTime.Now:HH:mm:ss} - {message}{vbCrLf}")
        LogTextBox.ScrollToEnd()
    End Sub
End Class