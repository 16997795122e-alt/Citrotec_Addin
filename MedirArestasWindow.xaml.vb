' Arquivo: MedirArestasWindow.xaml.vb (CORRIGIDO - Versão 2)
Imports Inventor
Imports System.Windows
Imports System.Text
Imports System.Runtime.InteropServices ' Para ReleaseComObject

Public Class MedirArestasWindow
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
        Log("Iniciando 'Medir Arestas'...")

        _invApp = Globals.g_inventorApplication
        Dim oDoc As Document = _invApp.ActiveDocument

        ' Especifica o namespace completo do Inventor
        Dim oCommandMgr As Inventor.CommandManager = _invApp.CommandManager

        Dim totalLength As Double = 0
        Dim continueSelecting As Boolean = True
        Dim edgeCount As Integer = 0
        Dim curveEval As CurveEvaluator = Nothing
        Dim firstEdge As Edge = Nothing
        Dim selectedEdge As Edge = Nothing

        Try
            ' Validação do Documento (Peça ou Montagem)
            If Not (TypeOf oDoc Is PartDocument OrElse TypeOf oDoc Is AssemblyDocument) Then
                Log("ERRO: Este comando só funciona em Peças (.ipt) ou Montagens (.iam).")
                FinalizeLog()
                Return
            End If

            ' Lógica adaptada de "Medir Arestas.txt"

            ' Esconde a janela para permitir a seleção no Inventor
            Me.Hide()

            ' Primeira seleção
            Try
                ' Agora .Pick será encontrado
                firstEdge = oCommandMgr.Pick(SelectionFilterEnum.kPartEdgeFilter, "Selecione uma aresta (ESC para finalizar)")
            Catch exPick As Exception
                Log("Seleção cancelada pelo usuário.")
                Me.Show() ' Mostra a janela novamente
                FinalizeLog()
                Return
            End Try

            ' Se nada for selecionado na primeira tentativa, cancela
            If firstEdge Is Nothing Then
                Log("Nenhuma aresta selecionada. Operação cancelada.")
                Me.Show()
                FinalizeLog()
                Exit Sub
            End If

            ' Calcula o comprimento da primeira aresta
            curveEval = firstEdge.Evaluator
            Dim minParam As Double
            Dim maxParam As Double
            curveEval.GetParamExtents(minParam, maxParam)
            Dim edgeLength As Double
            curveEval.GetLengthAtParam(minParam, maxParam, edgeLength)
            totalLength = totalLength + edgeLength
            edgeCount += 1
            Log($"Aresta {edgeCount} adicionada. Comprimento: {FormatLength(edgeLength)} mm")

            ' Loop para seleções adicionais
            While continueSelecting
                Try
                    ' Agora .Pick será encontrado
                    selectedEdge = oCommandMgr.Pick(SelectionFilterEnum.kPartEdgeFilter, $"Total: {FormatLength(totalLength * 10)} mm. Selecione outra aresta (ESC para finalizar)")
                Catch exLoopPick As Exception
                    ' Usuário pressionou ESC
                    selectedEdge = Nothing
                End Try

                If selectedEdge Is Nothing Then
                    continueSelecting = False
                Else
                    curveEval = selectedEdge.Evaluator
                    curveEval.GetParamExtents(minParam, maxParam)
                    curveEval.GetLengthAtParam(minParam, maxParam, edgeLength)

                    totalLength = totalLength + edgeLength
                    edgeCount += 1
                    Log($"Aresta {edgeCount} adicionada. Comprimento: {FormatLength(edgeLength)} mm")
                End If
                ReleaseComObject(selectedEdge)
            End While

            ' Converte para milímetros (totalLength está em cm)
            Dim totalLengthMM As Double = totalLength * 10

            ' Arredonda o valor
            Dim roundedLength As Integer = CInt(Math.Round(totalLengthMM))
            Dim resultText As String = roundedLength.ToString()

            ' Copia para o clipboard
            Clipboard.SetText(resultText)
            Log("-------------------------------------")
            Log($"Soma total ({edgeCount} arestas): {resultText} mm")
            Log("O valor foi copiado para o clipboard.")

            ' Mostra o resultado final em uma MessageBox
            Dim msgBox = New MessageBoxWindow($"Soma total: {resultText} mm{vbCrLf}{vbCrLf}O valor foi copiado para o clipboard.", "Medição Concluída", ModernMessageBoxButtons.OK)
            msgBox.Owner = Me
            Me.Show() ' Mostra a janela principal novamente
            msgBox.ShowDialog() ' Mostra o resultado

        Catch ex As Exception
            Log("======================================")
            Log("ERRO GERAL: " & ex.Message)
            Log(ex.StackTrace)
            Log("======================================")
            Me.Show() ' Garante que a janela reapareça em caso de erro
        Finally
            ReleaseComObject(firstEdge)
            ReleaseComObject(curveEval)
            FinalizeLog()
        End Try
    End Sub

    Private Function FormatLength(lengthCm As Double) As String
        Return (lengthCm * 10).ToString("N2") ' Formata com 2 casas decimais
    End Function

    Private Sub Log(message As String)
        _logBuilder.AppendLine($"{DateTime.Now:HH:mm:ss} - {message}")
        Me.Dispatcher.Invoke(Sub()
                                 If LogTextBox IsNot Nothing Then
                                     LogTextBox.Text = _logBuilder.ToString()
                                 End If
                             End Sub)
    End Sub

    Private Sub FinalizeLog()
        Me.Dispatcher.Invoke(Sub()
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
            ' Ignora erros na liberação, mas garante que a variável seja zerada
        Finally
            obj = Nothing
        End Try
    End Sub

End Class