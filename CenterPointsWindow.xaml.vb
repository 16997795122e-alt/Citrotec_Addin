' Arquivo: CenterPointsWindow.xaml.vb
Imports Inventor
Imports System.Windows
Imports System.Text ' Para o Log

Public Class CenterPointsWindow
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
        Log("Iniciando 'Criar Pontos Centrais'...")

        Try
            _invApp = Globals.g_inventorApplication
            Dim oDoc As Document = _invApp.ActiveDocument

            ' 1. Validação do Documento (Apenas Peças, conforme sua solicitação)
            If Not (TypeOf oDoc Is PartDocument) Then
                Log("ERRO: Este comando deve ser executado a partir de um documento de Peça (IPT).")
                FinalizeLog()
                Return
            End If

            ' 2. Validação do Esboço
            If Not TypeOf _invApp.ActiveEditObject Is Sketch Then
                Log("ERRO: O objeto ativo não é um esboço (Sketch). Por favor, edite um esboço para executar este comando.")
                FinalizeLog()
                Return
            End If

            Dim oSketch As Sketch = _invApp.ActiveEditObject
            Log("Esboço ativo: " & oSketch.Name)

            ' 3. Validação de Círculos
            If oSketch.SketchCircles.Count = 0 Then
                Log("Nenhum círculo encontrado no esboço ativo.")
                FinalizeLog()
                Return
            End If
            Log($"Encontrados {oSketch.SketchCircles.Count} círculos.")

            ' 4. Otimização e Transação
            Dim oTransaction As Transaction
            oTransaction = _invApp.TransactionManager.StartTransaction(oDoc, "iLogic - Criar Pontos nos Centros")
            _invApp.ScreenUpdating = False
            Dim pointsCreated As Integer = 0

            Try
                ' Loop principal
                For Each oCircle As SketchCircle In oSketch.SketchCircles
                    ' Adiciona um novo ponto de esboço (SketchPoint) na posição do centro do círculo
                    Dim oNewPoint As SketchPoint = oSketch.SketchPoints.Add(oCircle.Geometry.Center)

                    ' Obtém a referência ao ponto de esboço que define o centro do círculo
                    Dim oCircleCenterPoint As SketchPoint = oCircle.CenterSketchPoint

                    ' Adiciona uma restrição de coincidência
                    oSketch.GeometricConstraints.AddCoincident(oNewPoint, oCircleCenterPoint)
                    pointsCreated += 1
                Next

                ' Finaliza a transação
                oTransaction.End()
                Log($"Sucesso: {pointsCreated} pontos centrais foram criados.")
                Log("-------------------------------------")
                Log("Execução finalizada.")

            Catch ex As Exception
                ' Em caso de erro, desfaz
                If Not oTransaction Is Nothing Then
                    oTransaction.Abort()
                End If
                Log("======================================")
                Log("ERRO INESPERADO: " & ex.Message)
                Log("======================================")
            Finally
                ' Sempre reabilita a atualização da tela
                _invApp.ScreenUpdating = True
            End Try

        Catch ex As Exception
            Log("======================================")
            Log("ERRO GERAL: " & ex.Message)
            Log(ex.StackTrace)
            Log("======================================")
        Finally
            FinalizeLog()
        End Try
    End Sub

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