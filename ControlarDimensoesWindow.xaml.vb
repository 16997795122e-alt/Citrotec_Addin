' Arquivo: ControlarDimensoesWindow.xaml.vb (Unificado - Versão 3 - Correção Completa)

Imports Inventor
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports System.Text
Imports System.IO
Imports System.Diagnostics
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices ' Para Marshal.ReleaseComObject
Imports Microsoft.VisualBasic ' Para InputBox

Namespace CitrotecAddin

    Partial Public Class ControlarDimensoesWindow
        Inherits System.Windows.Window

        Private _logBuilder As New StringBuilder()
        Private _invApp As Inventor.Application
        Private _doc As Document ' Documento ativo (Peça ou Montagem)
        Private ReadOnly RULE_NAME As String = "Comprimento Peça"
        Private _isAssembly As Boolean = False

        Public Sub New()
            InitializeComponent()
            _invApp = Globals.g_inventorApplication ' Pega a aplicação global
            _doc = _invApp.ActiveDocument ' Pega o documento ativo na inicialização
        End Sub

        Private Sub Window_Loaded(sender As Object, e As RoutedEventArgs)
            UpdateUIForDocumentType() ' Configura a UI baseado no documento ativo
        End Sub

        Private Sub UpdateUIForDocumentType()
            _doc = _invApp.ActiveDocument ' Garante pegar o doc atual ao carregar a janela
            _isAssembly = False ' Reseta a flag

            If _doc Is Nothing Then
                DescriptionTextBlock.Text = "Nenhum documento ativo."
                SetOptionsAvailability(partEnabled:=False, assemblyEnabled:=False)
                btnExecute.IsEnabled = False
                btnAbrirVerificarRegras.IsEnabled = False
                Return
            End If

            If TypeOf _doc Is PartDocument Then
                DescriptionTextBlock.Text = "Selecione o tipo de regra 'Comprimento Peça' a ser criada/atualizada na PEÇA atual."
                SetOptionsAvailability(partEnabled:=True, assemblyEnabled:=False)
                If Not rbParam1.IsChecked.GetValueOrDefault() AndAlso Not rbParam2.IsChecked.GetValueOrDefault() AndAlso Not rbComprimentoPadrao.IsChecked.GetValueOrDefault() Then
                    rbSheetMetal.IsChecked = True
                End If
                btnExecute.IsEnabled = True
                btnAbrirVerificarRegras.IsEnabled = False
            ElseIf TypeOf _doc Is AssemblyDocument Then
                _isAssembly = True
                DescriptionTextBlock.Text = "Selecione a regra a ser aplicada nas peças SELECIONADAS na MONTAGEM."
                SetOptionsAvailability(partEnabled:=True, assemblyEnabled:=True)
                If Not rbSheetMetal.IsChecked.GetValueOrDefault() AndAlso Not rbParam1.IsChecked.GetValueOrDefault() AndAlso Not rbParam2.IsChecked.GetValueOrDefault() Then
                    rbComprimentoPadrao.IsChecked = True
                End If
                btnExecute.IsEnabled = True
                btnAbrirVerificarRegras.IsEnabled = True
            Else
                DescriptionTextBlock.Text = "Este comando funciona apenas com Peças (.ipt) ou Montagens (.iam)."
                SetOptionsAvailability(partEnabled:=False, assemblyEnabled:=False)
                btnExecute.IsEnabled = False
                btnAbrirVerificarRegras.IsEnabled = False
            End If
            RadioButton_CheckedChanged(Nothing, Nothing)
        End Sub

        Private Sub SetOptionsAvailability(partEnabled As Boolean, assemblyEnabled As Boolean)
            rbSheetMetal.IsEnabled = partEnabled Or assemblyEnabled
            rbParam1.IsEnabled = partEnabled Or assemblyEnabled
            rbParam2.IsEnabled = partEnabled Or assemblyEnabled
            ' ***** ALTERAÇÃO AQUI *****
            ' Habilita a opção para Peça (partEnabled) ou Montagem (assemblyEnabled)
            rbComprimentoPadrao.IsEnabled = partEnabled Or assemblyEnabled
            ' ***** FIM DA ALTERAÇÃO *****
        End Sub

        Private Sub Border_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs)
            Me.DragMove()
        End Sub

        Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs)
            Me.Close()
        End Sub

        Private Sub btnAbrirVerificarRegras_Click(sender As Object, e As RoutedEventArgs)
            WindowManager.ShowWindow(Of VerificarRegrasWindow)()
        End Sub

        Private Sub RadioButton_CheckedChanged(sender As Object, e As RoutedEventArgs)
            If Param1Panel Is Nothing OrElse Param2Panel Is Nothing OrElse rbParam1 Is Nothing OrElse rbParam2 Is Nothing Then Return
            Param1Panel.Visibility = If(rbParam1.IsChecked.GetValueOrDefault(), Visibility.Visible, Visibility.Collapsed)
            Param2Panel.Visibility = If(rbParam2.IsChecked.GetValueOrDefault(), Visibility.Visible, Visibility.Collapsed)
        End Sub

        Private Sub btnExecute_Click(sender As Object, e As RoutedEventArgs)
            _logBuilder.Clear()
            LogTextBox.Clear()
            _doc = _invApp.ActiveDocument

            If _doc Is Nothing Then
                Log("ERRO: Nenhum documento ativo.")
                FinalizeLog()
                Return
            End If

            Try
                If TypeOf _doc Is PartDocument Then
                    ExecuteRuleOnActivePart(CType(_doc, PartDocument))
                ElseIf TypeOf _doc Is AssemblyDocument Then
                    ExecuteRuleOnSelectedAssemblyComponents(CType(_doc, AssemblyDocument))
                Else
                    Log("ERRO: Tipo de documento não suportado.")
                End If
            Catch ex As Exception
                Log("======================================")
                Log("ERRO INESPERADO: " & ex.Message)
                Log(ex.StackTrace)
                Log("======================================")
            Finally
                FinalizeLog()
            End Try
        End Sub

#Region "Lógica Específica para PEÇA ATIVA"
        Private Sub ExecuteRuleOnActivePart(partDoc As PartDocument)
            Log($"Iniciando processo para a Peça: {partDoc.DisplayName}")
            Dim ruleText As String = ""
            Dim paramFx1 As String = ""
            Dim paramFx2 As String = ""

            If rbParam1.IsChecked.GetValueOrDefault() Then
                paramFx1 = txtParam1.Text.Trim()
                If String.IsNullOrWhiteSpace(paramFx1) Then Throw New Exception("O nome do parâmetro para '1 Parâmetro' não pode estar vazio.")
                Log($"Opção selecionada: 1 Parâmetro ('{paramFx1}')")
                ruleText = GenerateRuleText_1Param(paramFx1)
            ElseIf rbParam2.IsChecked.GetValueOrDefault() Then
                paramFx1 = txtParam1_2.Text.Trim()
                paramFx2 = txtParam2_2.Text.Trim()
                If String.IsNullOrWhiteSpace(paramFx1) OrElse String.IsNullOrWhiteSpace(paramFx2) Then Throw New Exception("Os nomes dos parâmetros para '2 Parâmetros' não podem estar vazios.")
                Log($"Opção selecionada: 2 Parâmetros ('{paramFx1}', '{paramFx2}')")
                ruleText = GenerateRuleText_2Param(paramFx1, paramFx2)

                ' ***** TRECHO ADICIONADO INÍCIO *****
            ElseIf rbComprimentoPadrao.IsChecked.GetValueOrDefault() Then
                Log("Opção selecionada: Comprimento Padrão (G_L/L)")
                Dim oParams As Parameters = partDoc.ComponentDefinition.Parameters
                Dim hasGL As Boolean = oParams.UserParameters.Cast(Of UserParameter)().Any(Function(p) p.Name = "G_L")
                Dim hasL As Boolean = oParams.UserParameters.Cast(Of UserParameter)().Any(Function(p) p.Name = "L")

                If hasGL AndAlso hasL Then
                    Log("   -> Ambos 'G_L' e 'L' encontrados.")

                    ' Usa a janela de seleção customizada
                    Dim prompt As String = $"Peça {partDoc.DisplayName}: Ambos 'G_L' e 'L' encontrados." & vbCrLf & "Qual parâmetro usar?"
                    Dim title As String = "Escolha Parâmetro"
                    Dim options As New List(Of String) From {"G_L", "L"}

                    Dim selectWindow As New SelectModelStateWindow(title, prompt, options, "G_L")
                    selectWindow.Owner = Me

                    Dim userChoice As String = ""
                    If selectWindow.ShowDialog().GetValueOrDefault(False) Then
                        userChoice = selectWindow.SelectedModelStateName
                    End If

                    If userChoice.ToUpper() = "G_L" OrElse userChoice.ToUpper() = "L" Then
                        paramFx1 = userChoice.ToUpper()
                    End If
                    Log($"   -> Usuário escolheu/resultado: {If(String.IsNullOrEmpty(paramFx1), "Inválido", paramFx1)}")

                ElseIf hasGL Then
                    paramFx1 = "G_L"
                    Log("   -> Parâmetro 'G_L' encontrado.")
                ElseIf hasL Then
                    paramFx1 = "L"
                    Log("   -> Parâmetro 'L' encontrado.")
                End If

                If String.IsNullOrWhiteSpace(paramFx1) Then
                    Throw New Exception("Nenhum parâmetro (G_L ou L) encontrado ou escolhido para Comprimento Padrão.")
                End If

                ruleText = GenerateRuleText_ComprimentoPadrao(paramFx1)
                Log($"   -> Aplicando regra Comprimento Padrão ({paramFx1})...")
                ' Configura unidades (necessário para a regra G_L/L)
                Log("   -> Aplicando unidades (MM e KG)...")
                partDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits
                partDoc.UnitsOfMeasure.MassUnits = UnitsTypeEnum.kKilogramMassUnits
                ' ***** TRECHO ADICIONADO FIM *****

            Else ' Opção Padrão: Sheet Metal
                Log("Opção selecionada: Sheet Metal")
                Dim smDef As SheetMetalComponentDefinition = TryCast(partDoc.ComponentDefinition, SheetMetalComponentDefinition)
                If smDef Is Nothing Then Throw New Exception("A opção 'Sheet Metal' só é válida para peças de chapa metálica (.ipt).")
                ruleText = GenerateRuleText_SheetMetal()
            End If

            Log("Tentando adicionar/substituir a regra iLogic...")
            AddOrReplaceILogicRule(partDoc, RULE_NAME, ruleText)

            Log("Configurando o gatilho 'Before Save Document'...")
            SetupEventTrigger(partDoc, RULE_NAME)

            Log("-------------------------------------")
            Log("Regra e gatilho configurados com sucesso na peça ativa!")
        End Sub
#End Region

#Region "Lógica Específica para MONTAGEM ATIVA (Seleção)"
        Private Sub ExecuteRuleOnSelectedAssemblyComponents(oAsmDoc As AssemblyDocument)
            Log($"Iniciando processo para componentes selecionados na Montagem: {oAsmDoc.DisplayName}")
            Dim processedParts As New List(Of String)
            Dim selectedItems As ObjectCollection = Nothing
            Dim iLogic As Object = Nothing
            Dim iLogicAddIn As ApplicationAddIn = Nothing
            Dim partDoc As PartDocument = Nothing

            Dim ruleType As String = ""
            Dim ruleDescription As String = ""
            Dim param1Input As String = ""
            Dim param2Input As String = ""

            If rbSheetMetal.IsChecked.GetValueOrDefault() Then
                ruleType = "SheetMetal"
                ruleDescription = "Sheet Metal (Automático)"
            ElseIf rbParam1.IsChecked.GetValueOrDefault() Then
                ruleType = "Param1"
                param1Input = txtParam1.Text.Trim()
                ruleDescription = $"1 Parâmetro ({param1Input})"
                If String.IsNullOrWhiteSpace(param1Input) Then Throw New Exception("O nome do parâmetro para '1 Parâmetro' não pode estar vazio.")
            ElseIf rbParam2.IsChecked.GetValueOrDefault() Then
                ruleType = "Param2"
                param1Input = txtParam1_2.Text.Trim()
                param2Input = txtParam2_2.Text.Trim()
                ruleDescription = $"2 Parâmetros ({param1Input}, {param2Input})"
                If String.IsNullOrWhiteSpace(param1Input) OrElse String.IsNullOrWhiteSpace(param2Input) Then Throw New Exception("Os nomes dos parâmetros para '2 Parâmetros' não podem estar vazios.")
            ElseIf rbComprimentoPadrao.IsChecked.GetValueOrDefault() Then
                ruleType = "ComprimentoPadrao"
                ruleDescription = "Comprimento Padrão (G_L/L)"
            Else
                Throw New Exception("Nenhum modo de operação válido selecionado.")
            End If
            Log($"Modo selecionado: {ruleDescription}")

            Try
                selectedItems = _invApp.TransientObjects.CreateObjectCollection()
                _invApp.SilentOperation = True

                Dim selectSet As SelectSet = oAsmDoc.SelectSet

                If selectSet.Count > 0 Then
                    Log($"Encontrados {selectSet.Count} itens pré-selecionados.")
                    For Each item In selectSet
                        If TypeOf item Is ComponentOccurrence Then
                            selectedItems.Add(item)
                        End If
                    Next
                Else
                    Log("Nenhum item pré-selecionado. Solicitando seleção manual (pressione ESC para parar)...")
                    Dim keepSelecting As Boolean = True
                    While keepSelecting
                        Dim oSelection As Object = _invApp.CommandManager.Pick(SelectionFilterEnum.kAssemblyLeafOccurrenceFilter, "Selecione as PEÇAS alvo e pressione ESC para continuar")
                        If oSelection IsNot Nothing Then
                            selectedItems.Add(oSelection)
                            Log($"Item selecionado: {oSelection.Name}")
                        Else
                            keepSelecting = False
                            Log("Seleção manual concluída.")
                        End If
                    End While
                End If

                If selectedItems.Count = 0 Then
                    Log("Nenhum item foi selecionado. Operação cancelada.")
                    Throw New OperationCanceledException("Nenhum item selecionado.")
                End If

                Log($"Processando {selectedItems.Count} itens selecionados...")

                iLogicAddIn = _invApp.ApplicationAddIns.ItemById("{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}")
                If iLogicAddIn Is Nothing Then Throw New Exception("Add-in iLogic não encontrado.")
                If Not iLogicAddIn.Activated Then iLogicAddIn.Activate()
                iLogic = iLogicAddIn.Automation

                For Each selectedItem In selectedItems
                    partDoc = Nothing
                    Dim partName As String = ""
                    Dim wasOpened As Boolean = False

                    Try
                        If TypeOf selectedItem Is ComponentOccurrence Then
                            Dim occurrence As ComponentOccurrence = CType(selectedItem, ComponentOccurrence)
                            Dim itemDef As ComponentDefinition = occurrence.Definition

                            If Not (TypeOf itemDef Is PartComponentDefinition) Then
                                Log($"AVISO: Item '{occurrence.Name}' não é uma peça (é {itemDef.Type}). Pulando.")
                                Continue For
                            End If

                            Dim partDef As PartComponentDefinition = CType(itemDef, PartComponentDefinition)
                            Dim partDocument As Document = partDef.Document

                            If partDocument Is Nothing Then
                                Log($"AVISO: Não foi possível obter o documento da peça '{occurrence.Name}'. Pulando.")
                                Continue For
                            End If

                            partName = occurrence.Name
                            Log($"--- Processando: {partName} ({System.IO.Path.GetFileName(partDocument.FullFileName)}) ---")

                            Dim isAlreadyOpen As Boolean = False
                            For Each openDoc As Document In _invApp.Documents
                                If openDoc.FullFileName = partDocument.FullFileName Then
                                    partDoc = TryCast(openDoc, PartDocument)
                                    isAlreadyOpen = True
                                    Exit For
                                End If
                            Next

                            If Not isAlreadyOpen Then
                                partDoc = TryCast(_invApp.Documents.Open(partDocument.FullFileName, True), PartDocument)
                                wasOpened = True
                            End If

                            If partDoc Is Nothing Then
                                Log($"   -> ERRO: Falha ao abrir o documento {partName}. Pulando.")
                                Continue For
                            End If

                            If Not (TypeOf partDoc Is PartDocument) Then
                                Log($"   -> AVISO: Documento '{partName}' não é uma peça válida. Pulando.")
                                If wasOpened Then partDoc.Close(True)
                                Continue For
                            End If

                            Dim ruleText As String = ""
                            Dim paramFx1 As String = ""

                            Select Case ruleType
                                Case "SheetMetal"
                                    ruleText = GenerateRuleText_SheetMetal()
                                    Log("   -> Aplicando regra Sheet Metal...")

                                Case "Param1"
                                    ruleText = GenerateRuleText_1Param(param1Input)
                                    Log($"   -> Aplicando regra 1 Parâmetro ({param1Input})...")

                                Case "Param2"
                                    ruleText = GenerateRuleText_2Param(param1Input, param2Input)
                                    Log($"   -> Aplicando regra 2 Parâmetros ({param1Input}, {param2Input})...")

                                Case "ComprimentoPadrao"
                                    Dim oParams As Parameters = partDoc.ComponentDefinition.Parameters
                                    Dim hasGL As Boolean = oParams.UserParameters.Cast(Of UserParameter)().Any(Function(p) p.Name = "G_L")
                                    Dim hasL As Boolean = oParams.UserParameters.Cast(Of UserParameter)().Any(Function(p) p.Name = "L")

                                    If hasGL AndAlso hasL Then
                                        Log("   -> Ambos 'G_L' e 'L' encontrados.")

                                        ' ***** ALTERAÇÃO INICIADA *****
                                        ' Substitui a InputBoxWindow pela SelectModelStateWindow

                                        Dim prompt As String = $"Peça {partDoc.DisplayName}: Ambos 'G_L' e 'L' encontrados." & vbCrLf & "Qual parâmetro usar?"
                                        Dim title As String = "Escolha Parâmetro"
                                        Dim options As New List(Of String) From {"G_L", "L"}

                                        ' Usa o novo construtor da SelectModelStateWindow
                                        Dim selectWindow As New SelectModelStateWindow(title, prompt, options, "G_L")
                                        selectWindow.Owner = Me ' Define esta janela como "pai"

                                        Dim userChoice As String = ""
                                        ' Exibe a janela modal
                                        If selectWindow.ShowDialog().GetValueOrDefault(False) Then
                                            ' Pega o item selecionado no ComboBox [cite: 726]
                                            userChoice = selectWindow.SelectedModelStateName
                                        End If
                                        ' ***** ALTERAÇÃO FINALIZADA *****

                                        If userChoice.ToUpper() = "G_L" OrElse userChoice.ToUpper() = "L" Then
                                            paramFx1 = userChoice.ToUpper()
                                        Else
                                            paramFx1 = ""
                                        End If
                                        Log($"   -> Usuário escolheu/resultado: {If(String.IsNullOrEmpty(paramFx1), "Inválido", paramFx1)}")
                                    End If
                                    If String.IsNullOrWhiteSpace(paramFx1) Then
                                        Log("   -> AVISO: Nenhum parâmetro (G_L ou L) encontrado ou escolhido para Comprimento Padrão. Pulando.")
                                        If wasOpened Then partDoc.Close(True)
                                        Continue For
                                    End If
                                    ruleText = GenerateRuleText_ComprimentoPadrao(paramFx1)
                                    Log($"   -> Aplicando regra Comprimento Padrão ({paramFx1})...")
                                    Log("   -> Aplicando unidades (MM e KG)...")
                                    partDoc.UnitsOfMeasure.LengthUnits = UnitsTypeEnum.kMillimeterLengthUnits
                                    partDoc.UnitsOfMeasure.MassUnits = UnitsTypeEnum.kKilogramMassUnits

                                Case Else
                                    Throw New Exception("Tipo de regra desconhecido.")
                            End Select

                            AddOrReplaceILogicRule(partDoc, RULE_NAME, ruleText)

                            If ruleType <> "ComprimentoPadrao" Then
                                Log("   -> Configurando gatilho 'Before Save Document'...")
                                SetupEventTrigger(partDoc, RULE_NAME)
                            Else
                                Log("   -> Gatilho 'Before Save' configurado dentro da própria regra.")
                            End If

                            Try
                                partDoc.Save()
                            Catch exSave As Exception
                                Log($"   -> AVISO: Falha ao salvar {partName}: {exSave.Message}")
                            End Try

                            If wasOpened Then
                                Try
                                    partDoc.Close(True)
                                Catch exClose As Exception
                                    Log($"   -> AVISO: Falha ao fechar {partName}: {exClose.Message}")
                                End Try
                            End If
                            processedParts.Add(partDoc.DisplayName)
                            Log($"   -> Processamento de {partName} concluído.")

                        Else
                            Log("AVISO: Item selecionado não é uma ComponentOccurrence válida. Pulando.")
                            Continue For
                        End If

                    Catch exPart As Exception
                        Log($"   -> ERRO ao processar {partName}: {exPart.Message}")
                        If partDoc IsNot Nothing AndAlso wasOpened Then
                            Try
                                partDoc.Close(True)
                            Catch exClose As Exception
                                Log($"   -> AVISO: Falha ao fechar {partName}: {exClose.Message}")
                            End Try
                        End If
                    Finally
                        ReleaseComObject(partDoc)
                    End Try
                Next

            Finally
                _invApp.SilentOperation = False
                If processedParts.Count > 0 Then
                    Log("-------------------------------------")
                    Log($"Regra '{ruleDescription}' aplicada com sucesso para:")
                    For Each partName In processedParts
                        Log($"- {partName}")
                    Next
                Else
                    Log("Nenhuma peça foi processada com sucesso (verifique erros, seleção ou pré-condições da regra).")
                End If
                ReleaseComObject(selectedItems)
                ReleaseComObject(iLogic)
                ReleaseComObject(iLogicAddIn)
            End Try
        End Sub
#End Region

#Region "Geração de Texto da Regra (Universal)"
        Private Function GenerateRuleText_1Param(paramName As String) As String
            Return $"' Atualiza a DIMENSAO1 com arredondamento seguro{vbCrLf}" &
           $"DIMENSAO1 = Math.Round({paramName}, 3){vbCrLf}" &
           $"DIMENSAO1 = If(DIMENSAO1 - Math.Floor(DIMENSAO1) < 0.01, Math.Floor(DIMENSAO1), Math.Ceiling(DIMENSAO1)){vbCrLf}" &
           $"iProperties.Value(""Custom"", ""DIMENSAO1"") = DIMENSAO1{vbCrLf}" &
           $"iProperties.Value(""Custom"", ""DIMENSAO2"") = """"{vbCrLf}" &
           $"InventorVb.DocumentUpdate(){vbCrLf}" &
           $"Parameter.UpdateAfterChange = True{vbCrLf}" &
           $"iLogicVb.UpdateWhenDone = True"
        End Function

        Private Function GenerateRuleText_2Param(param1Name As String, param2Name As String) As String
            Return $"' Atualiza DIMENSAO1 e DIMENSAO2 com arredondamento seguro{vbCrLf}" &
           $"DIMENSAO1 = Math.Round({param1Name}, 3){vbCrLf}" &
           $"DIMENSAO1 = If(DIMENSAO1 - Math.Floor(DIMENSAO1) < 0.01, Math.Floor(DIMENSAO1), Math.Ceiling(DIMENSAO1)){vbCrLf}" &
           $"DIMENSAO2 = Math.Round({param2Name}, 3){vbCrLf}" &
           $"DIMENSAO2 = If(DIMENSAO2 - Math.Floor(DIMENSAO2) < 0.01, Math.Floor(DIMENSAO2), Math.Ceiling(DIMENSAO2)){vbCrLf}" &
           $"If DIMENSAO1 > DIMENSAO2 Then{vbCrLf}" &
           $"    iProperties.Value(""Custom"", ""DIMENSAO1"") = DIMENSAO2{vbCrLf}" &
           $"    iProperties.Value(""Custom"", ""DIMENSAO2"") = DIMENSAO1{vbCrLf}" &
           $"Else{vbCrLf}" &
           $"    iProperties.Value(""Custom"", ""DIMENSAO1"") = DIMENSAO1{vbCrLf}" &
           $"    iProperties.Value(""Custom"", ""DIMENSAO2"") = DIMENSAO2{vbCrLf}" &
           $"End If{vbCrLf}" &
           $"InventorVb.DocumentUpdate(){vbCrLf}" &
           $"Parameter.UpdateAfterChange = True{vbCrLf}" &
           $"iLogicVb.UpdateWhenDone = True"
        End Function

        Private Function GenerateRuleText_SheetMetal() As String
            Return $"' Atualiza com base na planificação (Sheet Metal) com arredondamento seguro{vbCrLf}" &
           $"largura = Math.Round(SheetMetal.FlatExtentsWidth, 3){vbCrLf}" &
           $"comprimento = Math.Round(SheetMetal.FlatExtentsLength, 3){vbCrLf}" &
           $"DIMENSAO1 = If(largura - Math.Floor(largura) < 0.01, Math.Floor(largura), Math.Ceiling(largura)){vbCrLf}" &
           $"DIMENSAO2 = If(comprimento - Math.Floor(comprimento) < 0.01, Math.Floor(comprimento), Math.Ceiling(comprimento)){vbCrLf}" &
           $"If DIMENSAO1 > DIMENSAO2 Then{vbCrLf}" &
           $"    iProperties.Value(""Custom"", ""DIMENSAO1"") = DIMENSAO2{vbCrLf}" &
           $"    iProperties.Value(""Custom"", ""DIMENSAO2"") = DIMENSAO1{vbCrLf}" &
           $"Else{vbCrLf}" &
           $"    iProperties.Value(""Custom"", ""DIMENSAO1"") = DIMENSAO1{vbCrLf}" &
           $"    iProperties.Value(""Custom"", ""DIMENSAO2"") = DIMENSAO2{vbCrLf}" &
           $"End If{vbCrLf}" &
           $"InventorVb.DocumentUpdate(){vbCrLf}" &
           $"Parameter.UpdateAfterChange = True{vbCrLf}" &
           $"iLogicVb.UpdateWhenDone = True"
        End Function

        Private Function GenerateRuleText_ComprimentoPadrao(paramName As String) As String
            Return "' Esse código atualiza DIMENSAO1, configura evento BeforeDocSave e altera unidades" & vbCrLf & vbCrLf &
                   "' Parte 1: Atualização do Comprimento" & vbCrLf &
                   $"Try{vbCrLf}" &
                   $"    DIMENSAO1 = Parameter(""{paramName}""){vbCrLf}" &
                   $"Catch{vbCrLf}" &
                   $"    MessageBox.Show(""Parâmetro '{paramName}' (G_L ou L) não encontrado."", ""Erro na Regra '{RULE_NAME}'"", MessageBoxButtons.OK, MessageBoxIcon.Warning){vbCrLf}" &
                   $"    Exit Sub{vbCrLf}" &
                   $"End Try{vbCrLf}" &
                   "DIMENSAO1 = Ceil(DIMENSAO1)" & vbCrLf &
                   "iProperties.Value(""Custom"", ""DIMENSAO1"") = DIMENSAO1" & vbCrLf & vbCrLf &
                   "' Parte 2: Configuração do evento BeforeDocSave (embutido na regra)" & vbCrLf &
                   "Dim oDoc As Inventor.Document = ThisDoc.Document" & vbCrLf &
                   "Dim oRuleName As String = """ & RULE_NAME & """" & vbCrLf &
                   "Dim oAuto As Object = iLogicVb.Automation" & vbCrLf &
                   "Dim oRule As Object" & vbCrLf &
                   "Try" & vbCrLf &
                   "    oRule = oAuto.GetRule(oDoc, oRuleName)" & vbCrLf &
                   "Catch" & vbCrLf &
                   "    Exit Sub" & vbCrLf &
                   "End Try" & vbCrLf &
                   "Dim oETPropSet As PropertySet" & vbCrLf &
                   "Try" & vbCrLf &
                   "    oETPropSet = oDoc.PropertySets.Item(""iLogicEventsRules"")" & vbCrLf &
                   "Catch" & vbCrLf &
                   "    Try" & vbCrLf &
                   "        oETPropSet = oDoc.PropertySets.Add(""iLogicEventsRules"", ""{2C540830-0723-455E-A8E2-891722EB4C3E}"")" & vbCrLf &
                   "    Catch ex As Exception" & vbCrLf &
                   "        Try : oETPropSet = oDoc.PropertySets.Item(""_iLogicEventsRules"") : Catch : Exit Sub : End Try" & vbCrLf &
                   "    End Try" & vbCrLf &
                   "End Try" & vbCrLf &
                   "If oETPropSet Is Nothing Then Exit Sub" & vbCrLf &
                   "Dim oProperty As Inventor.Property" & vbCrLf &
                   "Dim oPropId As Long" & vbCrLf &
                   "Dim triggerExists As Boolean = False" & vbCrLf &
                   "For oPropId = 700 To 799" & vbCrLf &
                   "    Try" & vbCrLf &
                   "        oProperty = oETPropSet.ItemByPropId(oPropId)" & vbCrLf &
                   "        If oProperty.Value = oRuleName Then triggerExists = True : Exit For" & vbCrLf &
                   "    Catch" & vbCrLf &
                   "    End Try" & vbCrLf &
                   "Next" & vbCrLf &
                   "If Not triggerExists Then" & vbCrLf &
                   "    For oPropId = 700 To 799" & vbCrLf &
                   "        Try" & vbCrLf &
                   "            oProperty = oETPropSet.ItemByPropId(oPropId)" & vbCrLf &
                   "        Catch" & vbCrLf &
                   "            oProperty = oETPropSet.Add(oRuleName, ""BeforeDocSave"" & oPropId, oPropId)" & vbCrLf &
                   "            Exit For" & vbCrLf &
                   "        End Try" & vbCrLf &
                   "    Next" & vbCrLf &
                   "End If"
        End Function
#End Region

#Region "Interação com iLogic e Eventos (Compartilhado)"
        Private Function GetILogicAutomation() As Object
            Dim iLogicAddIn As ApplicationAddIn = Nothing
            Try
                iLogicAddIn = _invApp.ApplicationAddIns.ItemById("{3bdd8d79-2179-4b11-8a5a-257b1c0263ac}")
                If iLogicAddIn Is Nothing Then Throw New Exception("Add-in iLogic não encontrado.")
                If Not iLogicAddIn.Activated Then
                    Log("   -> Ativando Add-in iLogic...")
                    iLogicAddIn.Activate()
                End If
                Return iLogicAddIn.Automation
            Catch ex As Exception
                Throw New Exception("Falha ao obter automação iLogic: " & ex.Message)
            Finally
                ReleaseComObject(iLogicAddIn)
            End Try
        End Function

        Private Sub AddOrReplaceILogicRule(targetDoc As Document, ruleName As String, ruleText As String)
            Dim iLogicAuto As Object = Nothing
            Dim rules As Object = Nothing
            Dim ruleExists As Boolean = False

            Try
                iLogicAuto = GetILogicAutomation()
                If iLogicAuto Is Nothing Then Throw New Exception("Falha ao obter iLogic Automation.")

                rules = iLogicAuto.Rules(targetDoc)
                If rules IsNot Nothing Then
                    Dim tempRuleList As System.Collections.IEnumerable = TryCast(rules, System.Collections.IEnumerable)
                    If tempRuleList IsNot Nothing Then
                        For Each R As Object In tempRuleList
                            Try
                                If R.Name = ruleName Then
                                    ruleExists = True
                                    Exit For
                                End If
                            Finally
                                ReleaseComObject(R)
                            End Try
                        Next
                    Else
                        Log("   -> AVISO: Não foi possível iterar sobre a coleção de regras.")
                    End If
                    ReleaseComObject(rules)
                End If

                If ruleExists Then
                    Log($"   -> Regra '{ruleName}' já existe. Substituindo...")
                    iLogicAuto.DeleteRule(targetDoc, ruleName)
                    iLogicAuto.AddRule(targetDoc, ruleName, ruleText)
                    Log($"   -> Regra '{ruleName}' substituída com sucesso.")
                Else
                    Log($"   -> Adicionando nova regra '{ruleName}'...")
                    iLogicAuto.AddRule(targetDoc, ruleName, ruleText)
                    Log($"   -> Regra '{ruleName}' adicionada com sucesso.")
                End If

            Catch ex As Exception
                Throw New Exception($"Falha ao adicionar/substituir regra iLogic '{ruleName}': {ex.Message}", ex)
            Finally
                ReleaseComObject(rules)
                ReleaseComObject(iLogicAuto)
            End Try
        End Sub

        Private Sub SetupEventTrigger(targetDoc As Document, ruleName As String)
            Dim eventPropSet As PropertySet = Nothing
            Dim internalName As String = "{2C540830-0723-455E-A8E2-891722EB4C3E}"
            Dim prop As Inventor.Property = Nothing
            Dim oldPropSet As PropertySet = Nothing
            Dim needsCreation As Boolean = False

            Try
                Try
                    eventPropSet = targetDoc.PropertySets.Item("iLogicEventsRules")
                    If eventPropSet.InternalName <> internalName Then
                        Log("   -> PropertySet 'iLogicEventsRules' com InternalName incorreto. Será recriado.")
                        eventPropSet.Delete()
                        needsCreation = True
                        eventPropSet = Nothing
                    Else
                        Log("   -> PropertySet 'iLogicEventsRules' encontrado.")
                    End If
                Catch exNotFound As Exception
                    Try
                        oldPropSet = targetDoc.PropertySets.Item("_iLogicEventsRules")
                        Log("   -> PropertySet com nome antigo '_iLogicEventsRules' encontrado. Será recriado com nome novo.")
                        oldPropSet.Delete()
                        needsCreation = True
                    Catch exOldNotFound As Exception
                        Log("   -> PropertySet de eventos iLogic não encontrado. Será criado.")
                        needsCreation = True
                    End Try
                End Try

                If needsCreation Then
                    Log("   -> Criando PropertySet 'iLogicEventsRules'...")
                    eventPropSet = targetDoc.PropertySets.Add("iLogicEventsRules", internalName)
                End If

                If eventPropSet Is Nothing Then Throw New Exception("Falha crítica ao obter/criar PropertySet 'iLogicEventsRules'.")

                Dim triggerExists As Boolean = False
                Dim currentProp As Inventor.Property = Nothing
                For i As Integer = 1 To eventPropSet.Count
                    Try
                        currentProp = eventPropSet.Item(i)
                        If currentProp.PropId >= 700 AndAlso currentProp.PropId < 800 AndAlso currentProp.Value?.ToString() = ruleName Then
                            triggerExists = True
                            Log($"   -> Gatilho 'Before Save' para '{ruleName}' já existe (PropId: {currentProp.PropId}).")
                            Exit For
                        End If
                    Finally
                        ReleaseComObject(currentProp)
                    End Try
                Next

                If Not triggerExists Then
                    Log($"   -> Adicionando gatilho 'Before Save' para '{ruleName}'...")
                    Dim added As Boolean = False
                    For propId As Integer = 700 To 799
                        Dim newProp As Inventor.Property = Nothing
                        Try
                            newProp = eventPropSet.Add(ruleName, $"BeforeDocSave{propId}", propId)
                            added = True
                            Log($"      -> Gatilho adicionado com PropId: {propId}.")
                            Exit For
                        Catch exPropIdUsed As COMException When exPropIdUsed.ErrorCode = -2147352567
                        Catch exOther As Exception
                            Throw New Exception($"Erro inesperado ao tentar adicionar gatilho com PropId {propId}: {exOther.Message}", exOther)
                        Finally
                            ReleaseComObject(newProp)
                        End Try
                    Next
                    If Not added Then Throw New Exception("Não foi possível encontrar um PropId livre (700-799) para o gatilho 'Before Save'. Limite de gatilhos atingido?")
                End If

            Catch ex As Exception
                Throw New Exception($"Falha ao configurar gatilho iLogic para '{ruleName}': {ex.Message}", ex)
            Finally
                ReleaseComObject(prop)
                ReleaseComObject(eventPropSet)
                ReleaseComObject(oldPropSet)
            End Try
        End Sub
#End Region

#Region "Funções Auxiliares (Log, ReleaseCOM)"
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
        End Sub

        Private Sub Log(message As String)
            _logBuilder.AppendLine($"{DateTime.Now:HH:mm:ss} - {message}")
            Me.Dispatcher?.BeginInvoke(New Action(Sub()
                                                      If LogTextBox IsNot Nothing Then
                                                          LogTextBox.Text = _logBuilder.ToString()
                                                          LogTextBox.ScrollToEnd()
                                                      End If
                                                  End Sub))
        End Sub

        Private Sub FinalizeLog()
            Me.Dispatcher?.BeginInvoke(New Action(Sub()
                                                      If LogTextBox IsNot Nothing Then
                                                          LogTextBox.Text = _logBuilder.ToString()
                                                          LogTextBox.ScrollToEnd()
                                                      End If
                                                  End Sub))
        End Sub
#End Region

    End Class

End Namespace