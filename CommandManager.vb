' Arquivo: CommandManager.vb
Imports System.Linq
Imports System.Windows
Imports CitrotecAddin
Imports CitrotecAddin.CitrotecAddin ' Verifique se este namespace é necessário
Imports Inventor
Imports System.Runtime.InteropServices

Public Module CommandManager
    Public ReadOnly AllCommands As New List(Of InventorCommand)

    ' --- Funções para abrir janelas ---
    Private Sub OpenCriarPlanificacoesWindow()
        WindowManager.ShowWindow(Of CriarPlanificacoesWindow)()
    End Sub
    Private Sub OpenOrganizarBaloesWindow()
        WindowManager.ShowWindow(Of OrganizarBaloesWindow)()
    End Sub
    Private Sub OpenExportarDwgWindow()
        WindowManager.ShowWindow(Of ExportarDwgWindow)()
    End Sub
    Private Sub OpenExportarSectraWindow()
        WindowManager.ShowWindow(Of ExportarSectraWindow)()
    End Sub
    Private Sub OpenControlarDimensoesWindow()
        WindowManager.ShowWindow(Of ControlarDimensoesWindow)()
    End Sub
    Private Sub OpenCenterPointsWindow()
        WindowManager.ShowWindow(Of CenterPointsWindow)()
    End Sub
    Private Sub OpenMedirArestasWindow()
        WindowManager.ShowWindow(Of MedirArestasWindow)()
    End Sub
    Private Sub OpenVerificarRegrasWindow()
        WindowManager.ShowWindow(Of VerificarRegrasWindow)()
    End Sub
    Private Sub OpenChangeUnitsWindow()
        WindowManager.ShowWindow(Of ChangeUnitsWindow)()
    End Sub
    Private Sub OpenGrayOutPhantomsWindow()
        WindowManager.ShowWindow(Of GrayOutPhantomsWindow)()
    End Sub
    Private Sub OpenApplyCitrotecLabelWindow()
        WindowManager.ShowWindow(Of ApplyCitrotecLabelWindow)()
    End Sub
    Private Sub OpenPrepareNozzlePartWindow()
        WindowManager.ShowWindow(Of PrepareNozzlePartWindow)()
    End Sub
    Private Sub OpenCreateNozzleTableWindow()
        WindowManager.ShowWindow(Of CreateNozzleTableWindow)()
    End Sub


    Public Sub PopulateAllCommands()
        AllCommands.Clear()

        ' --- Comandos de Peça ---
        AllCommands.Add(New InventorCommand(
            "Controlar Dimensões", "CD",
            New List(Of String) From {"dimensao", "ilogic", "parametro", "comprimento", "chapa", "planificacao", "arredondar", "regra", "padrão", "g_l", "l", "lote"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kPartDocumentObject, DocumentTypeEnum.kAssemblyDocumentObject},
            AddressOf OpenControlarDimensoesWindow,
            "Cria/Atualiza regras iLogic (Peça) ou aplica regras (Seleção Montagem).",
            1))
        AllCommands.Add(New InventorCommand(
            "Criar Pontos Centrais", "CPC",
            New List(Of String) From {"centro", "ponto", "círculo", "sketch", "esboço"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kPartDocumentObject},
            AddressOf OpenCenterPointsWindow,
            "Cria pontos de esboço no centro de todos os círculos no sketch ativo.",
            1))
        AllCommands.Add(New InventorCommand(
            "Editar Propriedades Bocal", "EPB",
            New List(Of String) From {"bocal", "nozzle", "propriedade", "iproperty", "tabela", "editar", "tag", "qtd", "dn", "descricao"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kPartDocumentObject, DocumentTypeEnum.kAssemblyDocumentObject},
            AddressOf OpenPrepareNozzlePartWindow,
            "Edita/Cria as iProperties ('TAG_BOCAL', 'DESCRICAO_BOCAL', etc.) da peça para a Tabela de Bocais.",
            1))

        ' --- Comandos de Montagem ---
        AllCommands.Add(New InventorCommand(
           "Verificar/Limpar Regras iLogic", "VLR",
           New List(Of String) From {"ilogic", "regra", "limpar", "deletar", "apagar", "gerenciar", "verificar", "lista"},
           New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kAssemblyDocumentObject},
           AddressOf OpenVerificarRegrasWindow,
           "Lista peças com regras iLogic na montagem e permite limpá-las.",
           2))

        ' --- Comandos de Desenho ---
        AllCommands.Add(New InventorCommand(
            "Criar Planificações", "CP",
            New List(Of String) From {"chapa", "sheet", "metal", "planificar", "flat", "pattern"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kDrawingDocumentObject},
            AddressOf OpenCriarPlanificacoesWindow,
            "Cria vistas planificadas para todas as chapas da montagem.",
            1))
        AllCommands.Add(New InventorCommand(
            "Organizar Balões", "OB",
            New List(Of String) From {"balao", "balloon", "ordenar", "sort"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kDrawingDocumentObject},
            AddressOf OpenOrganizarBaloesWindow,
            "Organiza automaticamente balões com múltiplos líderes.",
            1))
        AllCommands.Add(New InventorCommand(
            "Exportar DWG", "ED",
            New List(Of String) From {"exportar", "dwg", "salvar", "desenho", "transferencia", "autocad"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kDrawingDocumentObject},
            AddressOf OpenExportarDwgWindow,
            "Exporta o desenho atual para DWG no padrão Citrotec.",
            1))
        AllCommands.Add(New InventorCommand(
            "Exportar SECTRA", "ES",
            New List(Of String) From {"exportar", "sectra", "excel", "xls", "lista", "materiais", "propriedades", "revisao"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kDrawingDocumentObject},
            AddressOf OpenExportarSectraWindow,
            "Exporta Lista de Materiais, Revisão e Propriedades para Excel (padrão SECTRA).",
            3))
        AllCommands.Add(New InventorCommand(
            "Realçar Phantoms/Referência", "RP",
            New List(Of String) From {"phantom", "referencia", "reference", "cinza", "cor", "realçar", "destacar", "bom", "estrutura", "vista", "linha"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kDrawingDocumentObject},
            AddressOf OpenGrayOutPhantomsWindow,
            "Altera a cor das linhas de peças Phantom/Reference para cinza nas vistas.",
            1))
        AllCommands.Add(New InventorCommand(
            "Aplicar Legenda Citrotec", "ALC",
            New List(Of String) From {"legenda", "label", "vista", "view", "item", "quantidade", "citrotec", "texto", "padrao"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kDrawingDocumentObject},
            AddressOf OpenApplyCitrotecLabelWindow,
            "Aplica legendas formatadas (com Item, Qtd, Escala) às vistas selecionadas.",
            1))
        AllCommands.Add(New InventorCommand(
            "Criar Tabela Bocais", "CTB",
            New List(Of String) From {"tabela", "bocal", "nozzle", "lista", "criar", "atualizar", "tag", "balao", "phantom", "reference"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kDrawingDocumentObject},
            AddressOf OpenCreateNozzleTableWindow,
            "Cria ou atualiza a 'TABELA DE BOCAIS' com base nos balões e componentes Phantom/Reference.",
            1))

        ' --- COMANDOS GERAIS (PEÇA E MONTAGEM) ---
        AllCommands.Add(New InventorCommand(
            "Medir Arestas", "MA",
            New List(Of String) From {"medir", "soma", "aresta", "comprimento", "total", "distancia", "edge", "measure"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kPartDocumentObject, DocumentTypeEnum.kAssemblyDocumentObject},
            AddressOf OpenMedirArestasWindow,
            "Soma o comprimento de múltiplas arestas selecionadas (em mm) e copia o valor.",
            1))
        AllCommands.Add(New InventorCommand(
            "Alterar Unidades", "AU",
            New List(Of String) From {"unidade", "medida", "mudar", "alterar", "metric", "imperial", "polegada", "milimetro", "kg", "libra", "sistema"},
            New List(Of DocumentTypeEnum) From {DocumentTypeEnum.kPartDocumentObject, DocumentTypeEnum.kAssemblyDocumentObject},
            AddressOf OpenChangeUnitsWindow,
            "Altera as unidades de comprimento e massa do documento ativo e seus componentes.",
            1))

    End Sub

End Module