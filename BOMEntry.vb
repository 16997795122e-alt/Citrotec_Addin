' Arquivo: BOMEntry.vb (VERSÃO ATUALIZADA)
Imports System.ComponentModel
Imports Inventor

''' <summary>
''' Representa uma única linha da BOM para exibição e edição na interface.
''' Implementa INotifyPropertyChanged para que a UI seja atualizada automaticamente quando um valor muda.
''' </summary>
Public Class BOMEntry
    Implements INotifyPropertyChanged

    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Public Property OriginalRow As BOMRow

    Private Sub RaisePropertyChanged(name As String)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub

    ' Propriedade para verificar se a quantidade foi customizada
    Public ReadOnly Property IsQtyCustom As Boolean
        Get
            Dim originalTotalQty As Integer
            Integer.TryParse(OriginalRow.TotalQuantity, originalTotalQty) ' TotalQuantity é a quantidade real da montagem
            Return Qtd <> originalTotalQty
        End Get
    End Property

    ' Propriedade para o ToolTip (dica de tela)
    Public ReadOnly Property QtyToolTip As String
        Get
            If IsQtyCustom Then
                Return $"Quantidade customizada. A quantidade original na montagem é {OriginalRow.TotalQuantity}."
            Else
                Return Nothing
            End If
        End Get
    End Property

    Private _item As Integer
    Public Property Item As Integer
        Get
            Return _item
        End Get
        Set(value As Integer)
            If _item <> value Then
                _item = value
                RaisePropertyChanged(NameOf(Item))
            End If
        End Set
    End Property

    Private _partNumber As String
    Public Property PartNumber As String
        Get
            Return _partNumber
        End Get
        Set(value As String)
            If _partNumber <> value Then
                _partNumber = value
                RaisePropertyChanged(NameOf(PartNumber))
            End If
        End Set
    End Property

    ' *** ADICIONADO: Propriedade Qtd (Item Quantity) ***
    ' Usamos Integer para garantir que apenas números sejam inseridos
    Private _qtd As Integer
    Public Property Qtd As Integer
        Get
            Return _qtd
        End Get
        Set(value As Integer)
            If _qtd <> value Then
                _qtd = value
                RaisePropertyChanged(NameOf(Qtd))
                ' *** IMPORTANTE: Notifica a UI para reavaliar o aviso de customização ***
                RaisePropertyChanged(NameOf(IsQtyCustom))
                RaisePropertyChanged(NameOf(QtyToolTip))
            End If
        End Set
    End Property

    Private _unitQty As String
    Public Property UnitQty As String
        Get
            Return _unitQty
        End Get
        Set(value As String)
            If _unitQty <> value Then
                _unitQty = value
                RaisePropertyChanged(NameOf(UnitQty))
            End If
        End Set
    End Property

    Private _qty As String
    Public Property Qty As String
        Get
            Return _qty
        End Get
        Set(value As String)
            If _qty <> value Then
                _qty = value
                RaisePropertyChanged(NameOf(Qty))
            End If
        End Set
    End Property

    ' *** ADICIONADO: Propriedade Peso ***
    Private _peso As String
    Public Property Peso As String
        Get
            Return _peso
        End Get
        Set(value As String)
            If _peso <> value Then
                _peso = value
                RaisePropertyChanged(NameOf(Peso))
            End If
        End Set
    End Property

    ' *** ADICIONADO: Propriedade QD ***
    Private _qd As String
    Public Property QD As String
        Get
            Return _qd
        End Get
        Set(value As String)
            If _qd <> value Then
                _qd = value
                RaisePropertyChanged(NameOf(QD))
            End If
        End Set
    End Property

    Private _codigo As String
    Public Property Codigo As String
        Get
            Return _codigo
        End Get
        Set(value As String)
            If _codigo <> value Then
                _codigo = value
                RaisePropertyChanged(NameOf(Codigo))
            End If
        End Set
    End Property

    Private _descricao As String
    Public Property Descricao As String
        Get
            Return _descricao
        End Get
        Set(value As String)
            If _descricao <> value Then
                _descricao = value
                RaisePropertyChanged(NameOf(Descricao))
            End If
        End Set
    End Property

    Private _dimensao1 As String
    Public Property DIMENSAO1 As String
        Get
            Return _dimensao1
        End Get
        Set(value As String)
            If _dimensao1 <> value Then
                _dimensao1 = value
                RaisePropertyChanged(NameOf(DIMENSAO1))
            End If
        End Set
    End Property

    Private _dimensao2 As String
    Public Property DIMENSAO2 As String
        Get
            Return _dimensao2
        End Get
        Set(value As String)
            If _dimensao2 <> value Then
                _dimensao2 = value
                RaisePropertyChanged(NameOf(DIMENSAO2))
            End If
        End Set
    End Property

    Private _complemento As String
    Public Property COMPLEMENTO As String
        Get
            Return _complemento
        End Get
        Set(value As String)
            If _complemento <> value Then
                _complemento = value
                RaisePropertyChanged(NameOf(COMPLEMENTO))
            End If
        End Set
    End Property
End Class