' Arquivo: BooleanToVisibilityNegatedConverter.vb
Imports System.Globalization
Imports System.Windows
Imports System.Windows.Data

Public Class BooleanToVisibilityNegatedConverter
    Implements IValueConverter

    Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
        Dim boolValue As Boolean = False
        If TypeOf value Is Boolean Then
            boolValue = CBool(value)
        End If

        ' Retorna Collapsed se for True (Admin), Visible se for False (Não Admin)
        Return If(boolValue, Visibility.Collapsed, Visibility.Visible)
    End Function

    Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
        Throw New NotImplementedException("Cannot convert back from Visibility to Boolean")
    End Function
End Class