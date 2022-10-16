Imports System.Math

Public Class Area
    Private _w As Double
    Private _h As Double

    Public Property W As Double
        Get
            Return _w
        End Get
        Set(value As Double)
            _w = value
        End Set
    End Property

    Public Property H As Double
        Get
            Return _h
        End Get
        Set(value As Double)
            _h = value
        End Set
    End Property

    ''' <summary>
    ''' Calculate perimeter.
    ''' </summary>
    ''' <returns>Perimeter of area.</returns>
    Public Function PArea()
        Return Ceiling((_w + _h) * 2 / Pow(10, 3))
    End Function

    ''' <summary>
    ''' Calculate superficies.
    ''' </summary>
    ''' <returns>Superficies of area.</returns>
    Public Function SArea()
        Return Ceiling(_w * _h / Pow(10, 6))
    End Function
End Class
