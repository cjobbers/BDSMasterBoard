Public Class Complex
    'Imports System.Numerics    ' may want to utilize this class

    Private _real As Double
    Private _imag As Double


    Public Property Real As Double
        Get
            Return _real
        End Get
        Set(value As Double)
            _real = value
        End Set
    End Property

    Public Property Imaginary As Double
        Get
            Return _imag
        End Get
        Set(value As Double)
            _imag = value
        End Set
    End Property



    Public Sub New()

    End Sub

    Public Sub New(ByVal dB As Double, ByVal angle As Double)
        Try
            _real = Calc_Gamma(dB) * Math.Cos(angle * Math.PI / 180)
            _imag = Calc_Gamma(dB) * Math.Sin(angle * Math.PI / 180)

        Catch ex As Exception
            Throw New Exception("Error: Complex.New: " + ex.Message)
        End Try
    End Sub


    Public Function Calc_Gamma(ByRef dB As Double) As Double
        Return (Math.Pow(10, (dB / 20)))
    End Function


    Public Function Calc_dB() As Double
        Return (20 * Math.Log10(Math.Sqrt((Math.Pow(_real, 2) + Math.Pow(_imag, 2)))))
    End Function


    Public Function Calc_Angle() As Double
        Return (Math.Atan2(_imag, _real) * 180 / Math.PI)
    End Function


    Public Function Add(ByRef Vector2 As Complex) As Complex
        Dim temp As New Complex

        temp.Real = Me.Real + Vector2.Real
        temp.Imaginary = Me.Imaginary + Vector2.Imaginary

        Return temp
    End Function


    Public Function Add(ByVal dbl As Double) As Complex
        Dim temp As New Complex

        temp.Real = Me.Real + dbl
        temp.Imaginary = Me.Imaginary

        Return (temp)
    End Function


    Public Function Subtract(ByRef Vector2 As Complex) As Complex
        Dim temp As New Complex

        temp.Real = Me.Real - Vector2.Real
        temp.Imaginary = Me.Imaginary - Vector2.Imaginary

        Return temp
    End Function


    Public Function Subtract(ByVal dbl As Double) As Complex
        Dim temp As New Complex

        temp.Real = Me.Real - dbl
        temp.Imaginary = Me.Imaginary

        Return (temp)
    End Function


    Public Function Multiply(ByRef Vector2 As Complex) As Complex
        Dim temp As New Complex
        Dim r1, r2, r3 As Double
        Dim t1, t2, t3 As Double

        'Convert Complex to Polar
        r1 = Math.Sqrt(Math.Pow(Me.Real, 2) + Math.Pow(Me.Imaginary, 2))
        r2 = Math.Sqrt(Math.Pow(Vector2.Real, 2) + Math.Pow(Vector2.Imaginary, 2))
        t1 = Math.Atan2(Me.Imaginary, Me.Real)
        t2 = Math.Atan2(Vector2.Imaginary, Vector2.Real)

        'Perform Math on Polar
        r3 = r1 * r2
        t3 = t1 + t2

        'Convert to Complex
        temp.Real = r3 * Math.Cos(t3)
        temp.Imaginary = r3 * Math.Sin(t3)

        Return temp
    End Function


    Public Function Multiply(ByVal dbl As Double) As Complex
        Dim temp As New Complex
        Dim r As Double
        Dim t As Double

        'Convert Complex to Polar
        r = Math.Sqrt(Math.Pow(Me.Real, 2) + Math.Pow(Me.Imaginary, 2))
        t = Math.Atan2(Me.Imaginary, Me.Real)

        'Perform Math on Polar and Convert to Complex
        temp.Real = (r * dbl) * Math.Cos(t)
        temp.Imaginary = (r * dbl) * Math.Sin(t)

        Return (temp)
    End Function


    Public Function Divide(ByRef Vector2 As Complex) As Complex
        Dim temp As New Complex
        Dim r1, r2, r3 As Double
        Dim t1, t2, t3 As Double

        'Convert Complex to Polar
        r1 = Math.Sqrt(Math.Pow(Me.Real, 2) + Math.Pow(Me.Imaginary, 2))
        r2 = Math.Sqrt(Math.Pow(Vector2.Real, 2) + Math.Pow(Vector2.Imaginary, 2))
        t1 = Math.Atan2(Me.Imaginary, Me.Real)
        t2 = Math.Atan2(Vector2.Imaginary, Vector2.Real)

        'Perform Math on Polar
        r3 = r1 / r2
        t3 = t1 - t2

        'Convert back to Complex
        temp.Real = r3 * Math.Cos(t3)
        temp.Imaginary = r3 * Math.Sin(t3)

        Return temp
    End Function


    Public Function Divide(ByVal dbl As Double) As Complex
        Dim temp As New Complex
        Dim r As Double
        Dim t As Double

        'Perform Function
        'Convert Complex to Polar
        r = Math.Sqrt(Math.Pow(Me.Real, 2) + Math.Pow(Me.Imaginary, 2))
        t = Math.Atan2(Me.Imaginary, Me.Real)

        'Perform Math on Polar and Convert to Complex
        temp.Real = (r / dbl) * Math.Cos(t)
        Temp.Imaginary = (r / dbl) * Math.Sin(t)

        Return (Temp)
    End Function
End Class
