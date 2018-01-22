Imports System.IO
'Imports System.Numerics

Public Class S2pFile

    Private _s2p As Dictionary(Of Long, S2P_Record)


    ' constructors
    Public Sub New()

    End Sub

    Public Sub New(ByVal fileName As String)
        Me.OpenS2pFileAsDictionary(fileName)
    End Sub


    Public Sub New(ByRef s2pArray2D As Double(,))
        Me.Transpose2DArrayToDictionary(s2pArray2D)
    End Sub


    Public Sub OpenS2pFileAsDictionary(ByVal fileName As String)
        Try
            Dim delims As Char() = {" "}
            Dim s2pList As Dictionary(Of Long, S2P_Record) = New Dictionary(Of Long, S2P_Record)
            Dim fileArray As String() = File.ReadAllLines(fileName)

            For Each line As String In fileArray
                Dim items As String() = line.Split(delims)
                If items(0) = "#" Or items(0) = "!" Or items(0) = " " Then Continue For

                Dim rec As S2P_Record = New S2P_Record
                rec.Freq = CLng(items(0))
                rec.S11M = CDbl(items(1))
                rec.S11P = CDbl(items(2))
                rec.S21M = CDbl(items(3))
                rec.S21P = CDbl(items(4))
                rec.S12M = CDbl(items(5))
                rec.S12P = CDbl(items(6))
                rec.S22M = CDbl(items(7))
                rec.S22P = CDbl(items(8))

                s2pList.Add(rec.Freq, rec)
            Next

            _s2p = s2pList
        Catch ex As Exception
            Throw New Exception("Error : S2pFile.OpenS2pFile : " + ex.Message)
        End Try
    End Sub


    Private Sub Transpose2DArrayToDictionary(ByRef s2p As Double(,))
        Try
            Dim s2pList As Dictionary(Of Long, S2P_Record) = New Dictionary(Of Long, S2P_Record)
            Dim nPoints As Integer = s2p.Length / 9

            For row As Integer = 0 To (nPoints - 1)
                Dim rec As S2P_Record = New S2P_Record
                rec.Freq = CLng(s2p(0, row))
                rec.S11M = s2p(1, row)
                rec.S11P = s2p(2, row)
                rec.S21M = s2p(3, row)
                rec.S21P = s2p(4, row)
                rec.S12M = s2p(5, row)
                rec.S12P = s2p(6, row)
                rec.S22M = s2p(7, row)
                rec.S22P = s2p(8, row)

                s2pList.Add(rec.Freq, rec)
            Next

            _s2p = s2pList
        Catch ex As Exception
            Throw New Exception("Error : S2pFile.Transpose2DArray : " + ex.Message)
        End Try
    End Sub


    Public Sub GetInBandData(BandStartFreq As Long, BandStopFreq As Long)
        Try
            BandStartFreq *= 1000000
            BandStopFreq *= 1000000

            Dim s2pList As Dictionary(Of Long, S2P_Record) = New Dictionary(Of Long, S2P_Record)
            For Each freq As Long In _s2p.Keys
                If freq >= BandStartFreq And freq <= BandStopFreq Then
                    s2pList.Add(freq, _s2p.Item(freq))
                End If
            Next

            _s2p = s2pList
        Catch ex As Exception
            Throw New Exception("Error : S2pFile.GetInBandData: " + ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' Data must be in polar notation (dB-degrees)
    ''' </summary>
    Public Sub UnwrapS21P()
        Try
            For i As Integer = 1 To _s2p.Keys.Count - 1
                Dim upper As Long = _s2p.Keys(i)
                Dim lower As Long = _s2p.Keys(i - 1)

                If _s2p.Item(upper).S21P - _s2p.Item(lower).S21P > 180 Then
                    For n As Integer = i To _s2p.Keys.Count - 1
                        Dim k As Long = _s2p.Keys(n)
                        _s2p.Item(k).S21P -= 360
                    Next
                ElseIf _s2p.Item(upper).S21P - _s2p.Item(lower).S21P < -180 Then
                    For n As Integer = i To _s2p.Keys.Count - 1
                        Dim k As Long = _s2p.Keys(n)
                        _s2p.Item(k).S21P += 360
                    Next
                End If
            Next
        Catch ex As Exception
            Throw New Exception("Error : S2pFile.UnwrapS21P: " + ex.Message)
        End Try
    End Sub


    Public Sub WrapS21P()   'add enum here?

    End Sub


    ''' <summary>
    ''' Subtract one S21 file from another - **Magnitude Only**
    ''' </summary>
    ''' <param name="file">Another instance of this class (S2pFile)</param>
    ''' <param name="corrFactor">scaling factor (multiplier) for data file passed in. Use 1 if not needed.</param>
    Public Sub SubtractS21M(ByRef file As S2pFile, ByVal corrFactor As Double)
        Try
            For Each freq As Long In _s2p.Keys
                _s2p(freq).S21M -= (file.S2P(freq).S21M * corrFactor)
            Next
        Catch ex As Exception
            Throw New Exception("Error: S2pFile.SubtractS21M: " + ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' Subtract one S21 file from another. Data must be in polar notation (dB-degrees)
    ''' </summary>
    ''' <param name="file">Another instance of this class (S2pFile)</param>
    ''' <param name="corrFactor">scaling factor (multiplier) for data file passed in. Use 1 if not needed.</param>
    Public Sub SubtractS21(ByRef file As S2pFile, ByVal corrFactor As Double)
        Try
            For Each freq As Long In _s2p.Keys

                ' DUT RAW Data
                Dim S21D As Complex = New Complex(_s2p(freq).S21M, _s2p(freq).S21P)
                'Dim S21D As Complex = PolarToComplex(_s2p(freq).S21M, _s2p(freq).S21P)

                ' Fixture/Cable Data 
                Dim F1M As Double = file.S2P(freq).S21M * corrFactor
                Dim F1P As Double = file.S2P(freq).S21P * corrFactor
                Dim S21F As Complex = New Complex(F1M, F1P)
                'Dim S21F As Complex = PolarToComplex(F1M, F1P)

                ' Complex Class
                S21D = S21D.Divide(S21F)
                _s2p(freq).S21M = S21D.Calc_dB
                _s2p(freq).S21P = S21D.Calc_Angle

                ' System.Numerics
                'S21D = Complex.Divide(S21D, S21F)
                '_s2p(freq).S21M = 20 * Complex.Log10(S21D).Magnitude
                '_s2p(freq).S21P = S21D.Phase

            Next
        Catch ex As Exception
            Throw New Exception("Error: S2pFile.SubtractS21: " + ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' Perform scalar subtraction on the Real part of the S21 vector
    ''' </summary>
    ''' <param name="dB"></param>
    Public Sub SubtractS21(ByVal dB As Double)
        Try
            For Each freq As Long In _s2p.Keys
                Dim S21D As Complex = New Complex(_s2p(freq).S21M, _s2p(freq).S21P)
                'Dim S21D As Complex = PolarToComplex(_s2p(freq).S21M, _s2p(freq).S21P)

                S21D = S21D.Divide(Math.Pow(10, (dB / 20)))
                _s2p(freq).S21M = S21D.Calc_dB
                _s2p(freq).S21P = S21D.Calc_Angle

                ' System.Numerics
                'S21D = Complex.Divide(S21D, Calc_Gamma(dB))
                '_s2p(freq).S21M = 20 * Complex.Log10(S21D).Magnitude
                '_s2p(freq).S21P = S21D.Phase

            Next
        Catch ex As Exception
            Throw New Exception("Error: S2pFile.SubtractS21: " + ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' Perform scalar addition on the Real part of the S21 vector
    ''' </summary>
    ''' <param name="dB"></param>
    Public Sub AddS21(ByVal dB As Double)
        Try
            For Each freq As Long In _s2p.Keys
                Dim S21D As Complex = New Complex(_s2p(freq).S21M, _s2p(freq).S21P)
                'Dim S21D As Complex = PolarToComplex(_s2p(freq).S21M, _s2p(freq).S21P)

                S21D = S21D.Multiply(Math.Pow(10, (dB / 20)))
                _s2p(freq).S21M = S21D.Calc_dB
                _s2p(freq).S21P = S21D.Calc_Angle

                ' System.Numerics
                'S21D = Complex.Multiply(S21D, Calc_Gamma(dB))
                '_s2p(freq).S21M = 20 * Complex.Log10(S21D).Magnitude
                '_s2p(freq).S21P = S21D.Phase

            Next
        Catch ex As Exception
            Throw New Exception("Error: S2pFile.AddS21: " + ex.Message)
        End Try
    End Sub


    'Private Function PolarToComplex(ByVal dB As Double, ByVal degrees As Double) As Complex
    '    Dim magnitude As Double = Math.Pow(10, (dB / 20))
    '    Dim phase As Double = degrees * (Math.PI / 180)
    '    Return Complex.FromPolarCoordinates(magnitude, phase)
    'End Function


    'Private Function Calc_Gamma(ByRef dB As Double) As Double
    '    Return (Math.Pow(10, (dB / 20)))
    'End Function


    Public Function interpolate(ByVal x As Double, ByVal x0 As Double, ByVal x1 As Double, ByVal y0 As Double, ByVal y1 As Double) As Double
        If (x1 - x0) = 0 Then
            Return (y0 + y1) / 2
        Else
            Return (y0 + (x - x0) * (y1 - y0) / (x1 - x0))
        End If
    End Function


    ''' <summary>
    ''' Returns new collection with keys at specified increment. Note: Phase can extend > +/-180 after interpolation
    ''' </summary>
    ''' <param name="startFreq">Start Frequency (MHz)</param>
    ''' <param name="stopFreq">Stop Frequency (MHz)</param>
    ''' <param name="increment">Output Step (MHz)</param>
    Public Sub GetInterpolatedData(ByVal startFreq As Long, ByVal stopFreq As Long, ByVal increment As Long)
        Try
            Dim s2pList As Dictionary(Of Long, S2P_Record) = New Dictionary(Of Long, S2P_Record)
            startFreq *= 1000000
            stopFreq *= 1000000
            increment *= 1000000

            If _s2p.Count < 2 Then Throw New Exception("Not enough data points to perform interpolation")

            ' unwrap before performing interpolation to avoid wrap-around errors
            'Me.UnwrapS21P()

            For freq As Long = startFreq To stopFreq Step increment
                Dim rec As S2P_Record = New S2P_Record

                ' see if the desired freq point already exists in the collection
                If _s2p.ContainsKey(freq) Then
                    rec = _s2p(freq)
                Else
                    ' find freq points in collection that are above and below desired freq
                    rec.Freq = freq
                    For i As Integer = 1 To (_s2p.Keys.Count - 1)
                        Dim upper As Long = _s2p.Keys(i)
                        Dim lower As Long = _s2p.Keys(i - 1)

                        If upper > freq Or i = (_s2p.Keys.Count - 1) Then
                            rec.S11M = interpolate(freq, _s2p(lower).Freq, _s2p(upper).Freq, _s2p(lower).S11M, _s2p(upper).S11M)
                            rec.S11P = interpolate(freq, _s2p(lower).Freq, _s2p(upper).Freq, _s2p(lower).S11P, _s2p(upper).S11P)
                            rec.S21M = interpolate(freq, _s2p(lower).Freq, _s2p(upper).Freq, _s2p(lower).S21M, _s2p(upper).S21M)
                            rec.S21P = interpolate(freq, _s2p(lower).Freq, _s2p(upper).Freq, _s2p(lower).S21P, _s2p(upper).S21P)
                            rec.S12M = interpolate(freq, _s2p(lower).Freq, _s2p(upper).Freq, _s2p(lower).S12M, _s2p(upper).S12M)
                            rec.S12P = interpolate(freq, _s2p(lower).Freq, _s2p(upper).Freq, _s2p(lower).S12P, _s2p(upper).S12P)
                            rec.S22M = interpolate(freq, _s2p(lower).Freq, _s2p(upper).Freq, _s2p(lower).S22M, _s2p(upper).S22M)
                            rec.S22P = interpolate(freq, _s2p(lower).Freq, _s2p(upper).Freq, _s2p(lower).S22P, _s2p(upper).S22P)
                            Exit For
                        End If
                    Next
                End If

                s2pList.Add(rec.Freq, rec)
            Next

            _s2p = s2pList
        Catch ex As Exception
            Throw New Exception("Error: S2pFile.GetInterpolatedData: " + ex.Message)
        End Try
    End Sub


    Public Sub RoundToDecimal(ByVal numPlaces As Integer)
        Try
            For Each freq As Long In _s2p.Keys
                _s2p(freq).S11M = Math.Round(_s2p(freq).S11M, numPlaces)
                _s2p(freq).S11P = Math.Round(_s2p(freq).S11P, numPlaces)
                _s2p(freq).S21M = Math.Round(_s2p(freq).S21M, numPlaces)
                _s2p(freq).S21P = Math.Round(_s2p(freq).S21P, numPlaces)
                _s2p(freq).S12M = Math.Round(_s2p(freq).S12M, numPlaces)
                _s2p(freq).S12P = Math.Round(_s2p(freq).S12P, numPlaces)
                _s2p(freq).S22M = Math.Round(_s2p(freq).S22M, numPlaces)
                _s2p(freq).S22P = Math.Round(_s2p(freq).S11P, numPlaces)
            Next
        Catch ex As Exception
            Throw New Exception("Error: S2pFile.RoundToDecimal: " + ex.Message)
        End Try
    End Sub


    Public Function GetTraceStats() As TraceStatsRecord
        Return Me.GetTraceStats(_s2p)
    End Function


    Public Function GetTraceStats(ByRef traceData As Dictionary(Of Long, S2P_Record)) As TraceStatsRecord
        Try
            Dim rec As New TraceStatsRecord
            rec.Points = traceData.Count
            rec.MinValue = 999
            rec.MaxValue = -999
            rec.MeanValue = 0
            rec.StartFreq = traceData.Keys.First
            rec.StopFreq = traceData.Keys.Last

            For Each freq As Long In traceData.Keys
                ' get max trace value
                If traceData(freq).S21M > rec.MaxValue Then
                    rec.MaxValue = traceData(freq).S21M
                    rec.MaxFreq = traceData(freq).Freq
                End If

                ' get min trace value
                If traceData(freq).S21M < rec.MinValue Then
                    rec.MinValue = traceData(freq).S21M
                    rec.MinFreq = traceData(freq).Freq
                End If

                ' sum for mean calc
                rec.MeanValue += traceData(freq).S21M
            Next

            rec.MeanValue /= rec.Points

            ' round values
            rec.MaxValue = Math.Round(rec.MaxValue, 2)
            rec.MinValue = Math.Round(rec.MinValue, 2)
            rec.MeanValue = Math.Round(rec.MeanValue, 2)

            Return rec
        Catch ex As Exception
            Throw New Exception("Error: S2pFile.GetTraceStats: " + ex.Message)
        End Try
    End Function


    Public Function S2pFileToStringList() As List(Of String())
        Try
            Dim strArray As List(Of String()) = New List(Of String())

            For Each freq As Long In _s2p.Keys
                Dim values As List(Of String) = New List(Of String)
                values.Add(_s2p(freq).Freq.ToString)
                values.Add(_s2p(freq).S11M.ToString())
                values.Add(_s2p(freq).S11P.ToString())
                values.Add(_s2p(freq).S21M.ToString())
                values.Add(_s2p(freq).S21P.ToString())
                values.Add(_s2p(freq).S12M.ToString())
                values.Add(_s2p(freq).S12P.ToString())
                values.Add(_s2p(freq).S22M.ToString())
                values.Add(_s2p(freq).S22P.ToString())

                strArray.Add(values.ToArray)
            Next

            Return strArray
        Catch ex As Exception
            Throw New Exception("Error : S2PFileToList : " + ex.Message)
        End Try
    End Function


    ' Properties ------------------

    Public ReadOnly Property S2P As Dictionary(Of Long, S2P_Record)
        Get
            Return _s2p
        End Get
        'Set(value As Dictionary(Of Long, S2P_Record))
        '    _s2pMatrix_d = value
        'End Set
    End Property



#Region "Unused Methods And Properties"

    'Public Sub New(ByRef s2pList As List(Of Double()))
    '    _s2pMatrix = s2pList
    'End Sub

    'Public Sub New(ByVal fileName As String, ByVal bandStart As Double, bandStop As Double)
    '    Me.OpenS2pFileAsDictionary(fileName)
    '    Me.GetInBandData(bandStart, bandStop)  ' consider doing in open file to save copying step
    'End Sub

    'Public Sub New(ByRef s2pArray2D As Double(,), ByVal bandStart As Double, bandStop As Double)
    '    Me.Transpose2DArrayToDictionary(s2pArray2D)
    '    Me.GetInBandData(bandStart, bandStop)
    'End Sub


    'Public Function OpenS2pFileAsText(ByVal fileName As String) As List(Of String())
    '    Try
    '        Dim delims As Char() = {" "}
    '        Dim strArray As List(Of String()) = New List(Of String())
    '        Dim fileArray As String() = File.ReadAllLines(fileName)

    '        For Each line As String In fileArray
    '            Dim delimLine As String() = line.Split(delims)
    '            strArray.Add(delimLine)
    '        Next

    '        Return strArray
    '    Catch ex As Exception
    '        Throw New Exception("Error Reading S2P File : " + ex.Message)
    '    End Try
    'End Function


    'Public Sub OpenS2pFileAsList(ByVal fileName As String)
    '    Try
    '        Dim delims As Char() = {" "}
    '        Dim dblArray As List(Of Double()) = New List(Of Double())
    '        Dim fileArray As String() = File.ReadAllLines(fileName)

    '        For Each line As String In fileArray
    '            Dim delimLine As String() = line.Split(delims)
    '            If delimLine(0).First = "#" Or delimLine(0).First = "!" Or delimLine(0).First = " " Then Continue For

    '            Dim values As List(Of Double) = New List(Of Double)
    '            For Each item As String In delimLine
    '                Dim value As Double = CDbl(item)
    '                values.Add(value)
    '            Next
    '            dblArray.Add(values.ToArray)
    '        Next

    '        _s2pList = dblArray
    '    Catch ex As Exception
    '        Throw New Exception("Error Reading S2P File : " + ex.Message)
    '    End Try
    'End Sub


    'Private Function Transpose2DArrayToList(ByRef s2p As Double(,)) As List(Of Double())
    '    Try
    '        Dim s2pList As List(Of Double()) = New List(Of Double())
    '        Dim nPoints As Integer = s2p.Length / 9

    '        For row As Integer = 0 To (nPoints - 1)
    '            Dim values(8) As Double

    '            For i As Integer = 0 To 8
    '                values(i) = s2p(i, row)
    '            Next

    '            s2pList.Add(values)
    '        Next

    '        Return s2pList
    '    Catch ex As Exception
    '        Throw New Exception("Error transposing 2D array : " + ex.Message)
    '    End Try
    'End Function


    'Private Function TransposeListToDictionary(ByRef s2p As List(Of Double())) As Dictionary(Of Long, S2P_Record)
    '    Try
    '        Dim s2pList As Dictionary(Of Long, S2P_Record) = New Dictionary(Of Long, S2P_Record)

    '        For i As Integer = 0 To (s2p.Count - 1)
    '            Dim rec As S2P_Record = New S2P_Record
    '            rec.Freq = CLng(s2p(i)(0))
    '            rec.S11M = s2p(i)(1)
    '            rec.S11P = s2p(i)(2)
    '            rec.S21M = s2p(i)(3)
    '            rec.S21P = s2p(i)(4)
    '            rec.S12M = s2p(i)(5)
    '            rec.S12P = s2p(i)(6)
    '            rec.S22M = s2p(i)(7)
    '            rec.S22P = s2p(i)(8)

    '            s2pList.Add(rec.Freq, rec)
    '        Next

    '        Return s2pList
    '    Catch ex As Exception
    '        Throw New Exception("Error transposing 2D array : " + ex.Message)
    '    End Try
    'End Function


    'Public Function getInBandData(BandStartFreq As Double, BandStopFreq As Double) As List(Of Double())
    '    Try
    '        Dim nPoints As Integer = _s2pList.Count
    '        Dim lowerIndex As Integer
    '        Dim upperIndex As Integer

    '        ' define in-band subarray indexes if band start and band stop values are valid. 
    '        ' else use all freq points
    '        If BandStartFreq > 0 Then
    '            lowerIndex = getIndex(_s2pList, BandStartFreq * 1000000.0)
    '            ' lowerIndex = s2p.IndexOf(...)
    '        Else
    '            lowerIndex = 0
    '        End If

    '        If BandStopFreq > 0 Then
    '            upperIndex = getIndex(_s2pList, BandStopFreq * 1000000.0)
    '        Else
    '            upperIndex = nPoints - 1
    '        End If

    '        Dim nInBandPoints As Integer = upperIndex - lowerIndex + 1
    '        If nInBandPoints <= 0 Then nInBandPoints = 1

    '        Dim _sxx As List(Of Double()) = New List(Of Double())
    '        For i As Integer = lowerIndex To upperIndex
    '            _sxx.Add(_s2pList(i))
    '        Next

    '        Return _sxx
    '    Catch ex As Exception
    '        Throw New Exception("Error : getInBandData function")
    '    End Try
    'End Function


    '''' <summary>
    '''' Returns array subset of an S2P matrix between two frequencies
    '''' </summary>
    '''' <param name="s2p"></param>
    '''' <param name="BandStartFreq">MHz</param>
    '''' <param name="BandStopFreq">MHz</param>
    '''' <param name="index">[0 to 8] (1 = S11M, 2 = S11P, 3 = S21M, etc)</param>
    '''' <returns></returns>
    'Public Function getInBandData(ByRef s2p As Double(,), BandStartFreq As Double, BandStopFreq As Double, index As Integer) As Double(,)
    '    Try
    '        Dim nPoints As Integer = s2p.Length / 9
    '        Dim lowerIndex As Integer
    '        Dim upperIndex As Integer

    '        ' define in-band subarray indexes if band start and band stop values are valid. 
    '        ' else use all freq points
    '        If BandStartFreq > 0 Then
    '            lowerIndex = getIndex(s2p, BandStartFreq * 1000000.0)
    '        Else
    '            lowerIndex = 0
    '        End If

    '        If BandStopFreq > 0 Then
    '            upperIndex = getIndex(s2p, BandStopFreq * 1000000.0)
    '        Else
    '            upperIndex = nPoints - 1
    '        End If

    '        Dim nInBandPoints As Integer = upperIndex - lowerIndex + 1
    '        If nInBandPoints <= 0 Then nInBandPoints = 1

    '        Dim _sxx(1, nInBandPoints - 1) As Double
    '        Array.Copy(s2p, 0 * nPoints + lowerIndex, _sxx, 0, nInBandPoints)                   ' freq
    '        Array.Copy(s2p, index * nPoints + lowerIndex, _sxx, nInBandPoints, nInBandPoints)   ' S21M

    '        Return _sxx
    '    Catch ex As Exception
    '        Throw New Exception("Error : getInBandData function")
    '    End Try
    'End Function


    'Private Function getIndex(ByRef s2p As List(Of S2P_Record), ByRef desiredFreq As Double) As Integer
    '    Try
    '        If s2p(0).Freq > desiredFreq Then Return 0
    '        If s2p(s2p.Count - 1).Freq < desiredFreq Then Return s2p(0).Freq

    '        For i As Integer = 0 To (s2p.Count - 1)
    '            If s2p(i).Freq >= desiredFreq Then Return i
    '        Next

    '        Return 0
    '    Catch ex As Exception
    '        Throw New Exception("Error : getIndex function")
    '    End Try
    'End Function

    '''' <summary>
    '''' Returns the array index closest to the desired frequency from an S2P matrix
    '''' </summary>
    '''' <param name="s2p"></param>
    '''' <param name="desiredFreq"></param>
    '''' <returns></returns>
    'Private Function getIndex(ByRef s2p(,) As Double, ByRef desiredFreq As Double) As Integer
    '    Try
    '        Dim nPoints As Integer = s2p.Length / 9
    '        If s2p(0, 0) > desiredFreq Then Return 0
    '        If s2p(0, nPoints - 1) < desiredFreq Then Return s2p(0, nPoints - 1)

    '        For i As Integer = 0 To (nPoints - 1)
    '            If s2p(0, i) >= desiredFreq Then Return i
    '        Next

    '        Return 0
    '    Catch ex As Exception
    '        Throw New Exception("Error : getIndex function")
    '    End Try
    'End Function


    'Private Function getIndex(ByRef s2p As List(Of Double()), ByRef desiredFreq As Double) As Integer
    '    Try
    '        If s2p(0).First > desiredFreq Then Return 0
    '        If s2p(s2p.Count - 1).First < desiredFreq Then Return s2p(0).Last

    '        For i As Integer = 0 To (s2p.Count - 1)
    '            If s2p(i)(0) >= desiredFreq Then Return i
    '        Next

    '        Return 0
    '    Catch ex As Exception
    '        Throw New Exception("Error : getIndex function")
    '    End Try
    'End Function


    'Public Function GetTraceStats(ByRef traceData As List(Of Double())) As TraceStatsRecord
    '    Try
    '        Dim rec As New TraceStatsRecord
    '        rec.Points = traceData.Count
    '        rec.MinValue = 999
    '        rec.MaxValue = -999
    '        rec.MeanValue = 0
    '        rec.StartFreq = traceData(0)(0)
    '        rec.StopFreq = traceData(traceData.Count - 1)(0)

    '        For n As Integer = 0 To traceData.Count - 1
    '            ' get max trace value
    '            If traceData(n)(3) > rec.MaxValue Then
    '                rec.MaxValue = traceData(n)(3)
    '                rec.MaxFreq = traceData(n)(0)
    '            End If

    '            ' get min trace value
    '            If traceData(n)(3) < rec.MinValue Then
    '                rec.MinValue = traceData(n)(3)
    '                rec.MinFreq = traceData(n)(0)
    '            End If

    '            ' sum for mean calc
    '            rec.MeanValue += traceData(n)(3)
    '        Next

    '        rec.MeanValue /= rec.Points

    '        ' round values
    '        rec.MaxValue = Math.Round(rec.MaxValue, 2)
    '        rec.MinValue = Math.Round(rec.MinValue, 2)
    '        rec.MeanValue = Math.Round(rec.MeanValue, 2)

    '        Return rec
    '    Catch ex As Exception
    '        Throw New Exception("Error: GetTraceStats")
    '    End Try
    'End Function



    'Public Function GetTraceStats(ByRef traceData As Double(,)) As TraceStatsRecord
    '    Try
    '        Dim rec As New TraceStatsRecord
    '        rec.Points = traceData.Length / 2
    '        rec.MinValue = 999
    '        rec.MaxValue = -999
    '        rec.MeanValue = 0
    '        rec.StartFreq = traceData(0, 0)
    '        rec.StopFreq = traceData(0, rec.Points - 1)

    '        For n As Integer = 0 To rec.Points - 1
    '            ' get max trace value
    '            If traceData(1, n) > rec.MaxValue Then
    '                rec.MaxValue = traceData(1, n)
    '                rec.MaxFreq = traceData(0, n)
    '            End If

    '            ' get min trace value
    '            If traceData(1, n) < rec.MinValue Then
    '                rec.MinValue = traceData(1, n)
    '                rec.MinFreq = traceData(0, n)
    '            End If

    '            ' sum for mean calc
    '            rec.MeanValue += traceData(1, n)
    '        Next

    '        rec.MeanValue /= rec.Points

    '        ' round values
    '        rec.MaxValue = Math.Round(rec.MaxValue, 2)
    '        rec.MinValue = Math.Round(rec.MinValue, 2)
    '        rec.MeanValue = Math.Round(rec.MeanValue, 2)

    '        Return rec
    '    Catch ex As Exception
    '        Throw New Exception("Error: GetTraceStats")
    '    End Try
    'End Function


    '''' <summary>
    '''' Creates a linear ramp pattern array from Fstart to Fstop using nPoints
    '''' </summary>
    '''' <param name="Fstart"></param>
    '''' <param name="Fstop"></param>
    '''' <param name="nPoints"></param>
    '''' <returns></returns>
    'Public Function getFreqArray(ByRef Fstart As Long, ByRef Fstop As Long, ByRef nPoints As Integer) As Long()
    '    Try
    '        Dim FreqArray(nPoints - 1) As Long
    '        Dim fStep As Long = (Fstop - Fstart) / (nPoints - 1)

    '        For i As Integer = 0 To (nPoints - 1)
    '            FreqArray(i) = Fstart + i * fStep
    '        Next

    '        Return FreqArray
    '    Catch ex As Exception
    '        Throw New Exception("Error : getFreqArray function")
    '    End Try
    'End Function


    'Public Function Array2DSlice(ByRef inputArray(,) As Double, index As Integer) As Double()
    '    Try
    '        Dim outArray(inputArray.Length - 1) As Double
    '        For i As Integer = 0 To (inputArray.Length - 1)
    '            outArray(i) = inputArray(index, i)
    '        Next

    '        Return outArray
    '    Catch ex As Exception
    '        Throw New Exception("Error : Array2DSlice function")
    '    End Try
    'End Function


    'Public Function Array2DSlice(ByRef inputArray(,) As Long, index As Integer) As Long()
    '    Try
    '        Dim outArray(inputArray.Length - 1) As Long
    '        For i As Integer = 0 To (inputArray.Length - 1)
    '            outArray(i) = inputArray(index, i)
    '        Next

    '        Return outArray
    '    Catch ex As Exception
    '        Throw New Exception("Error : Array2DSlice function")
    '    End Try
    'End Function


    'Public ReadOnly Property S2pToList As List(Of Double())
    '    Get
    '        Return _s2pList
    '    End Get
    '    'Set(value As List(Of Double()))
    '    '    _s2pList = value
    '    'End Set
    'End Property


    ' this only works for angle <90 deg
    'Public Sub SubtractS21MP(ByRef file As S2pFile, ByVal corrFactor As Double)
    '    Try
    '        For Each freq As Long In _s2p.Keys
    '            Dim A As Double = _s2p(freq).S21P * Math.PI / 180
    '            Dim B As Double = file.S2pToDictionary(freq).S21P * corrFactor * Math.PI / 180

    '            ' magnitude
    '            _s2p(freq).S21M -= (file.S2pToDictionary(freq).S21M * corrFactor)

    '            'phase
    '            Dim r = Math.Sin(A) * Math.Cos(B) - Math.Sin(B) * Math.Cos(A)
    '            _s2p(freq).S21P = Math.Asin(r) * 180 / Math.PI
    '        Next
    '    Catch ex As Exception
    '        Throw New Exception("Error: S2pFile.SubtractS21M: " + ex.Message)
    '    End Try
    'End Sub


#End Region


End Class
