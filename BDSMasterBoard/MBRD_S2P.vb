Imports Instruments
Imports TestExecutiveLibrary
Imports System.ComponentModel

Public Class MBRD_S2P
    Inherits TestsCommon

    Public Sub New(ByVal instrumentList As Dictionary(Of String, Instrument), ByVal params As TestParameters)
        MyBase.New(instrumentList, params)
    End Sub


    Public Overrides Sub Run(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim thread As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim iStep As Long = 10  'interpolation step size in MHz
        Dim runtest As Boolean = True


        Try
            If thread.CancellationPending Then Throw New AbortTestException()
            MyBase.PreTest()

            Dim prompt1 As String = "Insert Test Tile+ into MBRD port " & tp.Port
            If Not PromptUser(prompt1, MyBase.TestName, True) Then Throw New Exception("User aborted test")

            ' retrieve test files
            Dim SweepList As List(Of SweepParameters) = Globals.ReadXMLFileToList(Of SweepParameters)(tp.SweepListFile, "")

            ' initialize setup
            DFETB = CType(_instrumentList("DFETB1"), BdsDfeTestBoard)
            TTP = CType(_instrumentList("TTP1"), BdsTestTilePlus)
            DFETB.InitializeInstrument()
            TTP.InitializeInstrument()

            VNA.RecallReg(tp.stateName)
            VNA.IFBW = 10000
            VNA.Points = 401
            'VNA.Measure = vnaMeasure.S21
            'VNA.Format = vnaFormat.LogMag
            'Dim fStart As Long = VNA.StartFreq
            'Dim fStop As Long = VNA.StopFreq


            'While runtest

            For Each sp As SweepParameters In SweepList

                ' setup test board switching path
                DFETB.RfSwitchPath = sp.DfeTbSwitchPath
                TTP.RfSwitchPath = sp.TtpSwitchPath

                ' open input (fixture A) S2P file to be de-embedded
                Dim FA2p As S2pFile = New S2pFile(sp.InputCalFileName)
                FA2p.UnwrapS21P()
                FA2p.GetInterpolatedData(sp.BandStartFreq, sp.BandStopFreq, iStep)

                ' open output (Fixture B) S2P file to be de-embedded
                Dim FB2p As S2pFile = New S2pFile(sp.OutputCalFileName)
                FB2p.UnwrapS21P()
                FB2p.GetInterpolatedData(sp.BandStartFreq, sp.BandStopFreq, iStep)

                ' define s2p file path and file name
                Dim s2pPath As String = OutputDataPath & "SN" & DutBarcode & "\S2P\"
                Dim s2pFileName As String = s2pPath & tp.FileNamePrefix & DutBarcode & sp.FileNameSuffix & tp.Port & ".s2p"

                ' get S2P data from VNA
                VNA.Trigger = vnaTrigger.SingleSweep
                'Dim data() As Double = VNA.ReadTraceData()      'not working
                VNA.SaveS2PFile(TouchtoneFormat.DB, s2pFileName)

                ' instantiate S2P class; unwrap and interpolate within passband
                Dim s2pMeas As S2pFile = New S2pFile(s2pFileName)
                s2pMeas.UnwrapS21P()
                s2pMeas.GetInterpolatedData(sp.BandStartFreq, sp.BandStopFreq, iStep)

                'de-Embed fixture S21 Magnitude and Phase
                s2pMeas.SubtractS21(FA2p, sp.InputCalCorrFactor)
                s2pMeas.SubtractS21(FB2p, sp.OutputCalCorrFactor)

                ' factor in offset(dB) if required due to splitter/combiner
                s2pMeas.AddS21(sp.OffsetS21M)

                ' round all values to n places
                s2pMeas.RoundToDecimal(2)

                Dim stats As TraceStatsRecord = s2pMeas.GetTraceStats()

                ' write to status window
                MyBase.OnStatusMessageEvent(sp.FileNameSuffix & tp.Port & " In-Band Stats : ")
                MyBase.OnStatusMessageEvent("Mean : " & stats.MeanValue.ToString("0.00dB (") & (stats.StartFreq / 1000000.0).ToString("0 to ") & (stats.StopFreq / 1000000.0).ToString("0MHz)"))
                MyBase.OnStatusMessageEvent("Min : " & stats.MinValue.ToString("0.00dB @ ") & (stats.MinFreq / 1000000.0).ToString("0MHz"))
                MyBase.OnStatusMessageEvent("Max : " & stats.MaxValue.ToString("0.00dB @ ") & (stats.MaxFreq / 1000000.0).ToString("0MHz"))

                ' write data to grid
                Dim outHeader As String = sp.FileNameSuffix & tp.Port & " S21M  Mean"
                DataToGrid(outHeader, stats.MeanValue, sp.LL, sp.UL, sp.Units)

                tp.FileNameSuffix = sp.FileNameSuffix
                tp.InputCalFileName = sp.InputCalFileName
                tp.OutputCalFileName = sp.OutputCalFileName

                ' save data
                MyBase.WriteXmlS2PFile(s2pMeas)
                MyBase.WriteCsvS2PFile(s2pMeas)
                MyBase.OnStatusMessageEvent("")


                ' prompt user to retest if fail
                'If testResult = TestStatus.Fail Then
                '    runtest = PromptUserToRetry()
                'Else
                '    runtest = False
                'End If

                If thread.CancellationPending Then Throw New AbortTestException()

            Next
            'End While


            MyBase.PostTest()

        Catch ex As AbortTestException
            e.Cancel = True
            MyBase.HandleAbortTestException(ex)

        Catch ex As Exception
            MyBase.HandleTestFailException(ex)

        Finally
            If testResult <> TestStatus.Aborted Then WriteCsvLogFile()
        End Try
    End Sub

End Class
