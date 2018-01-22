Imports Instruments
Imports TestExecutiveLibrary
Imports System.ComponentModel

Public Class DFETB03_S2P
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

            Dim prompt1 As String = "Connect DFETB1 to DFETB2  Port " & tp.TestDescription
            If Not PromptUser(prompt1, MyBase.TestName, True) Then Throw New Exception("User aborted test")

            ' retrieve test files
            Dim SweepList As List(Of SweepParameters) = Globals.ReadXMLFileToList(Of SweepParameters)(tp.SweepListFile, "")

            ' initialize setup
            DFETB03 = CType(_instrumentList("DFETB1"), BdsDfeTestBoard03)
            DFETB03.InitializeInstrument()

            VNA.RecallReg(tp.stateName)
            VNA.IFBW = tp.VnaIfBw
            VNA.Points = tp.VnaPoints

            For Each sp As SweepParameters In SweepList

                ' setup test board switching path
                DFETB03.RfSwitchPath = sp.DFETB03SwitchPath
                DFETB03.SetOutputPath(sp.DFETB03OutputPath)

                ' define s2p file path and file name
                Dim s2pPath As String = OutputDataPath & "SN" & DutBarcode & "\S2P\"
                Dim s2pFileName As String = s2pPath & tp.FileNamePrefix & DutBarcode & sp.FileNameSuffix & tp.Port & ".s2p"

                ' get S2P data from VNA
                VNA.Trigger = vnaTrigger.SingleSweep
                VNA.SaveS2PFile(TouchtoneFormat.DB, s2pFileName)

                ' instantiate S2P class; unwrap and interpolate within passband
                Dim s2pMeas As S2pFile = New S2pFile(s2pFileName)
                s2pMeas.UnwrapS21P()
                s2pMeas.GetInterpolatedData(sp.BandStartFreq, sp.BandStopFreq, iStep)

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

                If thread.CancellationPending Then Throw New AbortTestException()
            Next

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
