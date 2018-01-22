Imports TimeDate = Microsoft.VisualBasic.DateAndTime
Imports Instruments
Imports Instruments.Globals
Imports TestExecutiveLibrary
Imports System.ComponentModel
Imports LabVIEW
Imports System.Collections.Generic
Imports System.IO

Public Class S2P
    Inherits TestsCommon

    Public Sub New(ByVal instrumentList As Dictionary(Of String, Instrument), ByVal params As TestParameters)
        MyBase.New(instrumentList, params)
    End Sub


    ' This is the manual version of MBRD testing using the master test board and test tile (no automated switching)

    Public Overrides Sub Run(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim thread As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim gridData As String()
        Dim params As TestParameters = tp
        Dim ports As String()
        Dim iStep As Long = 10  'interpolation step size in MHz
        Dim runtest As Boolean = True


        Try
            If thread.CancellationPending Then Throw New AbortTestException()

            If params.PortArray IsNot Nothing Then
                Dim delims As Char() = {","}
                ports = params.PortArray.Split(delims)
            Else
                Throw New Exception("Error: Test ports are not defined in config file")
            End If

            ' open input/output (fixture A and Fixture B) S2P files to be de-embedded
            Dim FA2p As S2pFile = New S2pFile(params.InputCalFileName)
            FA2p.UnwrapS21P()
            FA2p.GetInterpolatedData(params.BandStartFreq, params.BandStopFreq, iStep)

            Dim FB2p As S2pFile = New S2pFile(params.OutputCalFileName)
            FB2p.UnwrapS21P()
            FB2p.GetInterpolatedData(params.BandStartFreq, params.BandStopFreq, iStep)

            Debug.Print("FA2P File : " + params.InputCalFileName)
            Debug.Print("FB2P File : " + params.OutputCalFileName)
            Debug.Print("Band Start Freq : " + params.BandStartFreq.ToString("0MHz"))
            Debug.Print("Band Stop Freq : " + params.BandStopFreq.ToString("0MHz"))


            ' loop thru the ports specified in config file
            'For port As Integer = portStart To portStop
            For Each port As String In ports

                While runtest
                    MyBase.PreTest()
                    params.Port = port
                    Dim prompt1 As String = "Connect " & params.FileNameSuffix & port
                    If Not PromptUser(prompt1, MyBase.TestName, True) Then Throw New Exception("User aborted test")

                    Dim s2pPath As String = OutputDataPath & "SN" & DutBarcode & "\S2P\"
                    Dim s2pFileName As String = s2pPath & params.FileNamePrefix & DutBarcode & params.FileNameSuffix & port & ".s2p"


                    ' get S2P data from VNA
                    VNA.RecallReg(params.stateName)
                    'Dim fStart As Long = VNA.StartFreq
                    'Dim fStop As Long = VNA.StopFreq
                    VNA.IFBW = 10000
                    'VNA.Points = 401
                    'VNA.Measure = vnaMeasure.S21
                    'VNA.Format = vnaFormat.LogMag
                    VNA.Trigger = vnaTrigger.SingleSweep
                    'Dim data() As Double = VNA.ReadTraceData()      'not working

                    VNA.SaveS2PFile(TouchtoneFormat.DB, s2pFileName)

                    ' instantiate S2P calss and add returned values  
                    Dim s2pMeas As S2pFile = New S2pFile(s2pFileName)
                    'Dim s2pMeas As S2pFile = New S2pFile(CType(paramValues(3), Double(,)))
                    'Dim s2pMeas As S2pFile = New S2pFile("c:\matlab\mbrd\8329870006\mbrd-e1-8329870006_BDSF_C1_J1.s2p") ' temporary - pull measured s2p file for validation
                    s2pMeas.UnwrapS21P()
                    s2pMeas.GetInterpolatedData(params.BandStartFreq, params.BandStopFreq, iStep)

                    'de-Embed S21 Magnitude and Phase
                    s2pMeas.SubtractS21(FA2p, params.InputCalCorrFactor)
                    s2pMeas.SubtractS21(FB2p, params.OutputCalCorrFactor)

                    ' factor in offset(dB) if required (BDSF, BDSR -3.3dB)
                    s2pMeas.AddS21(params.OffsetS21M)

                    ' round all values to n places
                    s2pMeas.RoundToDecimal(2)

                    'Dim stats As TraceStatsRecord = New TraceStatsRecord()
                    Dim stats As TraceStatsRecord = s2pMeas.GetTraceStats()

                    ' write to status window
                    MyBase.OnStatusMessageEvent(params.FileNameSuffix & port & " In-Band Stats : ")
                    MyBase.OnStatusMessageEvent("Mean : " & stats.MeanValue.ToString("0.00dB (") & (stats.StartFreq / 1000000.0).ToString("0 to ") & (stats.StopFreq / 1000000.0).ToString("0MHz)"))
                    MyBase.OnStatusMessageEvent("Min : " & stats.MinValue.ToString("0.00dB @ ") & (stats.MinFreq / 1000000.0).ToString("0MHz"))
                    MyBase.OnStatusMessageEvent("Max : " & stats.MaxValue.ToString("0.00dB @ ") & (stats.MaxFreq / 1000000.0).ToString("0MHz"))

                    ' write data to grid
                    Dim outHeader As String = params.FileNameSuffix & port & " S21M  Mean"
                    Dim measData As String = stats.MeanValue.ToString("0.00")
                    gridData = {outHeader, params.LL.ToString, measData, params.UL.ToString, params.Units}
                    Dim testResult As TestStatus = GetTestResult(stats.MeanValue, params.LL, params.UL)
                    OnAddDataToGridEvent(gridData, testResult)


                    ' add additional file header information
                    'SqlTable = params.TestName

                    ' save data
                    MyBase.WriteXmlS2PFile(s2pMeas)
                    MyBase.WriteCsvS2PFile(s2pMeas)
                    MyBase.OnStatusMessageEvent("")

                    ' prompt user to retest if fail
                    If testResult = TestStatus.Fail Then
                        If Not PromptUserToRetry() Then runtest = False
                    Else
                        runtest = False
                    End If

                    If thread.CancellationPending Then Throw New AbortTestException()
                End While

                runtest = True
            Next

            MyBase.PostTest()

        Catch ex As AbortTestException
            e.Cancel = True
            MyBase.OnUpdateTestStatusEvent(MyBase.TestNumber, TestStatus.Aborted)
            MyBase.OnStatusMessageEvent(params.TestDescription + " Test Aborted!")

        Catch ex As Exception
            MyBase.HandleTestFailException(ex)
        End Try
    End Sub



    ' Labview VI Active X function
    'Public Overrides Sub Run(ByVal sender As Object, ByVal e As DoWorkEventArgs)
    '    Dim thread As BackgroundWorker = CType(sender, BackgroundWorker)
    '    Dim gridData As String()
    '    Dim params As TestParameters = _testParams
    '    Dim ports As String()
    '    Dim iStep As Long = 10  'interpolation step size in MHz
    '    Dim runtest As Boolean = True

    '    If thread.CancellationPending = True Then
    '        e.Cancel = True
    '        MyBase.OnStatusMessageEvent(params.TestName & " Test Aborted!")
    '        Exit Sub
    '    End If

    '    'MyBase.PreTest()

    '    Try
    '        If params.PortArray IsNot Nothing Then
    '            Dim delims As Char() = {","}
    '            ports = params.PortArray.Split(delims)
    '        Else
    '            Throw New Exception("Error: Test ports are not defined in config file")
    '        End If

    '        ' open input/output (fixture A and Fixture B) S2P files to be de-embedded
    '        Dim FA2p As S2pFile = New S2pFile(params.InputCalFileName)
    '        FA2p.UnwrapS21P()
    '        FA2p.GetInterpolatedData(params.BandStartFreq, params.BandStopFreq, iStep)

    '        Dim FB2p As S2pFile = New S2pFile(params.OutputCalFileName)
    '        FB2p.UnwrapS21P()
    '        FB2p.GetInterpolatedData(params.BandStartFreq, params.BandStopFreq, iStep)

    '        Debug.Print("FA2P File : " + params.InputCalFileName)
    '        Debug.Print("FB2P File : " + params.OutputCalFileName)
    '        Debug.Print("Band Start Freq : " + params.BandStartFreq.ToString("0MHz"))
    '        Debug.Print("Band Stop Freq : " + params.BandStopFreq.ToString("0MHz"))


    '        ' start LabView Active X server
    '        Dim labVIEWApp As Application = New Application()

    '        If labVIEWApp Is Nothing Then
    '            MyBase.OnStatusMessageEvent("RunLabVIEWVI : An error occurred getting labVIEWApp.")
    '            Throw New Exception()
    '        End If

    '        'labVIEWApp.AutomaticClose = True

    '        'Dim viPath As String = "C:\TestDevelopment\Labview\Generic_RF_Measurements\S_Parameters\RecallRegister_to_S2P_CMT304.vi"
    '        Dim vi As VirtualInstrument = labVIEWApp.GetVIReference(params.VIpath, "", False, 0)

    '        If labVIEWApp Is Nothing Then
    '            MyBase.OnStatusMessageEvent("RunLabVIEWVI : An error occurred getting the VI reference. Make sure VI path is correct.")
    '            Throw New Exception()
    '        End If


    '        ' loop thru the ports specified in config file
    '        'For port As Integer = portStart To portStop
    '        For Each port As String In ports

    '            While runtest
    '                MyBase.PreTest()
    '                params.Ports = port
    '                Dim prompt1 As String = "Connect " & params.FileNameSuffix & port
    '                If Not PromptUser(prompt1, MyBase.TestName, True) Then Throw New Exception("User aborted test")

    '                Dim s2pPath As String = OutputDataPath & "SN" & DutBarcode & "\S2P\"
    '                Dim s2pFileName As String = s2pPath & params.FileNamePrefix & DutBarcode & params.FileNameSuffix & port & ".s2p"

    '                Dim term1 As Object = "Configuration Characteristics"
    '                Dim term2 As Object = "Channel"
    '                Dim term3 As Object = "PowerOffset(dB)"
    '                'Dim term4 As Object = "Trace Characteristics"
    '                Dim term5 As Object = "S2P Data"
    '                Dim term6 As Object = "Fstart"
    '                Dim term7 As Object = "Fstop"
    '                Dim term8 As Object = "Test Duration"
    '                Dim term9 As Object = "VNA - Error Out"


    '                ' (Register Name, FileType(*.sta=0), Action (saveS2P=1), Path, Trigger Mode, Trigger Count, Trigger Point)    
    '                Dim in1 As Object = {params.stateName, 0, 1, s2pFileName, 3, 1, 0}    '
    '                Dim in2 As Object = 1
    '                Dim in3 As Object = 0
    '                'Dim in4 As Object = {"S11", "Magnitude", "S11", "Phase", "S21", "Magnitude", "S21", "Phase", "S12", "Magnitude", "S12", "Phase", "S22", "Magnitude", "S22", "Phase"}
    '                Dim out1 As Object = 0
    '                Dim out2 As Object = 0
    '                Dim out3 As Object = 0
    '                Dim out4 As Object = 0
    '                Dim errorOut As Object = 0

    '                Dim paramNames As Object = {term1, term2, term3, term5, term6, term7, term8, term9}
    '                Dim paramValues As Object = {in1, in2, in3, out1, out2, out3, out4, errorOut}

    '                ' run the vi
    '                vi.Call2(paramNames, paramValues, False, False, False, True)

    '                ' returned values from vi
    '                'Dim fStart As Long = paramValues(4)
    '                'Dim fStop As Long = paramValues(5)
    '                'Dim testTime As Double = Math.Round(paramValues(6), 1)

    '                ' vi error handling
    '                Dim errStatus As Boolean = CBool(paramValues(7)(0))
    '                Dim errCode As Integer = CInt(paramValues(7)(1))
    '                Dim errSource As String = CStr(paramValues(7)(2))

    '                If errStatus Then Throw New Exception(params.TestName & ": Code:" & errCode & " Source: " & errSource)

    '                ' instantiate S2P calss and add returned values  
    '                Dim s2pMeas As S2pFile = New S2pFile(CType(paramValues(3), Double(,)))
    '                'Dim s2pMeas As S2pFile = New S2pFile("c:\matlab\mbrd\8329870006\mbrd-e1-8329870006_BDSF_C1_J1.s2p") ' temporary - pull measured s2p file for validation
    '                s2pMeas.UnwrapS21P()
    '                s2pMeas.GetInterpolatedData(params.BandStartFreq, params.BandStopFreq, iStep)

    '                'de-Embed S21 Magnitude and Phase
    '                s2pMeas.SubtractS21(FA2p, params.InputCalCorrFactor)
    '                s2pMeas.SubtractS21(FB2p, params.OutputCalCorrFactor)

    '                ' factor in offset(dB) if required (BDSF, BDSR -3.3dB)
    '                s2pMeas.AddS21(params.OffsetS21M)

    '                ' round all values to n places
    '                s2pMeas.RoundToDecimal(2)

    '                'Dim stats As TraceStatsRecord = New TraceStatsRecord()
    '                Dim stats As TraceStatsRecord = s2pMeas.GetTraceStats()

    '                ' write to status window
    '                MyBase.OnStatusMessageEvent(params.FileNameSuffix & port & " In-Band Stats : ")
    '                MyBase.OnStatusMessageEvent("Mean : " & stats.MeanValue.ToString("0.00dB (") & (stats.StartFreq / 1000000.0).ToString("0 to ") & (stats.StopFreq / 1000000.0).ToString("0MHz)"))
    '                MyBase.OnStatusMessageEvent("Min : " & stats.MinValue.ToString("0.00dB @ ") & (stats.MinFreq / 1000000.0).ToString("0MHz"))
    '                MyBase.OnStatusMessageEvent("Max : " & stats.MaxValue.ToString("0.00dB @ ") & (stats.MaxFreq / 1000000.0).ToString("0MHz"))

    '                ' write data to grid
    '                Dim outHeader As String = params.FileNameSuffix & port & " S21M  Mean"
    '                Dim measData As String = stats.MeanValue.ToString("0.00")
    '                gridData = {outHeader, params.LL.ToString, measData, params.UL.ToString, params.Units}
    '                Dim testResult As TestStatus = GetTestResult(stats.MeanValue, params.LL, params.UL)
    '                OnAddDataToGridEvent(gridData, testResult)


    '                ' add additional file header information
    '                SqlTable = params.TestName

    '                ' save to xml file
    '                If OutputFileFormat.ToUpper = "XML" Then
    '                    Dim xmlPath As String = OutputDataPath & "SN" & DutBarcode & "\XML\"
    '                    If Not Directory.Exists(xmlPath) Then Directory.CreateDirectory(xmlPath)

    '                    Dim xmlFileName As String = xmlPath & params.FileNamePrefix & DutBarcode & params.FileNameSuffix & port & ".xml"
    '                    Dim xmlData As List(Of S2P_Record) = s2pMeas.S2P.Values.ToList()

    '                    WriteListToXMLFile(Of S2P_Record)(xmlFileName, GetFileComments().ToArray(), xmlData)
    '                    MyBase.OnStatusMessageEvent("Test Data File Saved As : " & xmlFileName)
    '                End If

    '                ' save to csv file
    '                If OutputFileFormat.ToUpper = "CSV" Then
    '                    Dim csvPath As String = OutputDataPath & "SN" & DutBarcode & "\CSV\"
    '                    If Not Directory.Exists(csvPath) Then Directory.CreateDirectory(csvPath)

    '                    Dim sqlTableName As String = params.SQLlogFileName ' "MBRD_TestLog"
    '                    Dim ticks As String = _startTime.Ticks.ToString("_0")
    '                    Dim csvFileName As String = csvPath & sqlTableName & "_" & DutBarcode & params.FileNameSuffix & port & ticks & ".csv"
    '                    Dim textData As List(Of String()) = BuildCsvLogFile(sqlTableName, stats, testResult)

    '                    WriteListToTextFile(csvFileName, textData, ",")
    '                    MyBase.OnStatusMessageEvent("Test Log File Saved As : " & csvFileName)

    '                    sqlTableName = params.SQLtableName  ' "MBRD_S2P"
    '                    csvFileName = csvPath & sqlTableName & "_" & DutBarcode & params.FileNameSuffix & port & ticks & ".csv"
    '                    textData = BuildCsvS2PFile(sqlTableName, s2pMeas)
    '                    WriteListToTextFile(csvFileName, textData, ",")
    '                    'MyBase.OnStatusMessageEvent("Test Data File Saved As : " & csvFileName)
    '                End If

    '                'MyBase.OnStatusMessageEvent("S2P File Saved As : " & s2pFileName)
    '                MyBase.OnStatusMessageEvent("")

    '                ' prompt user to retest if fail - not sure how to do this yet
    '                If testResult = TestStatus.Fail Then
    '                    If Not PromptUserToRetry() Then runtest = False
    '                Else
    '                    runtest = False
    '                End If

    '                ' check if Abort button pressed
    '                If thread.CancellationPending = True Then
    '                    e.Cancel = True
    '                    Throw New Exception(params.TestName & " Test Aborted!")
    '                End If
    '            End While

    '            runtest = True
    '        Next

    '        MyBase.PostTest()

    '    Catch ex As Exception
    '        MyBase.HandleTestFailException(ex)
    '    End Try
    'End Sub






#Region "bonepile"
    'WriteListToXMLFile(Of Double())(dataFileName, GetBaseComments(params).ToArray(), s2pMeas.S2pToList())

    ' check if Abort button pressed
    'If thread.CancellationPending = True Then
    '   e.Cancel = True
    '   'OnUpdateTestStatusEvent(TestNumber, TestStatus.Aborted)
    '   Throw New Exception(params.TestName & " Test Aborted!")
    'End If

#End Region

End Class
