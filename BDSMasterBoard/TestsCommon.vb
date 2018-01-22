Imports TimeDate = Microsoft.VisualBasic.DateAndTime
Imports System.ComponentModel
Imports Instruments
Imports Instruments.Globals
Imports System.Reflection
Imports TestExecutiveLibrary
Imports System.IO

' Test functions to be inherited by Test Classes
Public MustInherit Class TestsCommon
    Inherits Test

    Protected _startTime As DateTime = DateTime.Now
    Protected tp As TestParameters        ' _testParams
    Protected _calParams As CalParameters
    Protected testResult As TestStatus = TestStatus.NotTested

    ' test instruments allocation
    Protected WithEvents VNA As IVNA
    Protected WithEvents TTP As BdsTestTilePlus
    Protected WithEvents DFETB As BdsDfeTestBoard
    Protected WithEvents DFETB2 As BdsDfeTestBoard
    Protected WithEvents DFETB03 As BdsDfeTestBoard03

    Protected Sub New(ByVal instrumentList As Dictionary(Of String, Instrument))

    End Sub

    Public Sub New(ByVal instrumentList As Dictionary(Of String, Instrument), ByVal params As TestParameters)
        MyBase._instrumentList = instrumentList
        MyBase._testNum = params.TestNumber
        MyBase._testName = params.TestName
        MyBase._testDescription = params.TestDescription
        MyBase._outputFileName = params.OutputFileName
        tp = params
    End Sub

    Public Sub New(ByVal instrumentList As Dictionary(Of String, Instrument), ByVal params As CalParameters)
        MyBase._instrumentList = instrumentList
        MyBase._testNum = params.TestNumber
        MyBase._testName = params.TestName
        MyBase._testDescription = params.TestDescription
        MyBase._outputFileName = params.OutputFileName
        _calParams = params
    End Sub


    Protected Overrides Sub AllocateInstruments(ByVal instrumentList As Dictionary(Of String, Instrument))

        'Dim instrument As KeyValuePair(Of String, Instrument)
        'Dim assy As Assembly = Assembly.GetAssembly(GetType(Instrument))
        'Dim thisAssyName As String = Assembly.GetExecutingAssembly().GetName().Name

        ' must be a better way...
        'myVTM1 = CType(Assembly.GetExecutingAssembly().CreateInstance(instrument.Key), IMultimeter)

        Try
            VNA = CType(instrumentList("VNA"), IVNA)
            'VNA.Preset()
        Catch ex As Exception
            Throw New Exception("AllocateInstruments : Error allocating instruments")
        End Try

    End Sub


    Protected Sub PreTest()

        _startTime = DateTime.Now
        results.Clear()

        Me.AllocateInstruments(_instrumentList)
        'Dim prompt1 As String = "Connect Board Under Test "
        MyBase.OnStatusMessageEvent(MyBase.OutputFileName & " Test Started...")
        MyBase.OnStatusMessageEvent("")
        MyBase.OnUpdateTestStatusEvent(MyBase.TestNumber, TestStatus.Running)
        'If Not PromptUser(prompt1, MyBase.TestName, False) Then Throw New Exception("User aborted test")

    End Sub


    Protected Sub PostTest()
        MyBase.OnUpdateTestStatusEvent(MyBase.TestNumber, TestStatus.Pass)
        testResult = GetTestResultSummary()
        If testResult <> TestStatus.Pass Then MyBase.OnUpdateTestStatusEvent(MyBase.TestNumber, TestStatus.Fail)

        MyBase.OnStatusMessageEvent(MyBase.TestName & " Complete. Total test time = " & (DateTime.Now.Subtract(_startTime).TotalSeconds).ToString("0.00 seconds"))
        MyBase.OnStatusMessageEvent("")

        Try
            'mySource1.RfOn = False
            'mySpecAn.AutoAlign = state.TurnOn
        Catch ex As Exception
            MyBase.OnStatusMessageEvent("Error : Post-Test")
        End Try
    End Sub


    Protected Function PromptUserToRetry() As Boolean
        Dim resp As MsgBoxResult

        resp = MsgBox("Test Failed. Would you like to retry? ", MsgBoxStyle.RetryCancel, MyBase.TestName)
        If resp = vbCancel Then
            'OnUpdateTestStatusEvent(MyBase.TestNumber, TestStatus.FatalError)
            'OnStatusMessageEvent("Test Aborted!")
            'MyBase.OnStatusMessageEvent("")
            Return False
        End If

        Return True
    End Function


    Protected Sub HandleFatalErrorException(ByRef ex As Exception)
        testResult = TestStatus.FatalError
        MyBase.OnUpdateTestStatusEvent(MyBase.TestNumber, testResult)
        MyBase.OnStatusMessageEvent(MyBase.TestName & " Fatal Error Occurred!")
        MyBase.OnStatusMessageEvent(ex.Message)
        MyBase.OnStatusMessageEvent("")
    End Sub


    Protected Sub HandleTestFailException(ByRef ex As Exception)
        testResult = TestStatus.Fail
        MyBase.OnUpdateTestStatusEvent(MyBase.TestNumber, testResult)
        MyBase.OnStatusMessageEvent(MyBase.TestName & " Failed!")
        MyBase.OnStatusMessageEvent(ex.Message)
        MyBase.OnStatusMessageEvent("")
    End Sub

    Protected Sub HandleAbortTestException(ByRef ex As Exception)
        testResult = TestStatus.Aborted
        MyBase.OnUpdateTestStatusEvent(MyBase.TestNumber, testResult)
        MyBase.OnStatusMessageEvent(MyBase.TestName & " Test Aborted!")
        MyBase.OnStatusMessageEvent(ex.Message)
        MyBase.OnStatusMessageEvent("")
    End Sub


    Protected Function DataToGrid(testName As String, measurement As Double, LL As Double, UL As Double, units As String) As TestStatus
        Try
            Dim status As TestStatus = Me.GetTestResult(measurement, LL, UL)
            Dim gridData As String() = {testName, LL.ToString(), measurement.ToString(), UL.ToString(), units}
            MyBase.OnAddDataToGridEvent(gridData, status)
            ' MyBase.OnStatusMessageEvent(testName + " = " + measurement.ToString("0.00") + units)
            results.Add(status)

            Return status

        Catch ex As Exception
            MyBase.OnStatusMessageEvent("Error writing data to grid : " + ex.Message)
            Throw
        End Try
    End Function


    Public Function GetFileComments() As List(Of String)

        Dim header As List(Of String) = New List(Of String)
        header.Add("SqlTable:" & tp.SQLtableName)
        header.Add("Barcode:" & DutBarcode)
        header.Add("TestName:" & tp.TestName)
        header.Add("TestDescription:" & tp.TestDescription)
        header.Add("ProcessStep:" & ProcessStep)
        header.Add("TestOperator:" & TestOperator)
        header.Add("TestLocation:" & TestLocation)
        header.Add("TestStation:" & TestStation)
        header.Add("TestFixture:" & TestFixture)
        header.Add("TestExecRev:" & TestExecRev)
        header.Add("TestPlanRev:" & TestPlanRev)
        header.Add("TestLimitsRev:" & TestLimitsRev)
        header.Add("LastCalDate:" & LastCalDate)
        header.Add("HwRev:" & tp.FileNamePrefix)
        header.Add("Config:" & tp.FileNameSuffix)
        header.Add("Port:" & tp.Port.ToString())
        header.Add("Temp:" & tp.Temp)
        header.Add("InputCalFile:" & tp.InputCalFileName)
        header.Add("OutputCalFile:" & tp.OutputCalFileName)

        Return header
    End Function


    Public Sub WriteXmlS2PFile(ByRef s2pMeas As S2pFile)
        Try
            If OutputFileFormat.ToUpper <> "XML" Then Exit Sub

            Dim xmlPath As String = OutputDataPath & "SN" & DutBarcode & "\XML\"
            Dim xmlFileName As String = xmlPath & tp.FileNamePrefix & DutBarcode & tp.FileNameSuffix & tp.Port & ".xml"

            If Not Directory.Exists(xmlPath) Then Directory.CreateDirectory(xmlPath)

            'MyBase.SqlTable = tp.TestName
            Dim xmlData As List(Of S2P_Record) = s2pMeas.S2P.Values.ToList()

            WriteListToXMLFile(Of S2P_Record)(xmlFileName, GetFileComments().ToArray(), xmlData)
            MyBase.OnStatusMessageEvent("Test Data File Saved As : " & xmlFileName)
        Catch ex As Exception
            Throw New Exception("Error : S2P.WriteXmlS2PFile : " + ex.Message)
        End Try
    End Sub


    Public Sub WriteCsvS2PFile(ByRef s2pMeas As S2pFile)
        'Public Function BuildCsvS2PFile(ByRef s2pMeas As S2pFile) As List(Of String())
        Try
            If OutputFileFormat.ToUpper <> "CSV" Then Exit Sub

            Dim data As Dictionary(Of Long, S2P_Record) = s2pMeas.S2P
            Dim strArray As List(Of String()) = New List(Of String())
            Dim delims As Char() = {" ", "_", ","}
            Dim config As String() = tp.FileNameSuffix.Split(delims)   '_CAL_C1_

            ' header information (10 lines)
            strArray.Add({"SQLTable:", tp.SQLtableName, "TestName:", tp.TestName})
            strArray.Add({"Comments:", "First 10 lines are ignored by DB Import"})
            strArray.Add({"Barcode:", DutBarcode, "TestDescription:", TestDescription})
            strArray.Add({"ProcessStep:", ProcessStep, "TestOperator:", TestOperator, "InputCalFile:", tp.InputCalFileName, "OutputCalFile:", tp.OutputCalFileName})
            strArray.Add({"TestLocation:", TestLocation, "TestStation:", TestStation, "TestFixture:", TestFixture})
            strArray.Add({"TestExecRev:", TestExecRev, "TestPlanRev:", TestPlanRev, "TestLimitsRev:", TestLimitsRev})
            strArray.Add({"HwRev:", tp.FileNamePrefix, "LastCalDate:", LastCalDate})
            strArray.Add({"Config:", tp.FileNameSuffix, "Port:", tp.Port, "Temp:", tp.Temp})
            strArray.Add({"BRD_Id", "BRD_SN", "BRD_StartTime", "Channel", "TxRxPath", "JPort", "Freq_Mhz", "S11M", "S11P", "S21M", "S21P", "S12M", "S12P", "S22M", "S22P"})
            strArray.Add({"INT", "VARCHAR(32)", "DATETIME", "VARCHAR(8)", "VARCHAR(8)", "VARCHAR(8)", "DEC(9,3)", "DEC(6,2)", "DEC(6,2)", "DEC(6,2)", "DEC(6,2)", "DEC(6,2)", "DEC(6,2)", "DEC(6,2)", "DEC(6,2)"})

            For Each freq As Long In data.Keys
                Dim values As List(Of String) = New List(Of String)
                values.Add("0")                'Auto Increment     
                values.Add(DutBarcode)
                values.Add(_startTime.ToString("yyyy-MM-dd HH:mm:ss"))
                If config.Length > 2 Then values.Add(config(2))           'C1 or C2
                If config.Length > 1 Then values.Add(config(1))           'CAL, RX1, RX2..
                values.Add(tp.Port)   'J1, J2..

                values.Add((data(freq).Freq / 1000000.0).ToString)
                values.Add(data(freq).S11M.ToString())
                values.Add(data(freq).S11P.ToString())
                values.Add(data(freq).S21M.ToString())
                values.Add(data(freq).S21P.ToString())
                values.Add(data(freq).S12M.ToString())
                values.Add(data(freq).S12P.ToString())
                values.Add(data(freq).S22M.ToString())
                values.Add(data(freq).S22P.ToString())

                strArray.Add(values.ToArray)
            Next

            'Return strArray

            Dim csvPath As String = OutputDataPath & "SN" & DutBarcode & "\CSV\"
            Dim csvFileName As String = csvPath & tp.SQLtableName & "_" & DutBarcode & tp.FileNameSuffix & tp.Port & _startTime.Ticks.ToString("_0") & ".csv"
            If Not Directory.Exists(csvPath) Then Directory.CreateDirectory(csvPath)

            WriteListToTextFile(csvFileName, strArray, ",")
            'MyBase.OnStatusMessageEvent("Test Data File Saved As : " & csvFileName)

        Catch ex As Exception
            Throw New Exception("Error : S2P.WriteCsvS2PFile : " + ex.Message)
        End Try
    End Sub


    Public Sub WriteCsvLogFile()
        'Public Function BuildCsvLogFile(ByRef stats As TraceStatsRecord, ByVal testResult As TestStatus) As List(Of String())
        Try
            If OutputFileFormat.ToUpper <> "CSV" Then Exit Sub

            Dim strArray As List(Of String()) = New List(Of String())
            Dim delims As Char() = {" ", "_", ","}
            Dim config As String() = tp.FileNameSuffix.Split(delims)   '_CAL_C1_

            ' header information (10 lines)
            strArray.Add({"SQLTable:", tp.SQLlogFileName, "TestName:", tp.TestName})
            strArray.Add({"Comments:", "First 10 lines are ignored by DB Import"})
            strArray.Add({"\N"})
            strArray.Add({"InputCalFile:", tp.InputCalFileName, "OutputCalFile:", tp.OutputCalFileName})
            strArray.Add({"Temp:", tp.Temp})
            strArray.Add({"\N"})
            strArray.Add({"\N"})
            strArray.Add({"MBRD_SN", "MBRD_StartTime", "MBRD_Product_ID", "Location", "TestStation", "Fixture", "TestExecRev", "TestPlanRev",
                         "TestLimitsRev", "Operator", "LastCalDate", "ProcessStep", "Comments", "Channel", "TxRxPath", "JPort",
                         "S21M_Min", "S21M_Mean", "S21M_Max", "S21FreqMin", "S21FreqMax", "TestResult"})
            strArray.Add({"VARCHAR(32)", "DATETIME", "VARCHAR(32)", "VARCHAR(8)", "VARCHAR(8)", "VARCHAR(32)", "VARCHAR(32)", "VARCHAR(32)",
                         "VARCHAR(32)", "VARCHAR(16)", "DATETIME", "VARCHAR(16)", "VARCHAR(96)", "VARCHAR(8)", "VARCHAR(8)", "VARCHAR(8)",
                         "DEC(6,2)", "DEC(6,2)", "DEC(6,2)", "DEC(9,2)", "DEC(9,2)", "VARCHAR(16)"})
            strArray.Add({"\N"})

            ' data
            Dim values As List(Of String) = New List(Of String)
            values.Add(DutBarcode)
            values.Add(_startTime.ToString("yyyy-MM-dd HH:mm:ss"))
            values.Add(tp.FileNamePrefix)
            values.Add(TestLocation)
            values.Add(TestStation)
            values.Add(TestFixture)
            values.Add(TestExecRev)
            values.Add(TestPlanRev)
            values.Add(TestLimitsRev)
            values.Add(TestOperator)
            values.Add(LastCalDate)
            values.Add(ProcessStep)
            values.Add(TestDescription)

            values.Add("")      'config(2)
            values.Add("")      'config(1)
            values.Add(tp.Port)
            values.Add("")     '(stats.MinValue.ToString())
            values.Add("")     '(stats.MeanValue.ToString())
            values.Add("")     '(stats.MaxValue.ToString())
            values.Add("")     '((stats.MinFreq / 1000000.0).ToString())
            values.Add("")     '((stats.MaxFreq / 1000000.0).ToString())
            values.Add(testResult.ToString())

            strArray.Add(values.ToArray)

            'Return strArray

            Dim csvPath As String = OutputDataPath & "SN" & DutBarcode & "\CSV\"
            Dim csvFileName As String = csvPath & tp.SQLlogFileName & "_" & DutBarcode & "_" & tp.Port & _startTime.Ticks.ToString("_0") & ".csv"    '& tp.FileNameSuffix

            If Not Directory.Exists(csvPath) Then Directory.CreateDirectory(csvPath)
            WriteListToTextFile(csvFileName, strArray, ",")
            MyBase.OnStatusMessageEvent("Test Log File Saved As : " & csvFileName)

        Catch ex As Exception
            MyBase.OnStatusMessageEvent("Error : S2P.WriteCsvLogFile : " + ex.Message)
        End Try
    End Sub


#Region "Bonepile"
    'Protected Function GetCalRecord(ByVal freq As Double, ByVal CalData As List(Of CalRecord)) As CalRecord

    '    Dim defRec As New CalRecord

    '    Try
    '        ' see if freq is out of bounds. If so, return max or min
    '        If CalData.First.Freq > freq Then Return CalData.First
    '        If CalData.Last.Freq < freq Then Return CalData.Last

    '        ' iterate thru list and find closest frequency
    '        For Each rec As CalRecord In CalData
    '            If rec.Freq >= freq Then Return rec
    '        Next

    '    Catch ex As Exception
    '        MyBase.OnStatusMessageEvent("No calibration found for " & freq & "MHz in file")
    '        MyBase.OnStatusMessageEvent(ex.Message)
    '        Return defRec
    '    End Try


    '    Return defRec

    'End Function


    'Protected Function GetTestResult(ByVal measurement As Double, ByVal LL As Double, ByVal UL As Double) As TestStatus
    '    If measurement >= LL And measurement <= UL Then
    '        Return TestStatus.Pass
    '    Else
    '        Return TestStatus.Fail
    '    End If
    'End Function


    'Protected Function GetTestResult(ByVal measurement As Integer, ByVal LL As Integer, ByVal UL As Integer) As TestStatus
    '    If measurement >= LL And measurement <= UL Then
    '        Return TestStatus.Pass
    '    Else
    '        Return TestStatus.Fail
    '    End If
    'End Function
#End Region


End Class
