Imports System.Xml
Imports System.Xml.Serialization
Imports Instruments
Imports Instruments.Globals
Imports System.Reflection
Imports TestExecutiveLibrary


Public Class TestsCollection
    Inherits Tests

    Private _instrumentList As Dictionary(Of String, Instrument)


    Public Sub New()

    End Sub


    Public Overrides Sub Initialize(ByRef args() As Object)

        _instrumentList = CType(args(0), Dictionary(Of String, Instrument))
        _projectRecord = CType(args(1), ProjectRecord)
        _testList = New Dictionary(Of Integer, Test)

        Dim testFile As String = _projectRecord.TestParametersFileName
        Dim calFile As String = _projectRecord.CalParametersFileName
        Dim testList As List(Of TestParameters) = ReadXMLFileToList(Of TestParameters)(testFile, "")


        For Each rec As TestParameters In testList
            Dim test As Test
            Dim testName As String = rec.TestName
            Dim err As String = rec.TestName & " failed to load from " & testFile
            Dim testArgs As Object() = {_instrumentList, rec}

            Try
                'Assembly.GetExecutingAssembly().GetModules()
                test = CType(Assembly.GetExecutingAssembly().CreateInstance(testName, False, BindingFlags.Default, Nothing, testArgs, Nothing, Nothing), Test)
                If test Is Nothing Then Throw New Exception(err)

                test.ConfigPath = _projectRecord.ConfigPath
                test.CalPath = _projectRecord.CalPath
                test.OutputFileFormat = _projectRecord.OutputFileFormat
                test.TestLocation = _projectRecord.TestLocation
                test.TestStation = _projectRecord.TestStation
                test.TestFixture = _projectRecord.TestFixture
                test.TestExecRev = _projectRecord.TestExecRev
                test.TestLimitsRev = _projectRecord.TestLimitsRev
                test.TestPlanRev = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString()

                _testList.Add(rec.TestNumber, test)
            Catch ex As Exception
                test = Nothing
                MyBase.OnStatusMessageEvent(ex.Message)
            End Try
        Next


        Dim calList As List(Of CalParameters) = ReadXMLFileToList(Of CalParameters)(calFile, "")
        _calibrationList = New Dictionary(Of String, Test)

        For Each rec As CalParameters In calList
            Dim test As Test
            Dim testName As String = rec.TestName
            Dim err As String = rec.TestName & " failed to load from " & calFile
            Dim calArgs As Object() = {_instrumentList, rec}

            Try
                test = CType(Assembly.GetExecutingAssembly().CreateInstance(testName, False, BindingFlags.Default, Nothing, calArgs, Nothing, Nothing), Test)
                If test Is Nothing Then Throw New Exception(err)

                test.ConfigPath = _projectRecord.ConfigPath
                test.CalPath = _projectRecord.CalPath

                _calibrationList.Add(rec.TestDescription, test)
            Catch ex As Exception
                test = Nothing
                MyBase.OnStatusMessageEvent(ex.Message)
            End Try
        Next
    End Sub

    ' populate GUI Tools menu with these items - must match Sub name exactly which will be called
    Public Overrides Function GetToolsList() As List(Of String)
        Dim list As New List(Of String)
        'list.Add("OpenS2pFile")
        ' list.Add("WriteDefaultTestInfoToFile")
        ' list.Add("WriteDefaultFreqListToFile")

        Return list
    End Function


    'Public Sub OpenS2pFile()
    '    Try
    '        Dim file1 As S2pFile = New S2pFile("C:\test\mtbrd.s2p")
    '        Dim content As Dictionary(Of Long, S2P_Record) = file1.S2P

    '        MyBase.OnStatusMessageEvent(content.Item(0).Freq.ToString("Freq: 0MHz"))
    '    Catch ex As Exception
    '        MyBase.OnStatusMessageEvent(ex.Message)
    '    End Try
    'End Sub


    'Public Sub PowerSuppliesOn()
    '    Try
    '        Dim rec As New TestParameters
    '        Dim clsInstance As New PowerUpUnit(_instrumentList, rec)
    '        clsInstance.PowerUpUnit()
    '        MyBase.OnStatusMessageEvent("Power Supplies Turned On")
    '    Catch ex As Exception
    '        MyBase.OnStatusMessageEvent(ex.Message)
    '    End Try
    'End Sub

    'Public Sub PowerSuppliesOff()
    '    Try
    '        Dim rec As New TestParameters
    '        Dim clsInstance As New PowerOffUnit(_instrumentList, rec)
    '        clsInstance.PowerOffUnit()
    '        MyBase.OnStatusMessageEvent("Power Supplies Turned Off")
    '    Catch ex As Exception
    '        MyBase.OnStatusMessageEvent(ex.Message)
    '    End Try
    'End Sub




    'Public Function ReadTestsFromXmlFile(ByVal fileName As String) As Dictionary(Of Integer, TestParametersDCV014G2)
    '    Try
    '        Dim dict As New Dictionary(Of Integer, TestParameters)
    '        Dim list As List(Of TestParameters) = ReadXMLFileToList(Of TestParameters)(fileName, "")

    '        For Each item As TestParameters In list
    '            dict.Add(item.TestNumber, item)
    '        Next

    '        Return dict
    '    Catch ex As Exception
    '        Throw New Exception("Error retrieving Test Parameters File" & fileName)
    '    End Try
    'End Function


    'Public Function ReadCalInfoFromXmlFile(ByVal fileName As String) As Dictionary(Of Integer, CalParametersDCV014G2)
    '    Try
    '        Dim dict As New Dictionary(Of Integer, CalParameters)
    '        Dim list As List(Of CalParameters) = ReadXMLFileToList(Of CalParameters)(fileName, "")

    '        For Each item In list
    '            dict.Add(item.TestNumber, item)
    '        Next

    '        Return dict
    '    Catch ex As Exception
    '        Throw New Exception("Error retrieving Calibration Parameters File" & fileName)
    '    End Try
    'End Function


End Class
