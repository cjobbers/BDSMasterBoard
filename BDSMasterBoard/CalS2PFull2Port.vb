Imports TimeDate = Microsoft.VisualBasic.DateAndTime
Imports Instruments
Imports Instruments.Globals
Imports TestExecutiveLibrary
Imports System.ComponentModel
Imports LabVIEW
Imports System.Collections.Generic

Public Class CalS2PFull2Port
    Inherits TestsCommon

    Public Sub New(ByVal instrumentList As Dictionary(Of String, Instrument), ByVal params As CalParameters)
        MyBase.New(instrumentList, params)
    End Sub


    ' Test Executive Backgroundworker will call this sub on separate thread
    Public Overrides Sub Run(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim thread As BackgroundWorker = CType(sender, BackgroundWorker)
        Dim gridData As String()
        Dim params As CalParameters = _calParams

        If thread.CancellationPending = True Then
            e.Cancel = True
            MyBase.OnStatusMessageEvent(params.TestName & " Test Aborted!")
            Exit Sub
        End If


        Try
            Dim prompt1 As String = "Connect E-Cal between Port 1 and Port 2 VNA Cables"
            If Not PromptUser(prompt1, MyBase.TestName, True) Then Throw New Exception("User aborted test")

            MyBase.PreTest()
            VNA.RecallReg(params.CalRecallStateName)
            VNA.Calibrate(vnaCal.Full_2Port, vnaStandard.ECal)
            VNA.SaveReg(params.CalSaveStateName)

            gridData = {"Calibrate Full 2-Port", "", "1", "", "Bool"}

            ' check calibration values. Wrap limits around S21M min
            'VNA.RecallReg(params.CalSaveStateName)
            'VNA.Measure = vnaMeasure.S21
            'VNA.Format = vnaFormat.LogMag
            'VNA.Trigger = vnaTrigger.SingleSweep

            'read marker stats for S11 and S21?

            ' write data to grid
            Dim outHeader As String = "S21M  Minimum across band"
            'Dim measData As String = stats.MinValue.ToString("0.00")
            Dim measData As String = "0"
            gridData = {outHeader, params.LL.ToString, measData, params.UL.ToString, params.Units}
            Dim testResult As TestStatus = GetTestResult(0, params.LL, params.UL)
            OnAddDataToGridEvent(gridData, testResult)


            ' check if Abort button pressed
            If thread.CancellationPending = True Then
                e.Cancel = True
                Throw New Exception(params.TestName & " Test Aborted!")
            End If

            MyBase.PostTest()

        Catch ex As Exception
            MyBase.HandleFatalErrorException(ex)
        End Try
    End Sub


    ' Labview
    'Public Overrides Sub Run(ByVal sender As Object, ByVal e As DoWorkEventArgs)
    '    Dim thread As BackgroundWorker = CType(sender, BackgroundWorker)
    '    Dim gridData As String()
    '    Dim params As CalParameters = _calParams


    '    If thread.CancellationPending = True Then
    '        e.Cancel = True
    '        MyBase.OnStatusMessageEvent(params.TestName & " Test Aborted!")
    '        Exit Sub
    '    End If

    '    MyBase.PreTest()

    '    Try
    '        ' start LabView Active X server
    '        Dim labVIEWApp As Application = New Application()

    '        If labVIEWApp Is Nothing Then
    '            MyBase.OnStatusMessageEvent("RunLabVIEWVI : An error occurred getting labVIEWApp.")
    '            Throw New Exception()
    '        End If

    '        'Dim viPath As String = "C:\TestDevelopment\Labview\Generic_RF_Measurements\S_Parameters\Full2PortCal_CMT304.vi"
    '        Dim vi As VirtualInstrument = labVIEWApp.GetVIReference(params.VIpath, "", False, 0)

    '        If labVIEWApp Is Nothing Then
    '            MyBase.OnStatusMessageEvent("RunLabVIEWVI : An error occurred getting the VI reference. Make sure VI path is correct.")
    '            Throw New Exception()
    '        End If


    '        Dim prompt1 As String = "Connect E-Cal between Port 1 and Port 2 VNA Cables"
    '        If Not PromptUser(prompt1, MyBase.TestName, True) Then Throw New Exception("User aborted test")


    '        Dim term1 As Object = "Recall Register Name"
    '        Dim term2 As Object = "Save Register Name"
    '        Dim term3 As Object = "Stimulus Port"
    '        Dim term4 As Object = "Response Port"
    '        Dim term5 As Object = "error out"


    '        Dim in1 As Object = "State09.sta"
    '        Dim in2 As Object = "State01.sta"
    '        Dim in3 As Object = 1
    '        Dim in4 As Object = 2
    '        Dim errorOut As Object = 0


    '        Dim paramNames As Object = {term1, term2, term3, term4, term5}
    '        Dim paramValues As Object = {in1, in2, in3, in4, errorOut}

    '        ' run the vi
    '        vi.Call2(paramNames, paramValues, False, False, False, True)

    '        ' vi error handling
    '        Dim errStatus As Boolean = CBool(paramValues(4)(0))
    '        Dim errCode As Integer = CInt(paramValues(4)(1))
    '        Dim errSource As String = CStr(paramValues(4)(2))

    '        gridData = {"Calibrate Full 2-Port", "", Not errStatus, "", "Bool"}

    '        If Not errStatus Then
    '            OnAddDataToGridEvent(gridData, TestStatus.Pass)
    '        Else
    '            OnAddDataToGridEvent(gridData, TestStatus.FatalError)
    '            Throw New Exception(params.TestName & ": Code:" & errCode & " Source: " & errSource)
    '        End If

    '        ' check calibration values. Wrap limits around S21M min

    '        ' write data to grid
    '        Dim outHeader As String = "S21M  Minimum across band"
    '        'Dim measData As String = stats.MinValue.ToString("0.00")
    '        Dim measData As String = "0"
    '        gridData = {outHeader, params.LL.ToString, measData, params.UL.ToString, params.Units}
    '        Dim testResult As TestStatus = GetTestResult(0, params.LL, params.UL)
    '        OnAddDataToGridEvent(gridData, testResult)


    '        ' check if Abort button pressed
    '        If thread.CancellationPending = True Then
    '            e.Cancel = True
    '            Throw New Exception(params.TestName & " Test Aborted!")
    '        End If

    '        MyBase.PostTest()

    '    Catch ex As Exception
    '        MyBase.HandleTestFailException(ex)
    '    End Try
    'End Sub
End Class
