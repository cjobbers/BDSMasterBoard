Imports System.Xml
Imports System.Xml.Serialization
'Imports Instruments.BdsDfeTestBoard
'Imports Instruments.BdsTestTilePlus
Imports Instruments


<Serializable()>
<XmlType("TestParameters")>
Public Class TestParameters
    'Public Enum Port
    '    None = 0
    '    J1 = 1
    '    J2 = 2
    '    J3 = 3
    '    J4 = 4
    '    J5 = 5
    '    J6 = 6
    '    J7 = 7
    '    J8 = 8
    '    J9 = 9
    '    J10 = 10
    '    J11 = 11
    '    J12 = 12
    '    J1thruJ6 = 13
    '    J7thruJ12 = 14
    '    J1thruJ12 = 15
    'End Enum

    <XmlAttribute()>
    Public TestNumber As Integer
    <XmlElement()>
    Public TestName As String
    <XmlElement()>
    Public TestDescription As String
    <XmlElement()>
    Public SQLlogFileName As String
    <XmlElement()>
    Public SQLtableName As String
    <XmlElement()>
    Public VIpath As String
    <XmlElement()>
    Public ProductID As String
    <XmlElement()>
    Public CalRecallStateName As String
    <XmlElement()>
    Public CalSaveStateName As String
    <XmlElement()>
    Public stateName As String
    <XmlElement()>
    Public SweepListFile As String
    <XmlElement()>
    Public FileNamePrefix As String
    <XmlElement()>
    Public FileNameSuffix As String
    <XmlElement()>
    Public OutputFileName As String
    <XmlElement()>
    Public TestGroup As Integer
    <XmlElement()>
    Public Port As String
    <XmlElement()>
    Public PortArray As String
    <XmlElement()>
    Public BandStartFreq As Double
    <XmlElement()>
    Public BandStopFreq As Double
    <XmlElement()>
    Public RFpwr As Double
    <XmlElement()>
    Public Temp As Integer

    <XmlElement()>
    Public VnaIfBw As Integer
    <XmlElement()>
    Public VnaPoints As Integer

    <XmlElement()>
    Public InputCalFileName As String
    <XmlElement()>
    Public OutputCalFileName As String
    <XmlElement()>
    Public InputCalCorrFactor As Double
    <XmlElement()>
    Public OutputCalCorrFactor As Double
    <XmlElement()>
    Public OffsetS21M As Double

    <XmlElement()>
    Public LL As Double
    <XmlElement()>
    Public UL As Double
    <XmlElement()>
    Public Units As String
End Class


<Serializable()>
<XmlType("CalParameters")>
Public Class CalParameters
    Inherits TestParameters

    <XmlElement()>
    Public RfFreqListFile As String

End Class


<Serializable()>
<XmlType("S2P_Record")>
Public Class S2P_Record
    <XmlAttribute()>
    Public Freq As Long
    <XmlElement()>
    Public S11M As Double
    <XmlElement()>
    Public S11P As Double
    <XmlElement()>
    Public S21M As Double
    <XmlElement()>
    Public S21P As Double
    <XmlElement()>
    Public S12M As Double
    <XmlElement()>
    Public S12P As Double
    <XmlElement()>
    Public S22M As Double
    <XmlElement()>
    Public S22P As Double
End Class


<Serializable()>
<XmlType("GainRecord")>
Public Class GainRecordTypeDCV
    <XmlAttribute()>
    Public RfFreq As Double
    <XmlElement()>
    Public LoFreq As Double
    <XmlElement()>
    Public IfFreq As Double
    <XmlElement()>
    Public RfPwr As Double
    <XmlElement()>
    Public LoPwr As Double
    <XmlElement()>
    Public IfPwr As Double
    <XmlElement()>
    Public IfPwrSpecAn As Double
    <XmlElement()>
    Public Gain As Double
End Class


<Serializable()>
<XmlType("Calibration")>
Public Class CalRecord
    <XmlAttribute()>
    Public Freq As Double
    <XmlElement()>
    Public Reference As Double
    <XmlElement()>
    Public CableLoss As Double
    <XmlElement()>
    Public Offset1 As Double
    <XmlElement()>
    Public Offset2 As Double
End Class


<Serializable()>
<XmlType("Measurement")>
Public Class FreqRecord
    <XmlAttribute("Frequency")>
    Public Freq As Double
End Class


<Serializable()>
<XmlType("TraceData")>
Public Class TraceDataRecord
    <XmlAttribute()>
    Public Frequency As Double
    <XmlElement()>
    Public Amplitude As Double
End Class


<Serializable()>
<XmlType("TraceStats")>
Public Class TraceStatsRecord
    <XmlAttribute()>
    Public StartFreq As Double
    <XmlElement()>
    Public StopFreq As Double
    <XmlElement()>
    Public Points As Integer
    <XmlElement()>
    Public MaxValue As Double
    <XmlElement()>
    Public MaxFreq As Double
    <XmlElement()>
    Public MinValue As Double
    <XmlElement()>
    Public MinFreq As Double
    <XmlElement()>
    Public MeanValue As Double
End Class




<Serializable()>
<XmlType("SweepParameters")>
Public Class SweepParameters
    <XmlAttribute()>
    Public TestNumber As Integer

    <XmlElement()>
    Public DfeTbSwitchPath As BdsDfeTestBoard.dfetbSwPath
    <XmlElement()>
    Public DFETB03SwitchPath As BdsDfeTestBoard03.rfSwPath
    <XmlElement()>
    Public DFETB03OutputPath As BdsDfeTestBoard03.outputSwPath
    <XmlElement()>
    Public TtpSwitchPath As BdsTestTilePlus.ttpSwPath

    <XmlElement()>
    Public FileNameSuffix As String
    <XmlElement()>
    Public BandStartFreq As Double
    <XmlElement()>
    Public BandStopFreq As Double
    <XmlElement()>
    Public InputCalFileName As String
    <XmlElement()>
    Public OutputCalFileName As String
    <XmlElement()>
    Public InputCalCorrFactor As Double
    <XmlElement()>
    Public OutputCalCorrFactor As Double
    <XmlElement()>
    Public OffsetS21M As Double

    <XmlElement()>
    Public LL As Double
    <XmlElement()>
    Public UL As Double
    <XmlElement()>
    Public Units As String

End Class