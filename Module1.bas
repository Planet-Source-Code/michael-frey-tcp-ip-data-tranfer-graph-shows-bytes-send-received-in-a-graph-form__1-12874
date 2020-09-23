Attribute VB_Name = "Module1"
Public TCPSTATISTICS As TCPSTATISTICS
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTcpStatistics Lib "iphlpapi.dll" (pStats As TCPSTATISTICS) As Long

Public Type TCPSTATISTICS
    AlgorithmTimeout As Long
    MinTimeout As Long
    MaxTimeout As Long
    MaxConnections As Long
    ActiveOpenConnections As Long
    PassiveOpenConnections As Long
    FailedAttempts As Long
    EstablishedConnectionResets As Long
    CurrentEstablishedConnections As Long
    SegmentsReceived As Long
    SegmentsSent As Long
    SegmentsRetransmitted As Long
    IncomingErrors As Long
    OutgoingResets As Long
    CumulativeConnections As Long
End Type

Public Sub StayOnTop(Frm As Form)
    setontop = SetWindowPos(Frm.hwnd, -1, 0, 0, 0, 0, FLAGS)
End Sub

