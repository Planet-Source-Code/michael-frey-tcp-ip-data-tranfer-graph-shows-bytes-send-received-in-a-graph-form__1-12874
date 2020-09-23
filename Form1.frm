VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "TCP/IP transfer graph"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3195
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3195
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox P2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      ForeColor       =   &H000000FF&
      Height          =   1110
      Left            =   45
      ScaleHeight     =   1050
      ScaleWidth      =   3060
      TabIndex        =   1
      Top             =   1170
      Width           =   3120
   End
   Begin VB.PictureBox P1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      ForeColor       =   &H0000FFFF&
      Height          =   1110
      Left            =   45
      ScaleHeight     =   1050
      ScaleWidth      =   3060
      TabIndex        =   0
      Top             =   45
      Width           =   3120
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3015
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arr1(130) As Integer
Dim Arr2(130) As Integer
Dim RecVal, SenVal As Integer
Dim FirstRecVal, FirstSenVal As Integer

Private Sub Form_Load()
Call StayOnTop(Form1)

Me.Left = Screen.Width - Me.Width - 50
Me.Top = 50
Er = GetTcpStatistics(TCPSTATISTICS)
    If Er <> 0 Then
        Exit Sub
    End If
FirstRecVal = TCPSTATISTICS.SegmentsReceived
FirstSenVal = TCPSTATISTICS.SegmentsSent
End Sub

Private Sub Timer1_Timer()
Dim Dif1 As Integer
Dim Dif2 As Integer

Er = GetTcpStatistics(TCPSTATISTICS)
    If Er <> 0 Then
        Exit Sub
    End If
P1.Cls
P2.Cls
P1.Print "Bytes received"
P2.Print "Bytes send"
RecVal = TCPSTATISTICS.SegmentsReceived
SenVal = TCPSTATISTICS.SegmentsSent
Dif1 = RecVal - FirstRecVal
Dif2 = SenVal - FirstSenVal
Dif1 = Dif1 * 70    'Make this value lower if you have a
                    'faster connection or higher if you
Dif2 = Dif2 * 70    'have a slower connection.
ResortArray Dif1, Dif2
    For z = 0 To 130
        P1.Line (P1.Width - (z * 30), P1.Height - Arr1(z))-(P1.Width - (z * 30), P1.Height)
        P2.Line (P2.Width - (z * 30), P2.Height - Arr2(z))-(P2.Width - (z * 30), P2.Height)
    Next z
FirstRecVal = RecVal
FirstSenVal = SenVal
End Sub

Private Sub ResortArray(NewVal As Integer, NewVal2 As Integer)
    For z = 129 To 0 Step -1
        Arr1(z + 1) = Arr1(z)
        Arr2(z + 1) = Arr2(z)
    Next z
Arr1(0) = NewVal
Arr2(0) = NewVal2
End Sub


