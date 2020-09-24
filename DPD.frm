VERSION 5.00
Object = "{DDE234FE-685C-11D2-811D-00600891BAB0}#1.0#0"; "DNS.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Direct Proxy Daemon Version 1.0 Non Multi User"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2760
      Top             =   4920
   End
   Begin DNSControl.DNS DNS1 
      Left            =   2400
      Top             =   4800
      _ExtentX        =   450
      _ExtentY        =   238
   End
   Begin MSWinsockLib.Winsock RemoteC 
      Left            =   1920
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LocalC 
      Left            =   1320
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8493
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Connection History"
      TabPicture(0)   =   "DPD.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Text1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Direct Proxy Daemon"
      TabPicture(1)   =   "DPD.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Check1"
      Tab(1).Control(1)=   "Command4"
      Tab(1).Control(2)=   "Command3"
      Tab(1).Control(3)=   "Command2"
      Tab(1).Control(4)=   "Command1"
      Tab(1).Control(5)=   "Rport"
      Tab(1).Control(6)=   "Radd1"
      Tab(1).Control(7)=   "Text2"
      Tab(1).Control(8)=   "RaddL"
      Tab(1).Control(9)=   "RportL"
      Tab(1).Control(10)=   "Label8"
      Tab(1).Control(11)=   "Label7"
      Tab(1).Control(12)=   "CMNL"
      Tab(1).Control(13)=   "CIPL"
      Tab(1).Control(14)=   "Label6"
      Tab(1).Control(15)=   "Label5"
      Tab(1).Control(16)=   "Label4"
      Tab(1).Control(17)=   "Label3"
      Tab(1).Control(18)=   "DPort"
      Tab(1).Control(19)=   "ServerIP1"
      Tab(1).Control(20)=   "Label2"
      Tab(1).Control(21)=   "Label1"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "Help"
      TabPicture(2)   =   "DPD.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text3"
      Tab(2).ControlCount=   1
      Begin VB.TextBox Text3 
         Height          =   4335
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Text            =   "DPD.frx":0054
         Top             =   360
         Width           =   6855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Hide Your DNS Machine Name to Clients"
         Height          =   255
         Left            =   -73080
         TabIndex        =   23
         Top             =   4080
         Width           =   3255
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Disconnect"
         Height          =   255
         Left            =   -71040
         TabIndex        =   22
         Top             =   4440
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Reset"
         Height          =   255
         Left            =   -72000
         TabIndex        =   21
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "New"
         Height          =   255
         Left            =   -72840
         TabIndex        =   20
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Set Port And Address"
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox Rport 
         Height          =   285
         Left            =   -73920
         MaxLength       =   7
         TabIndex        =   12
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox Radd1 
         Height          =   285
         Left            =   -73560
         TabIndex        =   11
         Top             =   3720
         Width           =   5535
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   360
         Width           =   6855
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   4335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   1
         Top             =   360
         Width           =   6855
      End
      Begin VB.Label RaddL 
         Height          =   255
         Left            =   -71280
         TabIndex        =   18
         Top             =   2640
         Width           =   3135
      End
      Begin VB.Label RportL 
         Height          =   255
         Left            =   -71640
         TabIndex        =   17
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Remote Port:"
         Height          =   255
         Left            =   -72600
         TabIndex        =   16
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Remote Address:"
         Height          =   255
         Left            =   -72600
         TabIndex        =   15
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label CMNL 
         Height          =   255
         Left            =   -73320
         TabIndex        =   14
         Top             =   3360
         Width           =   5295
      End
      Begin VB.Label CIPL 
         Height          =   255
         Left            =   -74160
         TabIndex        =   13
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Remote Port:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   4080
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Remote Address:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Client Machine Name:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   8
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Client IP:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   7
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label DPort 
         Height          =   255
         Left            =   -73800
         TabIndex        =   6
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label ServerIP1 
         Height          =   255
         Left            =   -74040
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Server IP:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   4
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Daemon Port:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   3
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.Label servermn1 
      Height          =   15
      Left            =   6000
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Made By Mike G, email me at metallica999@metallica.com if you have anything GOOD to say or closely related to that

Dim YesNoDNS As Boolean
Dim YesNoDone As Boolean
Dim GoDangit As Boolean
Dim serverIP As String
Dim serverMN As String
Dim CIP As String
Dim CMN As String
Dim RemoteAdd As String
Dim RemotePort1 As String
Dim DaemonPort As String
Dim RxData As String
Dim RxData2 As String


Private Sub Check1_Click()
If Check1.Value = 0 Then
    If YesNoDNS = False Then
        MsgBox "DNS Control is in use, Please try again in a few seconds."
    Else:
        serverMN = DNS1.AddressToName(LocalC.LocalIP)
        servermn1.Caption = serverMN
        serverMN = servermn1.Caption
        MsgBox "Your Machine Name is: " & serverMN
    End If
Else:
    severmn = ""
    servermn1.Caption = ""
End If
End Sub

Private Sub Command1_Click()
If Radd1.Text = "" Then
    MsgBox "Please Specify an Address."
    Exit Sub
End If
If Rport.Text = "" Then
    MsgBox "Please Specify a Port."
    Exit Sub
End If
RemoteAdd = Radd1.Text
RemotePort1 = Rport.Text
RaddL.Caption = RemoteAdd
RportL.Caption = RemotePort1
End Sub

Private Sub Command2_Click()
If RemoteAdd = "" Then
    MsgBox "Please Specify an Address."
    Exit Sub
End If
If RemotePort1 = "" Then
    MsgBox "Please Specify a Port."
    Exit Sub
End If
LocalC.Close
RemoteC.Close
DaemonPort = InputBox$("Please enter in a Port to Listen on.")
If DaemonPort = "" Then DaemonPort = 8079
DPort.Caption = DaemonPort
LocalC.LocalPort = DaemonPort
LocalC.Listen
GoDangit = True
End Sub

Private Sub Command3_Click()
If GoDangit <> True Then
    MsgBox "Please Start a New Connection First."
    Exit Sub
End If
LocalC.Close
RemoteC.Close
LocalC.Listen
Text1.Text = Text1.Text & "Daemon Reset By Administrator." & vbCrLf
End Sub

Private Sub Command4_Click()
LocalC.Close
RemoteC.Close
Text1.Text = Text1.Text & "Daemon Closed By Administrator." & vbCrLf
End Sub

Private Sub DNS1_Error(ByVal Number As Long, Description As String)
Text2.Text = Text2.Text & "DNS Error: " & Number & ": " & Description & vbCrLf
YesNoDNS = True
End Sub
Private Sub DNS1_ResolveCompleted()
YesNoDNS = True
End Sub

Private Sub Form_Load()
serverIP = LocalC.LocalIP
ServerIP1.Caption = serverIP
serverIP = ServerIP1.Caption
YesNoDNS = True
serverMN = DNS1.AddressToName(serverIP)
servermn1.Caption = serverMN
serverMN = servermn1.Caption
End Sub

Private Sub LocalC_Close()
LocalC.Close
RemoteC.Close
LocalC.Listen
Text1.Text = Text1.Text & "Client Disconnected"
End Sub

Private Sub LocalC_ConnectionRequest(ByVal requestID As Long)
RemoteC.RemoteHost = RemoteAdd
RemoteC.RemotePort = RemotePort1
RemoteC.Connect
If LocalC.State <> sckClosed Then
    LocalC.Close
End If
LocalC.Accept requestID
CIP = LocalC.RemoteHostIP
CIPL.Caption = CIP
CIP = CIPL.Caption
CMN = DNS1.AddressToName(CIP)
CMNL.Caption = CMN
CMN = CMNL.Caption
LocalC.SendData "(" & ServerIP1.Caption & ")" & servermn1.Caption & ":" & DaemonPort & ": Proxy Address: " & RemoteAdd & ":" & RemotePort1 & vbCrLf
Text1.Text = Text1.Text & "Client IP: " & CIPL.Caption & vbCrLf & "Client Machine Name: " & CMNL.Caption & vbCrLf & "Remote Address: " & RemoteAdd & vbCrLf & "Remote Port: " & RemotePort1 & vbCrLf
End Sub
Private Sub LocalC_DataArrival(ByVal bytesTotal As Long)
On Error GoTo TimerStart
LocalC.GetData RxData
RemoteC.SendData RxData
Text2.Text = Text2.Text & "(" & CIPL.Caption & ")" & CMN & ": " & RxData & vbCrLf
YesNoDone = False
Exit Sub
TimerStart:
    Timer1.Enabled = True
    
End Sub
Private Sub LocalC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Text2.Text = Text2.Text & "Remote Client Winsock Control ERROR: " & Number & ": " & Description & vbCrLf
LocalC.Close
RemoteC.Close
LocalC.Listen
Text1.Text = Text1.Text & "Client Disconnected"
End Sub

Private Sub RemoteC_Close()
If YesNoDone <> True Then
    Do Until YesNoDone = True
        If YesNoDone = True Then
            Exit Do
        End If
    Loop
End If
LocalC.Close
RemoteC.Close
LocalC.Listen
Text1.Text = Text1.Text & "Client Disconnected" & vbCrLf
End Sub

Private Sub RemoteC_Connect()
LocalC.SendData "(" & ServerIP1.Caption & ")" & servermn1.Caption & ":" & DaemonPort & ": Connected to: " & RemoteAdd & ":" & RemotePort1 & vbCrLf
End Sub

Private Sub RemoteC_DataArrival(ByVal bytesTotal As Long)
RemoteC.GetData RxData2
LocalC.SendData RxData2
YesNoDone = True
Text2.Text = Text2.Text & RemoteAdd & ":" & RemotePort1 & ": " & RxData2 & vbCrLf
End Sub
Private Sub RemoteC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Text2.Text = Text2.Text & "Remote Connection Winsock Control ERROR: " & Number & ": " & Description & vbCrLf
LocalC.Close
RemoteC.Close
LocalC.Listen
Text1.Text = Text1.Text & "Client Disconnected"
End Sub
Private Sub Text1_Change()
Text1.SelStart = Len(Text1.Text)
If Len(Text1.Text) >= 40750 Then
     Text1.Text = ""
End If
End Sub

Private Sub Text2_Change()
Text2.SelStart = Len(Text2.Text)
If Len(Text2.Text) >= 40750 Then
     Text2.Text = ""
End If
End Sub

Private Sub Timer1_Timer()
On Error GoTo FinishThisKrud
RemoteC.SendData RxData
Timer1.Enabled = False
FinishThisKrud:
    Timer1.Enabled = False
End Sub
