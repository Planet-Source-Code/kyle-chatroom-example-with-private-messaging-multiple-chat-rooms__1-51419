VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Example - Login"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Ws 
      Left            =   120
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   4680
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Offline"
            TextSave        =   "Offline"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "2:41 AM"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame fraLogin 
      Caption         =   "Chat Login:"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   3255
      Begin VB.CheckBox chkRemember 
         Appearance      =   0  'Flat
         Caption         =   "Remember Nickname and Server"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.TextBox txtServer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "chatserv.serveftp.com"
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox txtNick 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblServer 
         BackStyle       =   0  'Transparent
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblNick 
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Image imgExample 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   120
      Picture         =   "frmClient.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3240
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
    'Attempt to Connect
    ChatUser = txtNick.Text
    Call Status("Connecting")
    Pause (1)
    Call SocketConnect("5106", Ws)
End Sub
Private Sub Form_Load()
    Call ReadLog
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Call CreateLog
    End
End Sub
Private Sub Ws_Connect()
    'Socket Connected, Send [Login] packet
    Call SendData(Login(ChatUser))
    Call Status("Connected")
End Sub
Private Sub Ws_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
    'Data is gotten.
    Call Ws.GetData(Data$, vbString)
    'All data is seperated and checked in a function in the modData module.
    Call HandleData(Data$, Ws)
End Sub
Private Sub Ws_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Socket error, this will show if the [Server] isn't running
    Call Status("Connection Error")
End Sub
