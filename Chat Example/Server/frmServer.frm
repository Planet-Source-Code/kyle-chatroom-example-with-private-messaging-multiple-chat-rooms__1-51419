VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Example - Server"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
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
   ScaleHeight     =   5775
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   5520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4939
            MinWidth        =   4939
            Text            =   "Sever: Running"
            TextSave        =   "Sever: Running"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "2:19 AM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "2/2/2004"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraData 
      Caption         =   "Data: (Raw)"
      Height          =   2295
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   5655
      Begin VB.TextBox txtData 
         Appearance      =   0  'Flat
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.Frame fraChat 
      Caption         =   "Chat Rooms:"
      Height          =   3015
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.CommandButton cmdCreate 
         Caption         =   "Create Room"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1575
      End
      Begin MSComctlLib.TreeView lstUsers 
         Height          =   2295
         Left            =   1800
         TabIndex        =   5
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   4048
         _Version        =   393217
         Indentation     =   176
         Style           =   1
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.TreeView lstRooms 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3201
         _Version        =   393217
         Indentation     =   176
         Style           =   1
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label lblUser 
         BackStyle       =   0  'Transparent
         Caption         =   "User/Chat Room:"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblRooms 
         BackStyle       =   0  'Transparent
         Caption         =   "Rooms:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame fraOnline 
      Caption         =   "Online:"
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
      Begin VB.ListBox lstOnline 
         Appearance      =   0  'Flat
         Height          =   2565
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSWinsockLib.Winsock Ws 
      Index           =   0
      Left            =   120
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Chat room example/Multiple chat rooms
'Created by: Kyle W.
'Age: 17
'Location: Fargo, North Dakota
'Genre: Punk Music
'=================================================================================
Option Explicit
Private Sub cmdCreate_Click()
Dim Mesg As String
    'Create a room
    Mesg = InputBox("Please type below the name of the chat room you would like to create", "Create a chatroom")
    If Mesg = "" Then
        Exit Sub
    ElseIf Mesg <> "" Then
        lstRooms.Nodes.Add , , , Mesg & " (0)"
    End If
End Sub
Private Sub Form_Load()
    'Listens for Clients to connect
    Call SockListen(Ws(0), "5106")
    Call AddRooms
End Sub
Private Sub Ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim Socket As Integer
    'Accepts the connection, loads another Socket
    Socket% = Ws().UBound + 1
    Load Ws(Socket%)
    Ws(Socket%).Accept (requestID)
End Sub
Public Function GatherRoomList()
Dim X As Integer
    'Gathers the list of Rooms
    For X% = 1 To lstRooms.Nodes.Count
        GatherRoomList = GatherRoomList & lstRooms.Nodes.Item(X%).Text & "/"
    Next X%
End Function
Function RoomCount(Rooms As String, AddOrRemove As Boolean)
Dim X As Integer, Room As String, Cont As String
    'Way of keeping track of how many users are in chat
    For X% = 1 To lstRooms.Nodes.Count
        'We split the Rooms, Everytime a user logins in it Adds +1 to the Usercount, when they log out, it
        'Subtracts -1
        Room = Split(lstRooms.Nodes.Item(X%).Text, " (")(0)
        Cont = Split(lstRooms.Nodes.Item(X%).Text, " (")(1)
        Cont = Replace(Cont, ")", "")
        If LCase(Room) = LCase(Rooms) Then
            If AddOrRemove = True Then
                Cont = Cont + 1
                lstRooms.Nodes.Item(X%).Text = Room & " (" & Cont & ")"
                Exit Function
            ElseIf AddOrRemove = False Then
                Cont = Cont - 1
                lstRooms.Nodes.Item(X%).Text = Room & " (" & Cont & ")"
                Exit Function
            End If
        End If
    Next X%
End Function
Public Function GetUsersInChat(Room As String)
Dim X As Integer, User As String, Rooms As String
    'Searchs through the User/Room list to see what users are in what room,
    'It gathers the users then is able to send the packet as one line.
    'Easy way of sending a lot of data.
    For X% = 1 To lstUsers.Nodes.Count
        User = Split(lstUsers.Nodes.Item(X%).Text, "/")(0)
        Rooms = Split(lstUsers.Nodes.Item(X%).Text, "/")(1)
        If LCase(Rooms) = LCase(Room) Then
            GetUsersInChat = GetUsersInChat & User & "/"
        End If
    Next X%
End Function
Private Sub Ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim Data As String, PacketType As String, Packet As String, PacketArray() As String
Dim I As Integer
    'Data is gotten.
    Call Ws(Index).GetData(Data, vbString)
    'Data is seperated
    PacketArray() = Split(Data, Sep)
    '[Second Packet]
    PacketType = PacketArray(2)
    '[Forth Packet]
    Packet = PacketArray(4)
    Select Case PacketType & Packet
        'Login Packet:
        Case Is = "A803"
            lstOnline.AddItem (PacketArray(6))
            Call SendData(Header("83" & Sep & "84" & Sep & "01", "á"), Ws(Index))
            Call AddData(Data, "Login")
            Ws(Index).Tag = LCase(PacketArray(6))
        'Get Room List Packet:
        Case Is = "L595"
            Call SendData(Header("72" & Sep & "00" & Sep & GatherRoomList & Sep, "8B"), Ws(Index))
            Call AddData(Data, "Get Room List")
        'Private Message Packet:
        Case Is = "*®91"
            Call SendData(Data$, Ws(FindSocket(PacketArray(8))))
        'Create Room Packet:
        Case Is = "N*04"
            lstRooms.Nodes.Add , , , PacketArray(6) & " (0)"
            Call AddData(Data, "Create Room")
        'Join Room Packet:
        Case Is = "G317"
            lstUsers.Nodes.Add , , , PacketArray(6) & "/" & PacketArray(8)
            Call SendData(Header("29" & Sep & "11" & Sep & "90" & Sep & GetUsersInChat(PacketArray(8)) & Sep, "1F"), Ws(Index))
            Pause (0.5)
            For I% = 1 To Ws().UBound
                Select Case Ws(I%).State
                    Case Is = sckConnected
                        Ws(I%).SendData (Header("90" & Sep & "15" & Sep & "75" & Sep & PacketArray(6) & Sep & PacketArray(8) & Sep, "9D"))
                        DoEvents%
                End Select
            Next I%
            Call RoomCount(PacketArray(8), True)
            Call AddData(Data, "Join Room")
        'Message to Room Packet:
        Case Is = "K923"
            For I% = 1 To Ws().UBound
                Select Case Ws(I%).State
                    Case Is = sckConnected
                        Ws(I%).SendData (Data$)
                        DoEvents%
                End Select
            Next I%
            Call AddData(Data, "Chat Text")
        'Logout Packet:
        Case Is = "*JN/E174"
            Call RemoveUser(PacketArray(7))
            Call RemoveUserFromRoom(PacketArray(7))
            Call RoomCount(PacketArray(9), False)
            Ws(Index).Tag = ""
            For I% = 1 To Ws().UBound
                Select Case Ws(I%).State
                    Case Is = sckConnected
                        Ws(I%).SendData (Data$)
                        DoEvents%
                End Select
            Next I%
        'Typing Packet:
        Case Is = "m/°48"
            Call SendData(Data$, Ws(FindSocket(PacketArray(7))))
    End Select
End Sub
