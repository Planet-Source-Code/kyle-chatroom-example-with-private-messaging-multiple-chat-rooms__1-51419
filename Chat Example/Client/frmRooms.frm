VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRooms 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Example - Rooms"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
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
   ScaleHeight     =   3375
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRooms 
      Caption         =   "Room (User Count):"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      Begin MSComctlLib.TreeView lstRooms 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   4895
         _Version        =   393217
         Style           =   7
         BorderStyle     =   1
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create Room"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmRooms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()
Dim Msg As String
    Msg = InputBox("Please type the name of the room you want to create:", "Create a room")
    If Msg = "" Then
        Exit Sub
    ElseIf Msg <> "" Then
        Call SendData(CreateRoom(Msg))
        lstRooms.Nodes.Add , , , Msg & " (0)"
    End If
End Sub
Private Sub lstRooms_DblClick()
Dim Rooms As String
    Rooms = Split(lstRooms.SelectedItem.Text, " (")(0)
    Call SendData(JoinRoom(Rooms, ChatUser))
    With frmChat
        .Caption = "Chat Room -- " & Rooms
        .txtChat.Text = ""
        .lstUsers.Nodes.Clear
        Call ChatEntry(Rooms, "Chat with new people, Enjoy!")
        .Show
    End With
    Unload Me
    frmClient.Visible = False
End Sub
