VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7815
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
   ScaleHeight     =   4455
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin RichTextLib.RichTextBox txtSend 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChat.frx":0000
   End
   Begin MSComctlLib.TreeView lstUsers 
      Height          =   4215
      Left            =   5760
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   7435
      _Version        =   393217
      Indentation     =   176
      Style           =   1
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   6376
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChat.frx":007B
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuLogout 
         Caption         =   "Logout"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
Dim Room As String
    If txtSend.Text <> "" Then
        Room = Split(Me.Caption, " -- ")(1)
        Call SendData(ChatSend(ChatUser, txtSend.Text, Room))
        Call ChatDisplay(ChatUser, txtSend.Text, True)
        txtSend.Text = ""
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Dim Room As String
    Room = Split(frmChat.Caption, " -- ")(1)
    Cancel = 1
    Call SendData(LogOut(ChatUser, Room))
    frmClient.Show
    Me.Visible = False
    txtChat.Text = ""
    lstUsers.Nodes.Clear
    Call Status("Offline")
End Sub
Private Sub lstUsers_DblClick()
On Error Resume Next
    Call OpenPM(lstUsers.SelectedItem.Text)
End Sub
Private Sub mnuAbout_Click()
    MsgBox "About Chat Example - By Kyle" & vbCrLf & vbCrLf & "Name: Kyle W." & vbCrLf & "Age: 18" & vbCrLf & "Location: Fargo, North Dakota" & vbCrLf & "Reason: I created this program to show people out there how easy it was to create a multi room chat client, email me at _Kyle_@msn.com if you have any questions, also my AIM is SayAnythingRox and Yahoo! ID is customize.", vbInformation, "About Chat Example"
End Sub
Private Sub mnuLogout_Click()
    Unload Me
End Sub
Private Sub txtSend_Change()
    If Len(txtSend.Text) > 0 Then
        cmdSend.Enabled = True
    ElseIf Len(txtSend.Text) <= 0 Then
        cmdSend.Enabled = False
    End If
End Sub
Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = 13
            Call cmdSend_Click
            KeyCode = 0
    End Select
End Sub
