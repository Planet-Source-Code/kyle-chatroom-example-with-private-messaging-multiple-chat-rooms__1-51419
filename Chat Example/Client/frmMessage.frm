VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMessage 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
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
   ScaleHeight     =   3540
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrType 
      Interval        =   1
      Left            =   1680
      Top             =   1560
   End
   Begin MSComctlLib.StatusBar Stat 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   3225
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8820
            MinWidth        =   8820
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin RichTextLib.RichTextBox txtSend 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMessage.frx":0000
   End
   Begin RichTextLib.RichTextBox txtChat 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMessage.frx":007B
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
    If txtSend.Text <> "" Then
        'Sends Message Data
        Call PMDisplay(Me, ChatUser, txtSend.Text, True)
        Call SendData(Message(ChatUser, Me.Caption, txtSend.Text))
        txtSend.Text = ""
        tmrType.Enabled = True
    End If
End Sub
Private Sub tmrType_Timer()
    'Checks when there is text in the textbox, then sends.
    If Len(txtSend.Text) > 0 Then
        Call SendData(Typing(ChatUser, Me.Caption))
        tmrType.Enabled = False
    End If
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
