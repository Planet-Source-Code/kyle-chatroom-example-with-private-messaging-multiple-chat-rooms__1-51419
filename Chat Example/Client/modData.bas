Attribute VB_Name = "modData"
'Chat Example Client: By/Kyle W.
'This may be a bit confusing, but I tried to make it easy to understand.
'The reason for a Seperator, and Header is for Security, it encrypts the packets
'so it's not all raw data incase you would want to create a real successful chat
'login.
'=================================================================================
Public Const Sep = "¥€"
Public ChatUser As String, Packet As String, NewPM(30) As New frmMessage, PmCount As Integer
Function SendData(Data As String)
    'Easier and cleaner way of sending Data.
    With frmClient
        Select Case .Ws.State
            Case Is = sckConnected
                .Ws.SendData (Data$)
                DoEvents%
        End Select
    End With
End Function
Function CreateLog()
    'Creates a log file that remembers your Nickname and Server
    With frmClient
        If .chkRemember.Value = 1 Then
            Close #1
            Open App.Path & "/Log.dat" For Output As #1
                Print #1, "Nick=" & .txtNick.Text & vbCrLf & "Serv=" & .txtServer.Text
            Close #1
        End If
    End With
End Function
Function OpenPM(User As String)
'Opens a new PM Window
Dim X As Integer
    With frmChat
        For X% = 0 To 30
            If LCase(NewPM(X%).Caption) = LCase(User) Then
                Exit Function
            End If
        Next X%
        NewPM(PmCount%).Caption = LCase(User)
        Load NewPM(PmCount%)
        NewPM(PmCount%).Show
        PmCount% = PmCount% + 1
    End With
End Function
Function ReadLog()
Dim Data As String
'Reads the log to Find the Nickname and Server
On Error GoTo ErrControl
    With frmClient
        Close #1
        Open App.Path & "/Log.dat" For Input As #1
            Do Until EOF(1)
                Input #1, Data
                Select Case Left(Data, 4)
                    Case Is = "Nick"
                        Data$ = Replace(Data$, "Nick=", "")
                        .txtNick.Text = Data$
                    Case Is = "Serv"
                        Data$ = Replace(Data$, "Serv=", "")
                        .txtServer.Text = Data$
                End Select
            Loop
        Close #1
    End With
ErrControl:
End Function
Sub Pause(interval)
'Pause (how many seconds)
Dim Current
    Current = Timer
    Do While Timer - Current < Val(interval)
    DoEvents
    Loop
End Sub
Function ChatDisplay(User As String, What As String, InOrOut As Boolean)
    'Displays the chat text in the room
    With frmChat.txtChat
        .SelStart = Len(.Text)
        .SelBold = True
        'I use a boolean so I dont have to waste space putting two functions for displaying text from a user
        'and your own text.
        If InOrOut = True Then
            .SelColor = vbRed
        ElseIf InOrOut = False Then
            .SelColor = vbBlue
        End If
        .SelText = User & ": "
        .SelBold = False
        .SelColor = vbBlack
        .SelText = What & vbCrLf
    End With
End Function
Function PMDisplay(frm As Form, User As String, What As String, InOrOut As Boolean)
    'Displays private message text
    With frm.txtChat
        .SelStart = Len(.Text)
        .SelBold = True
        'I use a boolean so I dont have to waste space putting two functions for displaying text from a user
        'and your own text.
        If InOrOut = True Then
            .SelColor = vbRed
        ElseIf InOrOut = False Then
            .SelColor = vbBlue
        End If
        .SelText = User & ": "
        .SelBold = False
        .SelColor = vbBlack
        .SelText = What & vbCrLf
    End With
End Function
Function HandleData(Data As String, Sock As Winsock)
Dim PacketArray() As String, ReadPacket As String, PacketType As String
Dim ChatRooms() As String, I As Integer
    PacketArray() = Split(Data, Sep)
    ReadPacket = PacketArray(4)
    PacketType = PacketArray(2)
    Select Case PacketType & ReadPacket
        'Nickname is Logged in, now must get the list of Rooms
        Case Is = "á84"
            Call SendData(GetRoomList)
        'Rooms Received
        Case Is = "8B00"
            With frmRooms
                .lstRooms.Nodes.Clear
                '/ is the seperator, I split the packet, use a Loop to get add the rooms, saves us time from
                'having the server socket keep sending each room, I just send it all at once and split it
                ' on the server end
                ChatRooms() = Split(PacketArray(5), "/")
                For I% = 0 To UBound(ChatRooms) - 1
                    .lstRooms.Nodes.Add , , , ChatRooms(I%)
                Next I%
                .Show
            End With
        'Userlist Received
        Case Is = "1F11"
            With frmChat
                '/ is the seperator, I split the packet to get the users
                ChatRooms() = Split(PacketArray(6), "/")
                For I% = 0 To UBound(ChatRooms) - 1
                    .lstUsers.Nodes.Add , , , ChatRooms(I%)
                Next I%
            End With
        'Chat room text
        Case Is = "K923"
            If frmChat.Caption = "Chat Room -- " & PacketArray(9) Then
                If LCase(PacketArray(6)) <> LCase(ChatUser) Then
                    Call ChatDisplay(PacketArray(6), PacketArray(8), False)
                End If
            End If
        'User joined the room
        Case Is = "9D15"
            If frmChat.Caption = "Chat Room -- " & PacketArray(7) Then
                If LCase(PacketArray(6)) <> LCase(ChatUser) Then
                    frmChat.lstUsers.Nodes.Add , , , PacketArray(6)
                    Call JoinLeaveRoom(PacketArray(6), True)
                End If
            End If
        'User has logged out
        Case Is = "*JN/E174"
            If frmChat.Caption = "Chat Room -- " & PacketArray(9) Then
                Call RemoveUser(PacketArray(7))
                Call JoinLeaveRoom(PacketArray(7), False)
            End If
        'Message from a User
        Case Is = "*®91"
            'Searchs through the PM's to see if an open message is already open with the user
            For I% = 0 To 30
                If LCase(NewPM(I%).Caption) = LCase(PacketArray(6)) Then
                    'If it is, it simply adds the text to that window, then exits this function
                    Call PMDisplay(NewPM(I%), PacketArray(6), PacketArray(9), False)
                    NewPM(I%).Stat.Panels.Item(1).Text = "Last message: " & Time & " / " & Date
                    Exit Function
                End If
            Next I%
            'If not, it simply adds the text, and pops open a new window.
            NewPM(PmCount%).Caption = LCase(PacketArray(6))
            Call PMDisplay(NewPM(PmCount%), PacketArray(6), PacketArray(9), False)
            NewPM(PmCount%).Stat.Panels.Item(1).Text = "Last message: " & Time & " / " & Date
            Load NewPM(PmCount%)
            NewPM(PmCount%).Show
            PmCount% = PmCount% + 1
        'Displays when a user is typing
        Case Is = "m/°48"
            'Searchs the pms for an open Message Box
            For I% = 0 To 30
                'If one is found, it simply shows that the user is typing.
                If LCase(NewPM(I%).Caption) = LCase(PacketArray(6)) Then
                    NewPM(I%).Stat.Panels.Item(1).Text = LCase(PacketArray(6)) & " is typing a message..."
                    Exit Function
                End If
            Next I%
    End Select
End Function
Function RemoveUser(User As String)
Dim X As Integer
    'Removes a user from the list
    With frmChat
        For X% = 1 To .lstUsers.Nodes.Count
            If LCase(.lstUsers.Nodes.Item(X%).Text) = LCase(User) Then
                .lstUsers.Nodes.Remove (X%)
                Exit Function
            End If
        Next X%
    End With
End Function
Function Status(Stat As String)
    'Changes the status on the client.
    With frmClient
        .Stat.Panels.Item(1).Text = Stat
    End With
End Function
Function ChatEntry(Room As String, EntryText As String)
    'Shows an entry status
    With frmChat.txtChat
        .SelStart = Len(.Text)
        .SelBold = False
        .SelColor = &H8000&
        .SelText = vbCrLf & vbCrLf & "  You have entered "
        .SelBold = True
        .SelText = Chr(34) & Room & Chr(34) & ". "
        .SelBold = False
        .SelColor = vbBlack
        .SelText = EntryText & vbCrLf & vbCrLf
    End With
End Function
Function JoinLeaveRoom(User As String, JoinLeave As Boolean)
    'Displays whether a user joins or leaves the room
    With frmChat.txtChat
        .SelStart = Len(.Text)
        .SelBold = False
        .SelColor = vbRed
        .SelText = User & " "
        .SelColor = vbBlack
        If JoinLeave = True Then
            .SelText = "has joined the room" & vbCrLf
        ElseIf JoinLeave = False Then
            .SelText = "has left the room" & vbCrLf
        End If
    End With
End Function
Function Header(Pack As String, PacketType As String)
    'I use this for security purposes, in case one user packet sniffs, gets the packets you use, you can
    'use this to help from people creating error's in your client.
    Header = "CHT" & Sep & "30" & Sep & PacketType & Sep & Pack
End Function
Function SocketConnect(Port As String, Sock As Winsock)
    'Connects the socket.
    Sock.Close
    Sock.Connect frmClient.txtServer.Text, Port
End Function
'Packets for Chat Example:By/Kyle W.
Function Login(User As String)
    Packet = "94" & Sep & "03" & Sep & "10" & Chr(38) & Sep & User & Sep
    Login = Header(Packet, "A8")
End Function
Function GetRoomList()
    Packet = "01" & Sep & "95" & Sep & "17" & Sep
    GetRoomList = Header(Packet, "L5")
End Function
Function CreateRoom(Room As String)
    Packet = "95" & Sep & "04" & Sep & "02" & Sep & Room & Sep
    CreateRoom = Header(Packet, "N*")
End Function
Function JoinRoom(Room As String, User As String)
    Packet = "85" & Sep & "17" & Sep & "71" & Sep & User & Sep & "55" & Sep & Room & Sep
    JoinRoom = Header(Packet, "G3")
End Function
Function ChatSend(User As String, What As String, Room As String)
    Packet = "99" & Sep & "23" & Sep & "02" & Sep & User & Sep & "05" & Sep & What & Sep & Room & Sep
    ChatSend = Header(Packet, "K9")
End Function
Function LogOut(User As String, Room As String)
    Packet = "888" & Sep & "174" & Sep & "49" & Sep & "19" & Sep & User & Sep & "392" & Sep & Room & Sep
    LogOut = Header(Packet, "*JN/E")
End Function
Function Message(From As String, ToUser As String, What As String)
    Packet = "832" & Sep & "91" & Sep & "81" & Sep & From & Sep & "18" & Sep & ToUser & Sep & What & Sep
    Message = Header(Packet, "*®")
End Function
Function Typing(From As String, ToUser As String)
    Packet = "983" & Sep & "48" & Sep & "63" & Sep & From & Sep & ToUser & Sep
    Typing = Header(Packet, "m/°")
End Function


