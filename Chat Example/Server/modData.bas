Attribute VB_Name = "modData"
'All Data created by: Kyle W.
Public Const Sep = "¥€"
Function SockListen(Sock As Winsock, Port As String)
'Listens for Connections
    Sock.Close
    Sock.LocalPort = Port
    Sock.Listen
End Function
Function AddRooms()
'Adds just a Few Chat rooms to "Start" off with
    With frmServer.lstRooms.Nodes
        .Add , , , "Teens (0)"
        .Add , , , "Health (0)"
        .Add , , , "Punk (0)"
        .Add , , , "Bored (0)"
        .Add , , , "School (0)"
        .Add , , , "Romance (0)"
    End With
End Function
Sub Pause(interval)
Dim Current
    Current = Timer
    Do While Timer - Current < Val(interval)
    DoEvents
    Loop
End Sub
Function Header(Pack As String, PacketType As String)
'For security Purposes, to encrypt packets to make them safe in case you plan
'on creating a secure chat client
    Header = "CHT" & Sep & "30" & Sep & PacketType & Sep & Pack
End Function
Function SendData(Data As String, Sock As Winsock)
'Sends the data from the sock you choose
    Select Case Sock.State
        Case Is = sckConnected
            Sock.SendData (Data$)
            DoEvents%
    End Select
End Function
Function AddData(Data As String, DataType As String)
'Adds the data being passed through the server into a textbox
    With frmServer
        .txtData.Text = .txtData.Text & DataType & ": " & Data & vbCrLf
        .txtData.SelStart = Len(.txtData.Text)
    End With
End Function
Function RemoveUser(User As String)
'Removes the user from the online users category
Dim X As Integer
    With frmServer
        For X% = 0 To .lstOnline.ListCount - 1
            If LCase(.lstOnline.List(X%)) = LCase(User) Then
                .lstOnline.RemoveItem (X%)
                Exit Function
            End If
        Next X%
    End With
End Function
Function FindSocket(User As String)
'When a user logs in, there username is put into the .Tag of the Winsock, so
'when a user private messages somebody, I have it search through the winsock.tag's
'for the user he/she is private messaging, then store it into FindSocket
Dim X As Integer
    With frmServer
        For X% = 1 To .Ws().UBound
            If LCase(.Ws(X%).Tag) = LCase(User) Then
                FindSocket = X%
                Exit Function
            End If
        Next X%
    End With
End Function
Function RemoveUserFromRoom(User As String)
'Removes user from chat
Dim X As Integer, Username As String
    With frmServer
        For X% = 1 To .lstUsers.Nodes.Count
            Username = Split(.lstUsers.Nodes.Item(X%).Text, "/")(0)
            If LCase(Username) = LCase(User) Then
                .lstUsers.Nodes.Remove (X%)
                Exit Function
            End If
        Next X%
    End With
End Function
