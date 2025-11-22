Option Explicit

Dim shell, connected, fakeIP, server, i, loadingAnim
Set shell = CreateObject("WScript.Shell")

connected = False
fakeIP = "51.178.112.33"

Do
    If Not connected Then
        server = InputBox( _
            "TrigerVPN" & vbCrLf & vbCrLf & _
            "Choose a server:" & vbCrLf & _
            "1 - France" & vbCrLf & _
            "2 - Germany" & vbCrLf & _
            "3 - USA" & vbCrLf & _
            "4 - Canada" & vbCrLf & vbCrLf & _
            "Type the number of the server:", _
            "TrigerVPN Server Selection")

        If server = "" Then WScript.Quit

        shell.Popup "Connecting to server...", 1, "TrigerVPN", 64

        For i = 1 To 5
            loadingAnim = String(i, ".")
            shell.Popup "Connecting" & loadingAnim, 1, "TrigerVPN", 64
            WScript.Sleep 500
        Next

        shell.Popup "Connected!" & vbCrLf & vbCrLf & _
                    "Server: " & GetServerName(server) & vbCrLf & _
                    "IP: " & fakeIP, 2, "TrigerVPN", 64

        connected = True

    Else
        Dim choice
        choice = shell.Popup( _
            "TrigerVPN" & vbCrLf & vbCrLf & _
            "Status: Connected" & vbCrLf & _
            "Server: " & GetServerName(server) & vbCrLf & _
            "IP: " & fakeIP & vbCrLf & vbCrLf & _
            "Do you want to disconnect?", _
            0, "TrigerVPN", 4 + 32)

        If choice = 6 Then
            shell.Popup "Disconnecting...", 1, "TrigerVPN", 64
            WScript.Sleep 1000
            shell.Popup "Disconnected!", 1, "TrigerVPN", 64
            connected = False
        Else
            WScript.Quit
        End If
    End If
Loop

Function GetServerName(code)
    Select Case code
        Case "1": GetServerName = "France"
        Case "2": GetServerName = "Germany"
        Case "3": GetServerName = "USA"
        Case "4": GetServerName = "Canada"
        Case Else: GetServerName = "Unknown"
    End Select
End Function
