'BeNoGa - www.isno.ch
'benoga@isno.ch
'Last Update: 12.09.2013

'Option Explicit
Dim Antwortzeit, Antwortzeit_show, Countdown, Countdown_Popup, WshShell, btn, strUser, strPasswords, Anmeldename_show, Server_IP


'Anmeldung anzeigen --> FALSE oder TRUE
	'--> Bei FALSE wird das Netzlaufwerk mit den Benutzerdaten vom Windows verbunden. Stimmen Windows Benutzername und PW nicht mit dem von ISNO.ch überein, schlägt die Authentifikation fehl.
	'--> Bei TRUE wird ein Dialogfenster aufgemacht, in dem Benutzername und PW selber eingegeben werden müssen.
Anmeldename_show = false

'Delay/Sleep für Autostart in Sekunden
	'--> Wird das Script in den Autostart Ordner verschoben, sollte die Zeit erhöht werden, damit die Netzwerkschnittstelle bereits aktiviert ist, wenn das Script startet.
Countdown = 3

'Delay/Sleep für Schliessen
	'--> Anzahl Sekunden für DialogBox bis es automatisch schliesst, wenn Netzlaufwerk Verbunden oder Fehlgeschlagen ist
	'--> 0 deaktiviert die automatische Schliessung der DialogBox
Countdown_Popup = 5

'Ping Zeit nach erfolgreichem Verbinden in Dialogbox anzeigen --> FALSE oder TRUE
Antwortzeit_show = false

'IP Adresse für den Ping und Connect
Server_IP = "192.168.11.2"


Set WshShell = CreateObject("WScript.Shell")
If Anmeldename_show = FALSE Then
	WshShell.popup "Laufwerke werden verbunden, bitte warten...", Countdown, "Netzlaufwerk verbinden", vbInformation
End If

'Funktion Ping()
Function Ping(strHost)
    Dim oPing, oRetStatus, bReturn
    Set oPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address='" & strHost & "'")
 
    For Each oRetStatus In oPing
        If IsNull(oRetStatus.StatusCode) Or oRetStatus.StatusCode <> 0 Then
            bReturn = False
        Else
            bReturn = True	
				If Antwortzeit_show = TRUE Then
					Antwortzeit =  (vbcrlf & vbcrlf & "Ping dauerte: " & oRetStatus.ResponseTime & "ms")
				End If
        End If
        Set oRetStatus = Nothing
    Next
    Set oPing = Nothing
    Ping = bReturn
End Function

'Funktion Mount()
Function Mount()
	If Anmeldename_show = TRUE Then
		'Input Dialog für Username & Passwort
		strUser = InputBox("Benutzername eingeben", "Netzlaufwerk verbinden")
			If strUser = "" Then
				WScript.Quit
			End If
		strPassword = InputBox("Passwort eingeben", "Netzlaufwerk verbinden")
			If strPassword = "" Then
				WScript.Quit
			End If
	End If
		'Disconnect alte Netzlaufwerke
		set shell = CreateObject("WScript.Shell") 
		shell.run "net use * /d /y",0
		
		' 2 Sekunden warten nach dem Disconnect
		WScript.Sleep 2000

		'Netzlaufwerke verbinden
		Dim objNetwork
		Set objNetwork = WScript.CreateObject("WScript.Network")
		'Err.Clear
		On Error Resume Next
		strLocalDrive = "E:"
		strRemoteShare = "\\" & Server_IP & "\backup"
		objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, False, strUser, strPassword
		strLocalDrive = "H:"
		strRemoteShare = "\\" & Server_IP & "\home"
		objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, False, strUser, strPassword
		strLocalDrive = "M:"
		strRemoteShare = "\\" & Server_IP & "\music"
		objNetwork.MapNetworkDrive strLocalDrive, strRemoteShare, False, strUser, strPassword

			
		Set oShell = CreateObject("Shell.Application")
		Set oShFolder = oShell.Namespace(17)

		' Laufwerknamen vergeben
		For Each oShFolderItem In oShFolder.Items
		   select  case oShFolderItem.Path
		   case "E:\"      oShFolderItem.Name = "Backup"
		   case "H:\"      oShFolderItem.Name = "My Home Folder"
		   case "M:\"      oShFolderItem.Name = "Music"
		   End select
		Next

		Set oShell = Nothing
		Set oShFolder = Nothing
		Set oShFolderItem = Nothing
End Function

'Funktion geglückt
Function Jep()
	btn = WshShell.popup("Netzlaufwerke Verbunden" & vbcrlf & vbcrlf & "Fenster schliesst in " & Countdown_Popup & " Sekunden automatisch..." & Antwortzeit & " ", Countdown_Popup, "Netzlaufwerke Verbunden", vbInformation)
End Function

'Funktion abbrechen, wenn kein Ping
Function Nieet()
	btn = WshShell.popup("Host " & Server_IP & " nicht erreichbar!" & vbcrlf & vbcrlf & "Fenster schliesst in " & Countdown_Popup & " Sekunden automatisch...", Countdown_Popup, "Netzlaufwerke verbinden fehlgeschlagen", vbcritical + vbRetryCancel)
End Function

'Function ErrorISNO
Function ErrorISNO
	'Fehler Erkennung '-2147024891 --> Zugriff Verweigert ausblenden, alle anderen Verbinden
	If Err.Number <> 0 Then
		If Err.Number = -2147024891 Then
		Else
			WshShell.popup " "& Err.Description,0, "Fehler", vbcritical
			Err.Clear
		End If
	ElseIf strUser = vbCancel Then
		WScript.Quit
	ElseIf strPassword = vbCancel Then
		WScript.Quit
	Else
		Jep()
	End If
End Function

'Funktion PingISNO
Function PingISNO()
	If Ping(Server_IP) Then
		Mount()
		ErrorISNO()
	Else
		Nieet()
	End if
End Function

'Host wird gepingt und danach gemountet oder Fehlermeldung angezeigt
PingISNO()

'Schleife für Reconnect
Do While btn = vbRetry
	PingISNO()
	WScript.Sleep 2000 
Loop

'Fertig :)