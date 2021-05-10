REM  *****  BASIC  *****

Sub StaubsaugerVerkaeufer

	Dim intStaubsauger As Integer
	Dim strText As String
	
' ---- Die Variablen werden deklariert ----

intStaubsauger = InputBox("Wieviele Staubsauger?")

' ---- Abfrage nach Anzahl der Staubsauger ----

If intStaubsauger < 5 Then

' ---- 1. Fall: Anzahl der Staubsauger

	strText = "Fauler Hund!"	

Else

' ---- Falls mehr als 5 verkauft wurden ----

strText = "Sie erhalten "
	If intStaubsauger < 10 Then
	' ---- 2. Fall zwischen 5 und 10 ----
		strText = strText & _
		"100,00 Euro Provision!"
		
Else
' ---- 3.Fall: zwischen 10 und 20 ----

	If intStaubsauger < 20 Then
		strText = strText & _
		"200,00 Euro Provision!"
		
	Else
	' ---- 4. Fall: mehr als 20 ----
		strText = strText & _
		"500,00 Euro Provision!"
	End If
	End If
End If
MsgBox strText, 64

End Sub
