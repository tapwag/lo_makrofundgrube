REM  *****  BASIC  *****

Sub Riestersparbetrag
	
	'Variablendeklaration
	Dim intGespart As Integer
	Dim strAusgabe As String
	
	'Dies ist ein Kommentar durch einen Apostroph gekennzeichnet
	REM Dies ist auch ein Kommentar - diesmal eingeleteitet mit REM (Remark)
	
	'Dies ist ein Meldungsfenster
	MsgBox "Mein Ziel ist es, für die Riester-Rente mehr als " & Format (20000, "#,##0.00") &_
	" Euro anzusparen!"

	'Abfrage des angesparten Betrags
	intGespart = inputBox ("Wieviel hast Du schon angespart?")
	
	
	'Verschachtelte if-Abfrage prüft ob zuerst unter 10000 Euro danach zwischen 10000 und 19999 und dann gleich oder über 20000
	If intGespart < 10000 Then
		strAusgabe = "Es ist noch nicht die Häfte erreicht."
	Else

		If intGespart > 10000 And intGespart < 19999 Then
			strAusgabe = "Immerhin über die Hälfte"
	Else
				If intGespart >= 20000 Then
					strAusgabe = "Herzlichen Glückwunsch. Du hast dein Ziel erreicht."

	End If
	End If
	End If
	'Ausgabedialog
	MsgBox strAusgabe, 64
	
End Sub
