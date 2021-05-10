REM  *****  BASIC  *****

Sub Main

	Dim AltKm As Double
	Dim NeuKm As Double
	Dim Liter#
	Dim Verbrauch#
	Dim GefahreneKm#
	' # Shortcut für Double

	MsgBox "Mit diesem Programm wird der durschnittliche Benzinverbrauch auf 100 km berechnet.", 0 , "Benzinverbrauchmakro"
	' , 0 (Nur Okay angeben)
	
	AltKm = InputBox("Kilometerstand beim vorletzten Tanken?", "Vorletzter Kilometerstand")
	NeuKm = InputBox ("Kilometerstand beim letzten Tanken?", "Letzter Kilometerstand")
	Liter = InputBox ("Wieviele Liter haben Sie getankt?", "Tankauffüllung")
   GefahreneKm = Neukm - Altkm
   Verbrauch = Liter / GefahreneKm * 100
   MsgBox "Der Benzinverbrauch auf 100 km betrug "  & Verbrauch 

End Sub
