REM  *****  BASIC  *****

Sub Welche_Datenbanken
	Dim objDatabaseContext As Object
	Dim objNamen
	Dim i As Integer
	Dim strListe As String
	objDatabaseContext = _
		createUnoService("com.sun.star.sdb.DatabaseContext")
	objNamen = objDatabaseContext.getElementNames()
	For i = 0 To UBound(objNamen())
		strListe = strListe & Chr(13) & objNamen(i)
	Next
	MsgBox strListe
End Sub


