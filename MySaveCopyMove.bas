REM  *****  BASIC  *****

Sub MySaveCopyMove()
   Dim oDoc As Object
   Dim sDir As String
   Dim sFName As String
   Dim sFPath As String
   Dim oSimpleFileAccess As Object
   Dim oDateTime As Object
   Dim currentDate As String
   Dim currentTime As String
   
   oDoc = ThisComponent
   sDir = oDoc.URL
   sDir = Left(sDir, LastInStr(sDir, "/"))
   sDir = sDir & "_historie/"
   
   oSimpleFileAccess = createUnoService("com.sun.star.ucb.SimpleFileAccess")
   
   ' Ordner erstellen, falls er nicht existiert
   If Not oSimpleFileAccess.Exists(ConvertToURL(sDir)) Then
       oSimpleFileAccess.CreateFolder(ConvertToURL(sDir))
   End If
   
   ' Erstellen des neuen Dateinamens
   oDateTime = createUnoService("com.sun.star.util.DateTime")
   currentDate = Format(Date, "yyyy-MM-dd")
   currentTime = Format(Time, "hh-mm-ss")
   
   sFName = Left(oDoc.Title, Len(oDoc.Title) - 4) & " - " & _
            currentDate & "_" & _
            currentTime & ".ods"
   sFPath = sDir & sFName
   
   ' Speichern einer Kopie der Datei
   oDoc.storeToURL(ConvertToURL(sFPath), Array())
End Sub
Function LastInStr(sText As String, sFind As String) As Integer
   Dim i As Integer
   For i = Len(sText) To 1 Step -1
       If Mid(sText, i, 1) = sFind Then
           LastInStr = i
           Exit Function
       End If
   Next i
   LastInStr = 0
End Function
