Class TextIOBaseLite
	Private ostream, bclosed, nomfichier, anewline
	Private mecriture
	
	Private Sub Class_Initialize()
		set ostream = CreateObject("ADODB.Stream")
		closed = True
		mecriture = False
	End Sub
	
	
	Public property get closed
		closed = bclosed
	End Property
	
	Private Property let closed(value)
		bclosed = value
	End Property
  
	Public Sub open(nom_fichier, mode, encoding, newline)
		If mode = "r" Then
			ostream.Charset = encoding
			ostream.open()
			nomfichier = nom_fichier
			ostream.LoadFromFile(nomfichier)
			closed = False
			anewline = newline
		End If
		If mode = "w" Then
			ostream.Charset = encoding
			ostream.open()
			nomfichier = nom_fichier
			closed = False
			anewline = newline
		End If
	End Sub
	
	Public Function readable()
		readable = Not closed
	End Function
	
	Private Function LineSeparator()
		If anewline = "\n" Then
			LineSeparator = 10
		ElseIf anewline = "\r" Then
			LineSeparator = 13
		Else
			LineSeparator = -1
		End If
	End Function
	
	private Function LineSeparatorChar()
		If anewline = "\n" Then
			LineSeparatorChar = VbLf
		ElseIf anewline = "\r" Then
			LineSeparatorChar = vbCr
		Else
			LineSeparatorChar = vbCrLf
		End If
	End Function
	
	Public Function readline()
		If readable() Then
			ostream.LineSeparator = LineSeparator()
			readline = ostream.ReadText(-2)
		Else
			Err.Raise vbObjectError + 513,,"Fichier fermé"
		End If
	End Function
	
	Public Function readlines()
		If readable() Then
			dim content : content = ostream.ReadText ' Lecture du contenu
			readlines = Split(content, LineSeparatorChar()) ' Séparer le contenu en lignes
		Else
			Err.Raise vbObjectError + 513,,"Fichier fermé"
		End If
	End Function
	
	Public Function writable()
		writable = Not closed
	End Function
	
	Public Sub write(strtexte)
		If writable() Then
			ostream.WriteText strtexte
			mecriture = True
		Else
			Err.Raise vbObjectError + 513,,"Fichier fermé"
		End If
	End Sub
	
	Public Function seekable()
		seekable = True
	End Function
	
	Public Function seek(offset)
		if seekable() Then
			ostream.position = offset
			seek = offset
		Else
			Err.Raise vbObjectError + 513,,"Not seekable"
		End if
	End Function
	
	Public Function tell()
		if seekable() Then
			tell = ostream.position
		Else
			Err.Raise vbObjectError + 513,,"Not seekable"
		End if
	End Function
	
	Public Sub flush
		if not closed Then
			if mecriture Then
				dim tmp 
				tmp = tell()
				ostream.SaveToFile nomfichier, 2
				seek(tmp)
			End If
		End if
	End Sub

	Public Sub close()
		if not closed Then
			flush
			ostream.close()
			closed = True
		End if
	End Sub
	
	
	Public Sub Class_Terminate()
		close()
	End Sub
End Class
