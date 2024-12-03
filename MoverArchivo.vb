Sub MoverArchivo()
	Dim FSO As Object
	Set FSO = CreateObject("scripting.FileSystemObject")

	Dim UltmaLine, Usuario, Origen, Destino, FechaGestion, MesGestion, NombreMP3, CarDes, CarDesRut, CarOriRut, ExisteCarpeta, ExisteArchivo, CopiArch As String


	Origen = InputBox("Ingresa la ruta del origen de los archivos")

	Usuario = Environ("username")
	Destino = "C:\Users\" & Usuario & "\Documents\Recuperacion\Audios" & "\"

	UltmaLine = ThisWorkbook.Sheets("BDD").UsedRange.Columns("BN").Rows.Count

		For i = 3 To UltmaLine
			FechaGestion = ThisWorkbook.Sheets("BDD").Cells(i, 10)
			NameMesGestion = StrConv(MonthName(Month(FechaGestion)), vbProperCase)
			
			NombreMP3 = ThisWorkbook.Sheets("BDD").Cells(i, 67) & ".mp3"
			
			'Concatena la ruta de los audios originales y indicar cual es el audio
			CarOriRut = Origen & "\" & NombreMP3
			
			'Genera el nombre de la carpeta segun la fecha de gestion
			CarDes = Replace(FechaGestion, "/", ".")
			
			CarDesRutMes = Destino & NameMesGestion
			ExisteCarpetaMes = Dir(CarDesRutMes, vbDirectory)
			
			If ExisteCarpetaMes = "" Then
				'Crea la ruta si no existe
				MkDir CarDesRutMes
			Else
				CarDesRutFecha = CarDesRutMes & "\" & CarDes
				ExisteCarpetaFecha = Dir(CarDesRutFecha, vbDirectory)
				CopiArch = CarDesRutFecha & "\" & NombreMP3
				
				'Verifica si la ruta existe, segun la variable: ExisteCarpeta
				If ExisteCarpetaFecha = "" Then
					'Crea la ruta si no existe
					MkDir CarDesRutFecha
				Else
					'Guarda la ruta de los audios con el nombre del archivo
					ExisteArchivo = Dir(CarOriRut)
					
					'Verifica que no exista el archivo, segun la variable: ExisteArchivo
					If ExisteArchivo = "" Then
						'MsgBox "No existe el archivo " & NombreMP3
					Else
						On Error Resume Next
								FSO.CopyFile Source:=CarOriRut, Destination:=CopiArch
						On Error GoTo 0
					End If
				End If
			End If
		Next
End Sub