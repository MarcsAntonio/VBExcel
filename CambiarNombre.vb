Sub CambiarNombre()
	'Antes de correr la macro, elije las celdas que tengan la ruta
	'del nombre actual, es decir a partir de A2
	Dim NombreNuevo As String
	Dim NombreAnterior As String
	Set RCAR = ActiveSheet.UsedRange.Columns("BN").Cells
	directorio = InputBox("Ingresa la ruta donde quieres crear las carpetas")
	'Si no encuentra algún archivo, continuará con el siguiente
	On Error Resume Next
	For Each Celdas In Range("BN3:BN1000")
		NombreAnterior = directorio & "\" & Celdas.Value
	'El dato del nombre nuevo será la columna D, especificado con 3
	text1 = Celdas.Offset(0, 1).Value
		NombreNuevo = directorio & "\" & text1 & ".mp3"
		Name NombreAnterior As NombreNuevo
	Next Celdas
	On Error GoTo 0
End Sub
