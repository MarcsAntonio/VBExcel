Sub CrearCarpetas()
	Ruta = InputBox("Ingresa la ruta donde quieres crear las carpetas")
	Celda = InputBox("Primera celda")
	Range(Celda).Select
	Do While ActiveCell.Value <> ""
	MkDir (Ruta & "/" & ActiveCell.Value)
	ActiveCell.Offset(1, 0).Select
	Loop
End Sub