Sub Main
	Call DirectExtraction()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Datos: Extracción directa
Function DirectExtraction
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "Detalle_mas_5000.IMD"
	task.AddExtraction dbName, "", "TOTAL >5000"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function