Sub Main
	Call DirectExtraction()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Datos: Extracción directa
Function DirectExtraction
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	task.AddField "STATUS", "", WI_EDIT_CHAR, 10, 0, """Cuarentena"""
	dbName = "Extraccion_Cuarentena.IMD"
	task.AddExtraction dbName, "", "( FECHA_FACT  >= ""20150316"" .AND.  FECHA_FACT  <= ""20150624"") .AND.  (TOTAL  >= 300 .AND.  TOTAL  <= 10500) .AND. (COD_PROD  = ""04"" .OR.  COD_PROD  = ""05"" .OR.  COD_PROD  = ""06"")"
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function