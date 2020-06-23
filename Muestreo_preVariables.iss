Sub Main
	IgnoreWarning(True)
	Call DirectExtraction()	'ED-Ventas-2010-L4-Ventas diarias por método de pago.IMD
	Call DirectExtraction1()	'ED-Ventas-2010-L4-Ventas diarias por método de pago.IMD
	Call AppendDatabase()	'2011 Anulaciones.IMD
	Call AppendField()	'Anulaciones de 2010 y 2011.IMD
	Client.RefreshFileExplorer
End Sub


' Datos: Extracción directa
Function DirectExtraction
	Set db = Client.OpenDatabase("ED-Ventas-2010-L4-Ventas diarias por método de pago.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "2010 Anulaciones.IMD"
	task.AddExtraction dbName, "", "METODO_PAGO == ""ANULADO"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Datos: Extracción directa
Function DirectExtraction1
	Set db = Client.OpenDatabase("ED-Ventas-2011-L4-Ventas diarias por método de pago.IMD")
	Set task = db.Extraction
	task.IncludeAllFields
	dbName = "2011 Anulaciones.IMD"
	task.AddExtraction dbName, "", "METODO_PAGO == ""ANULADO"""
	task.CreateVirtualDatabase = False
	task.PerformTask 1, db.Count
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Archivo: Anexar bases de datos
Function AppendDatabase
	Set db = Client.OpenDatabase("2010 Anulaciones.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "2011 Anulaciones.IMD"
	dbName = "Anulaciones de 2010 y 2011.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Anexar campo
Function AppendField
	Set db = Client.OpenDatabase("Anulaciones de 2010 y 2011.IMD")
	Set task = db.TableManagement
	Set field = db.TableDef.NewField
	field.Name = "ANULACIONES_ABS"
	field.Description = ""
	field.Type = WI_VIRT_NUM
	field.Equation = "@Abs( SUMA_SIN_IMP )"
	field.Decimals = 2
	task.AppendField field
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Set field = Nothing
End Function