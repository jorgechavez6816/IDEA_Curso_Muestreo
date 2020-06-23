Sub Main
	Call RandomSample()	'ED-Ventas-2011-L4-Ventas diarias por producto (solo cupones).IMD
End Sub


' Muestreo: Aleatorio
Function RandomSample
	Set db = Client.OpenDatabase("ED-Ventas-2011-L4-Ventas diarias por producto (solo cupones).IMD")
	Set task = db.RandomSample
	task.IncludeAllFields
	dbName = "MuesAleat.IMD"
	task.CreateVirtualDatabase = False
	task.PerformTask dbName, "", 50, 1, db.Count, 3412, False
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function