Sub Main
	Call SystematicSample()	'ED-Ventas-2010-L4.IMD
End Sub


' Muestreo: Muestra sistemática
Function SystematicSample
	Set db = Client.OpenDatabase("ED-Ventas-2010-L4.IMD")
	Set task = db.SystematicSample
	task.IncludeAllFields
	dbName = "MuesSis.IMD"
	task.CreateVirtualDatabase = False
	task.PerformTask dbName, "", 60, 1, db.Count, 3754
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function