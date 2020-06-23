Sub Main
	Call Stratification1()	'ED-Ventas-2010-L4-Ventas diarias por método de pago.IMD
	Call StratifiedRandomSample()	'Estratificación2.IMD
End Sub


' Análisis: Estratificación
Function Stratification1
	Set db = Client.OpenDatabase("ED-Ventas-2010-L4-Ventas diarias por método de pago.IMD")
	Set task = db.Stratification
	task.IncludeAllFields
	dbName = "Estratificación1.IMD"
	task.OutputDBName = dbName
	task.IncludeInterval = FALSE
	task.SupportStratifiedRandomSample = TRUE
	task.FieldToStratify = "SUMA_CON_IMP"
	task.AddFieldToTotal "SUMA_CON_IMP"
	task.LowerLimit -1433.33
	task.AddUpperLimit -1233.33
	task.AddUpperLimit -1033.33
	task.AddUpperLimit -833.33
	task.AddUpperLimit -633.33
	task.AddUpperLimit -433.33
	task.AddUpperLimit -233.33
	task.AddUpperLimit -33.33
	task.AddUpperLimit 166.67
	task.AddUpperLimit 366.67
	task.AddUpperLimit 566.67
	task.AddUpperLimit 766.67
	task.AddUpperLimit 966.67
	task.AddUpperLimit 1166.67
	task.AddUpperLimit 1366.67
	task.AddUpperLimit 1566.67
	task.AddUpperLimit 1766.67
	task.AddUpperLimit 1966.67
	task.AddUpperLimit 2166.67
	task.AddUpperLimit 2366.67
	task.AddUpperLimit 2566.67
	task.CreateVirtualDatabase = False
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function

' Muestreo: Muestra aleatoria estratificada
Function StratifiedRandomSample
	Set db = Client.OpenDatabase("Estratificación1.IMD")
	Set task = db.StratifyRndSample
	task.StratifyOnBand 1,  1
	task.StratifyOnBand 2,  1
	task.StratifyOnBand 3,  1
	task.StratifyOnBand 4,  2
	task.StratifyOnBand 5,  3
	task.StratifyOnBand 6,  2
	task.StratifyOnBand 7,  4
	task.StratifyOnBand 8,  12
	task.IncludeAllFields
	dbName = "MuesAleat1.IMD"
	task.CreateVirtualDatabase = False
	task.PerformTask dbName, "", "SUMA_CON_IMP", 16984
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function