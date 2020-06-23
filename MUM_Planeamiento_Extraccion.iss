Sub Main
	IgnoreWarning(True)
	Call MUSExtraction()	'ED-Ventas-2010-L4.IMD
	Client.RefreshFileExplorer
End Sub

' Muestreo: Unidad Monetaria
Function MUSExtraction
	Const WI_HighValueHandling_AGGREGATE = 0
	Const WI_HighValueHandling_FILE = 1
	Const WI_RangeOfValues_POSITIVES = 0
	Const WI_RangeOfValues_NEGATIVES = 1
	Const WI_RangeOfValues_ABSOLUTES = 2
	Const WI_TaskType_FIXED = 0
	Const WI_TaskType_CELL = 1
	
	Set db = Client.OpenDatabase("ED-Ventas-2010-L4.IMD")
	Set task = db.MUSExtraction
	task.IncludeAllFields
	task.TaskType = WI_TaskType_FIXED
	task.RangeOfValues = WI_RangeOfValues_POSITIVES
	task.HighValueHandling = WI_HighValueHandling_FILE
	task.HighValueFilename = "Valores altos.IMD"
	task.SampleInterval = 91023.81
	task.RandomValue = 7187.15
	task.FieldToSample = "CON_IMP"
	dbName = "Muestra monetaria.IMD"
	task.MUSExtractionFilename = dbName
	task.CreateVirtualDatabase = False
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function
