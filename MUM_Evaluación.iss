Sub Main
	IgnoreWarning(True)
	Call MUSEvaluate()	'Muestra monetaria.IMD
	Call MUSEvaluate1()	'Muestra monetaria.IMD
	Call MUSEvaluate2()	'Muestra monetaria.IMD
	Client.RefreshFileExplorer
End Sub

' Muestreo: Unidad Monetaria
Function MUSEvaluate
	Const WI_HighValueHandling_AGGREGATE = 0
	Const WI_HighValueHandling_FILE = 1
	Const WI_PrecisionLimits_Upper = 1
	Const WI_PrecisionLimits_UpperandLower = 2
	Const WI_MUS_CLASSICAL_PPS_EVALUATION = 1
	Const WI_MUS_CELL_EVALUATION = 2
	Set db = Client.OpenDatabase("Muestra monetaria.IMD")
	Set task = db.MUSEvaluate
	task.Method = WI_MUS_CELL_EVALUATION
	task.AuditAmountField = "CANT_AUDIT"
	task.BookField = "CON_IMP"
	task.ReferenceField = "REFERENCIA"
	task.ConfidenceLevel = 90.00
	task.PopulationValue = 3822999.85
	task.SampleSize = 42
	task.ResultName = db.UniqueResultName("Muestreo por unidades monetarias - Evaluación de celda")
	task.BasicPrecisionPricing = 100.00
	task.HighValueHandling = WI_HighValueHandling_FILE
	task.HighValueFilename = "Valores altos.IMD"
	task.HighValueAuditAmountField = "CANT_AUDIT"
	task.HighValueBookField = "CON_IMP"
	task.HighValueReferenceField = "REFERENCIA"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' Muestreo: Unidad Monetaria
Function MUSEvaluate1
	Const WI_HighValueHandling_AGGREGATE = 0
	Const WI_HighValueHandling_FILE = 1
	Const WI_PrecisionLimits_Upper = 1
	Const WI_PrecisionLimits_UpperandLower = 2
	Const WI_MUS_CLASSICAL_PPS_EVALUATION = 1
	Const WI_MUS_CELL_EVALUATION = 2
	
	Set db = Client.OpenDatabase("Muestra monetaria.IMD")
	Set task = db.MUSEvaluate
	task.Method = WI_MUS_CLASSICAL_PPS_EVALUATION
	task.AuditAmountField = "CANT_AUDIT"
	task.BookField = "CON_IMP"
	task.ReferenceField = "REFERENCIA"
	task.ConfidenceLevel = 90.00
	task.PopulationValue = 3822999.85
	task.SampleSize = 42
	task.ResultName = db.UniqueResultName("Muestreo por unidades monetarias - Evaluación PPS clasica")
	task.PrecisionLimits = WI_PrecisionLimits_Upper
	task.HighValueHandling = WI_HighValueHandling_FILE
	task.HighValueFilename = "Valores altos.IMD"
	task.HighValueAuditAmountField = "CANT_AUDIT"
	task.HighValueBookField = "CON_IMP"
	task.HighValueReferenceField = "REFERENCIA"
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

' Muestreo: Unidad Monetaria
Function MUSEvaluate2
	Set db = Client.OpenDatabase("Muestra monetaria.IMD")
	Set task = db.MUSCombinedEvaluate
	task.ResultName = db.UniqueResultName("Muestreo por unidades monetarias – Evaluación de cota de Stringer")
	task.AddSampleToEvaluation "Muestra monetaria.IMD", "CON_IMP", "CANT_AUDIT", 3822999.85, 42.00, 100.00, "Valores altos.IMD", 91023.81, "CON_IMP", "CANT_AUDIT"
	task.ConfidenceLevel = 90.00
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function

