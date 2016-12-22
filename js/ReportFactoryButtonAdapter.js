function doDownload(sessionId, templateId, reportName, recordId) {
	console.log('[doDownload] sessionId=' + sessionId);
	console.log('[doDownload] templateId=' + templateId);
	console.log('[doDownload] reportName=' + reportName);
	console.log('[doDownload] recordId=' + recordId);
	
	ReportFactory.create(
		sessionId,
		templateId,
		reportName,
		[{"w":[{"f":"ID","t":"id","o":"equals","v":recordId}]}]
	);
}