function doDownload(sessionId, templateUrl, reportName, recordId) {
	console.log('[doDownload] sessionId=' + sessionId);
	console.log('[doDownload] templateId=' + templateId);
	console.log('[doDownload] reportName=' + reportName);
	console.log('[doDownload] recordId=' + recordId);
	
	let index = templateUrl.lastIndexOf('/');
	let templateId = templateUrl.substring(index+1);
	let bseUrl = templateUrl.substring(0, index);
	
	ReportFactory.init(sessionId, bseUrl);
	ReportFactory.create(
		sessionId,
		templateId,
		reportName,
		[{"w":[{"f":"ID","t":"id","o":"equals","v":recordId}]}]
	);
}