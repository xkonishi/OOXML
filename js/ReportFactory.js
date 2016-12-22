if (!ReportFactory) {
	function ReportFactory() {}
}

(function() {
	
	let sessionId;
	let salesforceBaseUrl;
	
	ReportFactory.init= function(sessionId_, salesforceBaseUrl_) {
		sessionId= sessionId_;
		salesforceBaseUrl= salesforceBaseUrl_;
	};
	
	ReportFactory.create= function(templateId, reportName, conditionsJson, originalJson) {
		console.log('[ReportFactory.create] templateId=' + templateId);
		console.log('[ReportFactory.create] salesforceBaseUrl=' + salesforceBaseUrl);
		$.ajax({
			type: 'GET',
			url: '/services/proxy',
			headers: {
				'Authorization': 'Bearer ' + sessionId,
				'SalesforceProxy-Endpoint': salesforceBaseUrl + '/services/apexrest/ReportTemplate/' + templateId
			},
			success: function(template) {
//				alert('[template] ok');
				if (!template) {
					alert('[ReportFactory.create] Not authorized or not found: template=' + templateId);
					return;
				}
				getData(template, reportName, conditionsJson, originalJson);
			},
			error: function(jqXHR, textStatus, errorThrown) {
				alert('[template] ' + textStatus);
			}
		});
	};
	
	function getData(template, reportName, conditionsJson, originalJson) {
		let dataDefinitionId= template.DataDefinition__c;
		let conditions= JSON.stringify(conditionsJson);
		console.log('[getData] dataDefinitionId=' + dataDefinitionId);
		
		$.ajax({
			type: 'POST',
			url: '/services/proxy',
			headers: {
				'Authorization': 'Bearer ' + sessionId,
				'Content-Type': 'application/json',
				'SalesforceProxy-Endpoint': salesforceBaseUrl + '/services/apexrest/ReportData/' + dataDefinitionId
			},
			data: conditions,
			success: function(data) {
//				alert('[data] ok');
				if (originalJson) {
					save(template, reportName, data, originalJson);
				} else {
					getOriginalJsonAndSave(template, reportName, data, dataDefinitionId);
				}
			},
			error: function(jqXHR, textStatus, errorThrown) {
				alert('[data] ' + textStatus);
			}
		});
	}
	
	function getOriginalJsonAndSave(template, reportName, data, dataDefinitionId) {
		$.ajax({
			type: 'GET',
			url: '/services/proxy',
			headers: {
				'Authorization': 'Bearer ' + sessionId,
				'SalesforceProxy-Endpoint': salesforceBaseUrl + '/services/apexrest/ReportDataDefinition/' + dataDefinitionId
			},
			success: function(originalJson) {
//				alert('[dataDefinition] ok');
				console.log('[dataDefinition] originalJson=' + originalJson);
				save(template, reportName, data, JSON.parse(originalJson));
			},
			error: function(jqXHR, textStatus, errorThrown) {
				alert('[dataDefinition] ' + textStatus);
			}
		});
	}
	
	function save(template, reportName, data, originalJson) {
		let fileType= template.FileType__c;
		let t= template.Template__c;
		
		let Factory= ReportFactory[fileType];
		if (Factory == null) {
			alert('Not found: ReportFactory extension for ' + fileType);
			return;
		}
		
		let f= new Factory(t, originalJson);
		f.merge(data);
		f.save(reportName);
	}
	
})();
