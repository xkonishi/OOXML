if (!ReportTemplateFactory) {
	function ReportTemplateFactory() {}
}

(function() {
	
	let sessionId;
	let salesforceBaseUrl;
	
	ReportTemplateFactory.init= function(sessionId_, salesforceBaseUrl_) {
		sessionId= sessionId_;
		salesforceBaseUrl= salesforceBaseUrl_;
	};
	
	ReportTemplateFactory.create= function(title, dataDefinitionName, fileType) {
		let factory= ReportTemplateFactory[fileType];
		if (!factory) {
			alert('Not supported file type: ' + fileType);
			return;
		}
		
		$.ajax({
			type: 'POST',
			url: '/services/proxy',
			headers: {
				'Authorization': 'Bearer ' + sessionId,
				'SalesforceProxy-Endpoint': salesforceBaseUrl + '/services/apexrest/ReportDataDefinition/'
			},
			data: {
				dataDefinitionName: dataDefinitionName
			},
			success: function(queryArrayAsString) {
	//			alert('[queryArray] success: ' + data);
				let template= factory.create(JSON.parse(queryArrayAsString));
				uploadTemplate(dataDefinitionName, title, fileType, template);
				/* Save the template locally.
				new openXml.OpenXmlPackage(template).saveToBlobAsync(function(blob) {
					saveAs(blob, templateName);
				});
				*/
			},
			error: function(jqXHR, textStatus, errorThrown) {
				alert('[queryArray] ' + textStatus);
			}
		});
	};
	
	
	function uploadTemplate(dataDefinitionName, title, fileType, template) {
		let fileSuffix= fileType == 'Excel' ? '.xlsx' : '.docx';
		let templateName= title + fileSuffix;
		
		$.ajax({
			type: 'POST',
			url: '/services/proxy',
			headers: {
				'Authorization': 'Bearer ' + sessionId,
				'SalesforceProxy-Endpoint': salesforceBaseUrl + '/services/apexrest/ReportTemplate/'
			},
			data: {
				dataDefinitionName: dataDefinitionName,
				title: title,
				fileType: fileType,
				template: template
			},
			success: function(templateId) {
	//			alert('[template] success: templateId=' + templateId);
				redirectToTemplateDetailPage(templateId);
			},
			error: function(jqXHR, textStatus, errorThrown) {
				alert('[template] ' + textStatus);
			}
		});
	}
	
	function redirectToTemplateDetailPage(templateId) {
		$('[id$=":templateId"]').val(templateId);
		redirect();
	}
	
})();