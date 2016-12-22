function Editor() {}

(function() {

	//////////
	
	const apiVersion= 'v37.0';
	
	// cache
	const sobjectNameToData= {};
	
	let sessionId;
	
	// View settings
	// true -> Hide "Add Definition" and "Remove Definition" buttons.
	Editor.fixDataDefinitions= false;
	// true -> Replace the dropdown of object names by simple label.
	Editor.fixObjects= false;
	// true -> Show only selected child relationships and hide checkboxes.
	Editor.fixChildRelationships= false;
	// true -> Show fields.
	Editor.showFields= true;
	// true -> Show parent conditions.
	Editor.showParentConditions= true;
	
	// JSON model settings
	// true -> Include q.f and q.c[].f properties.
	Editor.includeObjects= true;
	// true -> Include q.c even if q.c is empty (q.c = []).
	Editor.includeChildRelationships= true;
	
	Editor.init= function(sessionId_, qArray_) {
		sessionId= sessionId_;
		let qArray= qArray_ || [{"f":""}];
		
		console.log('[Editor.init] sessionId=' + sessionId);
		
		if (!Editor.fixDataDefinitions) {
			$('.dataDefinitionRowRemoveRow').show();
			$('.dataDefinitionRowAddRow').show();
		}
		
		if (Editor.showFields) {
			$('.field').show();
			$('.childField').show();
		}
		
		if (Editor.includeChildRelationships) {
			$('.childRelationship').show();
		}
		
		Editor.initDataDefinitionRowAddButton();
		
		$.ajax({
			url: '/services/data/' + apiVersion + '/sobjects',
			headers: {
				'Authorization': 'Bearer ' + sessionId
			},
			success: function(data) {
				onSuccess(data);
			},
			error: function(jqXHR, textStatus, errorThrown) {
				console.log('[Editor] data error: status=' + textStatus);
			}
		});
		
		
		function onSuccess(data) {
			let option= getSelectOneOption();
			let objectName= $('.objectName');
			objectName.children().remove();
			objectName.on('change', onObjectChanged);
			objectName.append(option.clone());
			data.sobjects.forEach(function(sobject) {
				if (!sobject.deprecatedAndHidden && sobject.queryable && sobject.retrieveable && sobject.searchable) {
					objectName.append(createOption(option.clone(), sobject, false));
				}
			});
			
			if (qArray) {
				qArray.forEach(function(q) {
					let dataDefinitionRow= add($('.dataDefinition'));
					let objectName= dataDefinitionRow.find('.objectName');
					objectName.val(q.f);
					handleObjectChanged(objectName, q);
					
					if (Editor.fixObjects) {
						data.sobjects.forEach(function(sobject) {
							if (sobject.name === q.f) {
								let object= dataDefinitionRow.find('.objectRow');
								// Need more elegant way to replace the select element
								// and elements inserted by Visualforce.
								object.replaceWith(
									'<tr class="object">'
										+ '<td class="labelCol empty last"></td>'
										+ '<td>' + sobject.label + '</td>'
										+ '<td>'
											+ '<input type="hidden" class="objectName" value="' + sobject.name + '"/>'
										+ '</td>'
									+ '</tr>');
								objectName.remove();
							}
						});
					}
				});
			}
		}
	};
	
	function onObjectChanged() {
		handleObjectChanged($(this));
	}
	
	function handleObjectChanged(source, q) {
		let t= $(source || this);
		let sobjectName= t.val();
//		console.log('[handleObjectChanged] source=' + source + ', this=' + this + ', sobjectName=' + sobjectName);
		
		let dataDefinitionRow= t.closest('.dataDefinitionRow');
		initDataDefinitionRow(dataDefinitionRow);
		
		if (sobjectName === '__optSelectOne') {
			return;
		}
		
		let field= dataDefinitionRow.find('.field');
		let condition= dataDefinitionRow.find('.condition');
		let childRelationship= dataDefinitionRow.find('.childRelationship');
		
		updateFieldsConditionsAndChildRelationships(sobjectName, field, condition, childRelationship, q);
		
		if (Editor.showFields) {
			field.show();
		}
		if (Editor.showParentConditions) {
			condition.show();
		}
	}
	
	function onFieldChanged() {
		handleFieldChanged($(this));
	}
	
	function handleFieldChanged(source) {
		let t= $(source);
		let val= t.val();
		
		// nothing to do
	}
	
	function onConditionFieldChanged() {
		handleConditionFieldChanged($(this));
	}
	
	function handleConditionFieldChanged(source, w) {
		let t= $(source);
		let val= t.val();
		
		let type= val === '__optSelectOne' ? '__optNone' : val.substring(val.indexOf(':') + 1);
		
		let conditionFieldNameClass= t.attr('class');
		let conditionClass= conditionFieldNameClass.substring(0, conditionFieldNameClass.length - 'FieldName'.length);
		let conditionRowClass= conditionClass + 'Row';
		let conditionRow= t.closest('.' + conditionRowClass);
		
		let conditionOperatorClass= conditionClass + 'Operator';
		let conditionOperator= conditionRow.find('.' + conditionOperatorClass);
		conditionOperator.children().remove();
		conditionOperator.append(getOperatorOptions(type));
		
		let conditionValueClass= conditionClass + 'Value';
		let conditionValue= conditionRow.find('.' + conditionValueClass);
		
		if (w) {
			if (w.o) {
				conditionOperator.val(w.o);
			}
			
			if (w.v) {
				conditionValue.val(w.v);
			}
		}
	}
	
	function onChildObjectChanged() {
		handleChildObjectChanged($(this));
	}
	
	function handleChildObjectChanged(source, q) {
		let t= $(source);
		
		let childRelationshipRow= t.closest('.childRelationshipRow');
		let childFieldAndCondition= childRelationshipRow.find('.childFieldAndCondition');
		
		let relationshipName= childRelationshipRow.find('.relationshipName').text();
		let childSObject= childRelationshipRow.find('.relationshipChildSObject').text();
		
		if (t.prop('checked')) {
			if (t.data('state') != 'initialized') {
				let childField= childFieldAndCondition.find('.childField');
				let childFieldName= childField.find('.childFieldName')
				
				let childCondition= childFieldAndCondition.find('.childCondition');
				let childConditionFieldName= childCondition.find('.childConditionFieldName');
				
				updateFieldsConditionsAndChildRelationships(childSObject, childField, childCondition, null, q, relationshipName);
				
				t.data('state', 'initialized');
			}
		
			childFieldAndCondition.show();
		} else {
			childFieldAndCondition.hide();
		}
	}
	
	function updateFieldsConditionsAndChildRelationships(sobjectName, field, condition, childRelationship, q, relationshipName) {
		let data= sobjectNameToData[sobjectName];
		if (data) {
			onSuccess(data);
		} else {
			$.ajax({
				url: '/services/data/' + apiVersion + '/sobjects/' + sobjectName + '/describe',
				headers: {
					'Authorization': 'Bearer ' + sessionId
				},
				success: function(data) {
					sobjectNameToData[sobjectName]= data;
					onSuccess(data);
				},
				error: function(jqXHR, textStatus, errorThrown) {
					console.log('[updateFieldsConditionsAndChildRelationships] data error: status=' + textStatus);
				}
			});
		}
		
		
		function onSuccess(data) {
			let option= getSelectOneOption();
			
			let fieldName= field.find('.' + field.attr('class') + 'Name');
			fieldName.children().remove();
			fieldName.on('change', onFieldChanged);
			fieldName.append(option.clone());
			
			let conditionFieldName= condition.find('.' + condition.attr('class') + 'FieldName');
			conditionFieldName.children().remove();
			conditionFieldName.on('change', onConditionFieldChanged);
			conditionFieldName.append(option.clone());
			
			data.fields.forEach(function(field) {
				if (!field.deprecatedAndHidden) {
					fieldName.append(createOption(option.clone(), field));
					
					if (field.filterable) {
						conditionFieldName.append(createOption(option.clone(), field));
					}
				}
			});
			
			if (q) {
				if (q.s) {
					q.s.forEach(function(s, index) {
						let fieldRow= add(field);
						let fieldName= fieldRow.find('.' + field.attr('class') + 'Name');
						fieldName.val(s.f + ':' + getType(sobjectName, s.f));
//						fieldName.val(s);
						handleFieldChanged(fieldName, q);
					});
				}
				
				if (q.w) {
					q.w.forEach(function(w, index) {
						let conditionRow= add(condition);
						let conditionFieldName= conditionRow.find('.' + condition.attr('class') + 'FieldName');
						conditionFieldName.val(w.f + ':' + getType(sobjectName, w.f));
//						conditionFieldName.val(w.f + ':' + getType(sobjectName, w.f, relationshipName));
						handleConditionFieldChanged(conditionFieldName, w);
					});
				}
			} else {
				if (sobjectName && 0 < sobjectName.length) {
					showField= true;
					showCondition= true;
				}
			}
			
			if (childRelationship) {
				let exists= false;
				
				for (let i= 0, size= data.childRelationships.length; i < size; ++i) {
					let relationship= data.childRelationships[i];
					if (!relationship.deprecatedAndHidden && relationship.relationshipName != null) {
						let childRelationshipRow= add(childRelationship);
						let childObjectName= childRelationshipRow.find('.childObjectName');
						childObjectName.on('change', onChildObjectChanged);
						
						if (Editor.fixChildRelationships) {
							childObjectName.hide();
						}
						
						let relationshipName= childRelationshipRow.find('.relationshipName');
						relationshipName.text(relationship.relationshipName);
						let relationshipChildSObject= childRelationshipRow.find('.relationshipChildSObject');
						relationshipChildSObject.text(relationship.childSObject);
						
						if (q && q.c) {
							let matches= -1;
							
							for (let i= 0, size= q.c.length; i < size; ++i) {
								if (q.c[i].f == relationship.relationshipName) {
									matches= i;
									break;
								}
							}
							
							if (matches != -1) {
								childObjectName.prop('checked', true);
								handleChildObjectChanged(childObjectName, q.c[matches]);
							} else {
								if (Editor.fixChildRelationships) {
									childRelationshipRow.remove();
									continue;
								}
							}
						} else {
							if (Editor.fixChildRelationships) {
								childRelationshipRow.remove();
								continue;
							}
						}
						
						exists= true;
					}
				}
				
				if (exists && Editor.includeChildRelationships) {
					childRelationship.show();
				} else {
					childRelationship.hide();
				}
			}
		}
	}
	
	
	function getType(sobjectName, fieldName) {
		let data= sobjectNameToData[sobjectName];
		if (data) {
			let fields= data.fields;
			for (let i= 0, size= data.fields.length; i < size; ++i) {
				if (fields[i].name === fieldName) {
					return fields[i].type;
				}
			}
		}
		return null;
	}
	
	function createOption(option, field, includeType) {
		includeType= includeType !== false;
		
		if (field.idLookup) { // e.g. Account.Id, Contact.Id, Contact.Email
//			return options;
		}
		
		let name;
		let type;
		let label;
		
		switch (field.type) {
		case 'id': // ?
			name= field.name;
			type= field.type;
			label= field.label;// + ' => ' + field.name + '.Name[type="id"]';
			break;
		case 'reference': // Account: Parent Account ID => Parent.Name
			// TODO: WARNING!! There are objects without Name field.
			name= field.relationshipName + '.Name';
			type= 'string';
			label= name;
			break;
		default:
			name= field.name;
			type= field.type;
			label= field.label;
		}
		
		return option.val(includeType ? (name + ':' + type) : name).text(label);
	}
	
	
	//////////
	
	Editor.initDataDefinitionRowAddButton= function() {
		initButtons($('.dataDefinitionRowAddRow'));
	};
	
	function initDataDefinitionRow(dataDefinitionRow) {
		let field= dataDefinitionRow.find('.field');
		field.hide();
		removeAll(field);
		
		let condition= dataDefinitionRow.find('.condition');
		condition.hide();
		removeAll(condition);
		
		let childRelationship= dataDefinitionRow.find('.childRelationship');
		childRelationship.hide();
		removeAll(childRelationship);
	}
	
	function add(container) {
		let containerClass= container.attr('class');
//		console.log('[add] containerClass=' + containerClass);
		let rowClass= containerClass + 'Row';
		let rowTemplateClass= rowClass + 'Template';
		let rowAddRowClass= rowClass + 'AddRow';
		
		let rowTemplate= container.find('.' + rowTemplateClass);
		let rowAddRow= container.find('.' + rowAddRowClass);
		
		let row= rowTemplate.clone(true).attr('class', rowClass);
		initButtons(row);
		container.append(row).append(rowAddRow);
		row.show();
		return row;
	}
	
	function removeAll(container) {
		let containerClass= container.attr('class');
//		console.log('[removeAll] containerClass=' + containerClass);
		let rowClass= containerClass + 'Row';
		let row= container.find('.' + rowClass);
		row.remove();
	}
	
	function getSelectOneOption() {
		let operatorOptionsTemplates= $('.operatorOptionsTemplates');
		let select= operatorOptionsTemplates.find('.__optSelectOne');
		return select.children().clone();
	}
	
	function getOperatorOptions(type) {
		let operatorOptionsTemplates= $('.operatorOptionsTemplates');
		let select= operatorOptionsTemplates.find('.' + type);
		return select.children().clone();
	}
	
	
	Editor.getResult= function() {
		let qArray= [];
		
		$('.dataDefinitionRow').each(function() {
			let dataDefinitionRow= $(this);
			let objectName= dataDefinitionRow.find('.objectName').val();
			if (objectName !== '__optSelectOne') {
				addQ(qArray, dataDefinitionRow);
			}
		});
		
		return qArray;
		
		
		function addQ(qArray, dataDefinitionRow) {
			let q= {};
			
			let s= [];
			dataDefinitionRow.find('.fieldRow').each(function() {
				let fieldRow= $(this);
				let fieldNameValue= fieldRow.find('.fieldName').val();
				if (fieldNameValue !== '__optSelectOne') {
					let si= {};
					let colon= fieldNameValue.indexOf(':');
					si.f= fieldNameValue.substring(0, colon);
					si.t= fieldNameValue.substring(colon + 1);
					s.push(si);
				}
			});
			if (s.length == 0) {
				s.push({"f":"Id","t":"id"});
			}
			q.s= s;
			
			let c= [];
			dataDefinitionRow.find('.childRelationshipRow').each(function() {
				let childRelationshipRow= $(this);
				if (childRelationshipRow.find('.childObjectName').prop('checked')) {
					addCQ(c, childRelationshipRow);
				}
			});
			if (0 < c.length || Editor.includeChildRelationships) {
				q.c= c;
			}
			
			let object= dataDefinitionRow.find('.objectName').val();
			if (Editor.includeObjects) {
				q.f= object;
			}
			
			let w= [];
			dataDefinitionRow.find('.conditionRow').each(function() {
				let conditionRow= $(this);
				let conditionFieldNameValue= conditionRow.find('.conditionFieldName').val();
				let conditionOperatorValue= conditionRow.find('.conditionOperator').val();
				if (conditionFieldNameValue !== '__optSelectOne' && conditionOperatorValue !== '__optNone') {
					let wi= {};
					let colon= conditionFieldNameValue.indexOf(':');
					wi.f= conditionFieldNameValue.substring(0, colon);
					wi.t= conditionFieldNameValue.substring(colon + 1);
					wi.o= conditionOperatorValue;
					wi.v= conditionRow.find('.conditionValue').val();
					w.push(wi);
				}
			});
			if (0 < w.length) {
				q.w= w;
			}
			
			qArray.push(q);
		}
		
		function addCQ(qArray, childRelationshipRow) {
			let q= {};
			
			let s= [];
			childRelationshipRow.find('.childFieldRow').each(function() {
				let fieldRow= $(this);
				let fieldNameValue= fieldRow.find('.childFieldName').val();
				if (fieldNameValue !== '__optSelectOne') {
					let si= {};
					let colon= fieldNameValue.indexOf(':');
					si.f= fieldNameValue.substring(0, colon);
					si.t= fieldNameValue.substring(colon + 1);
					s.push(si);
				}
			});
			if (s.length == 0) {
				s.push({"f":"Id","t":"id"});
			}
			q.s= s;
			
			let object= childRelationshipRow.find('.relationshipName').text();
			if (Editor.includeObjects) {
				q.f= object;
			}
			
			let w= [];
			childRelationshipRow.find('.childConditionRow').each(function() {
				let conditionRow= $(this);
				let conditionFieldNameValue= conditionRow.find('.childConditionFieldName').val();
				let conditionOperatorValue= conditionRow.find('.childConditionOperator').val();
				if (conditionFieldNameValue !== '__optSelectOne' && conditionOperatorValue !== '__optNone') {
					let wi= {};
					let colon= conditionFieldNameValue.indexOf(':');
					wi.f= conditionFieldNameValue.substring(0, colon);
					wi.t= conditionFieldNameValue.substring(colon + 1);
					wi.o= conditionOperatorValue;
					wi.v= conditionRow.find('.childConditionValue').val();
					w.push(wi);
				}
			});
			if (0 < w.length) {
				q.w= w;
			}
			
			qArray.push(q);
		}
	}
	
	function initButtons(element) {
		$(element).find('.btn').each(function() {
			let btn= $(this);
			let parentClass= btn.parent().attr('class');
			let row= parentClass.lastIndexOf('Row');
			let action= parentClass.substring(row + 'Row'.length);
			btn.on('click', function() {
				Actions[action](btn);
			});
		});
	}
	
	
	Actions= {};
	
	Actions.Add= function(button) {
		let rowAdd= $(button).parent();
		let rowAddClass= rowAdd.attr('class');
		let containerClass= rowAddClass.substring(0, rowAddClass.length - 'RowAdd'.length);
		let container= rowAdd.closest('.' + containerClass);
		add(container);
	};
	
	Actions.Remove= function(button) {
		let rowRemove= $(button).parent();
		let rowRemoveClass= rowRemove.attr('class');
		let rowClass= rowRemoveClass.substring(0, rowRemoveClass.length - 'Remove'.length);
		let row= rowRemove.closest('.' + rowClass);
		row.remove();
	};
	
})();