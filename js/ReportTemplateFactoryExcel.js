// Auxiliary constants definitions

(function(openXml) {
	
	const XNamespace= Ltxml.XNamespace;
	const XName= Ltxml.XName;
	
	// /xl/xmlMaps.xml
	openXml.NoNamespace.ID= new XName('ID');
	openXml.NoNamespace.Name= new XName('Name');
	openXml.NoNamespace.RootElement= new XName('RootElement');
	openXml.NoNamespace.SchemaID= new XName('SchemaID');
	openXml.NoNamespace.ShowImportExportValidationErrors= new XName('ShowImportExportValidationErrors');
	openXml.NoNamespace.AutoFit= new XName('AutoFit');
	openXml.NoNamespace.Append= new XName('Append');
	openXml.NoNamespace.PreserveSortAFLayout= new XName('PreserveSortAFLayout');
	openXml.NoNamespace.PreserveFormat= new XName('PreserveFormat');
	
	let sNs= openXml.sNs;
	openXml.S.Map= new XName(sNs, "Map");
	openXml.S.Schema= new XName(sNs, "Schema");
	
	openXml.NoNamespace.minOccurs= new XName('minOccurs');
	openXml.NoNamespace.maxOccurs= new XName('maxOccurs');
	
	openXml.xsdNs= new XNamespace('http://www.w3.org/2001/XMLSchema');
	let xsdNs= openXml.xsdNs;
	openXml.XSD= {
		all: new XName(xsdNs, 'all'),
		complexType: new XName(xsdNs, 'complexType'),
		element: new XName(xsdNs, 'element'),
		schema: new XName(xsdNs, 'schema'),
		sequence: new XName(xsdNs, 'sequence'),
	};
	
	// /xl/worksheets/sheet#.xml
	openXml.NoNamespace.tabSelected= new XName('tabSelected');
	
	// /xl/tables/table#.xml
	openXml.NoNamespace.tableType= new XName('tableType');
	openXml.NoNamespace.insertRow= new XName('insertRow');
	
})(openXml);


// ReportTemplateFactory.Excel

if (!ReportTemplateFactory) {
	function ReportTemplateFactory() {}
}

(function() {
		
	ReportTemplateFactory.Excel= function() {};
	
	ReportTemplateFactory.Excel.create= function(queryArray) {
//		alert('[ReportTemplateFactory.Excel.create] called.');
		
		const XDocument= Ltxml.XDocument;
		const XElement= Ltxml.XElement;
		const XAttribute= Ltxml.XAttribute;
		
		// The BASE64-encoded Excel template base file from which templates will be generated.
		// This Excel template base file can be opened by Excel application. 
		const templateBase
= "UEsDBBQABgAIAAAAIQCkU8XPTgEAAAgEAAATAAgCW0NvbnRlbnRfVHlwZXNdLnhtbCCiBAIo"
+ "oAACAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACsk8tOwzAQRfdI/EPkLYrdskAINe2CxxK6"
+ "KB9g4kli1S953NL+PRP3sUChFWo3sWLP3HM9M57MNtYUa4iovavYmI9YAa72Sru2Yp+Lt/KR"
+ "FZikU9J4BxXbArLZ9PZmstgGwIKyHVasSyk8CYF1B1Yi9wEcnTQ+WpnoN7YiyHopWxD3o9GD"
+ "qL1L4FKZeg02nbxAI1cmFa8b2t45iWCQFc+7wJ5VMRmC0bVM5FSsnfpFKfcETpk5Bjsd8I5s"
+ "MDFI6E/+BuzzPqg0USso5jKmd2nJhtgY8e3j8sv7JT8tMuDSN42uQfl6ZakCHEMEqbADSNbw"
+ "vHIrtTv4PsHPwSjyMr6ykf5+WfiMj0T9BpG/l1vIMmeAmLYG8Nplz6KnyNSvefQBaXIj/J9+"
+ "GM0+uwwkBDFpOA7nUJOPRJr6i68L/btSoAbYIr/j6Q8AAAD//wMAUEsDBBQABgAIAAAAIQC1"
+ "VTAj9AAAAEwCAAALAAgCX3JlbHMvLnJlbHMgogQCKKAAAgAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA"
+ "AAAAAAAArJJNT8MwDIbvSPyHyPfV3ZAQQkt3QUi7IVR+gEncD7WNoyQb3b8nHBBUGoMDR3+9"
+ "fvzK2908jerIIfbiNKyLEhQ7I7Z3rYaX+nF1ByomcpZGcazhxBF21fXV9plHSnkodr2PKqu4"
+ "qKFLyd8jRtPxRLEQzy5XGgkTpRyGFj2ZgVrGTVneYviuAdVCU+2thrC3N6Dqk8+bf9eWpukN"
+ "P4g5TOzSmRXIc2Jn2a58yGwh9fkaVVNoOWmwYp5yOiJ5X2RswPNEm78T/XwtTpzIUiI0Evgy"
+ "z0fHJaD1f1q0NPHLnXnENwnDq8jwyYKLH6jeAQAA//8DAFBLAwQUAAYACAAAACEATiMbaPIA"
+ "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"
+ "AHoCAABjHgAAAAA=";
		
		return generate(name, queryArray, templateBase);
		
		
		function generate(name, queryArray, templateBase) {
			let template= new openXml.OpenXmlPackage(templateBase);
			
			generateAll(queryArray, template);
			
			return template.saveToBase64();
		}
		
		function generateAll(queryArray, pkg) {
			const xsdPrefix= 'xsd';
			
			let workbookPart= pkg.workbookPart();
			let customXmlMappingsPart= workbookPart.getPartByRelationshipType(openXml.relationshipTypes.customXmlMappings);
			let customXmlMappingsPartXDoc= customXmlMappingsPart.getXDocument();
			let MapInfo= customXmlMappingsPartXDoc.root;
			
			MapInfo.removeNodes();
			
			///// BEGIN Layout tables /////
			let workbookStylesPart= workbookPart.workbookStylesPart();
			let workbookStylesPartXDoc= workbookStylesPart.getXDocument();
			let defaultTableStyle= workbookStylesPartXDoc.root.element(openXml.S.tableStyles).attribute(openXml.NoNamespace.defaultTableStyle).value;
			
			let workbookContext= {
				// array of the shared strings
				sharedStrings: generateSharedStrings(queryArray, pkg),
				
				defaultTableStyle: defaultTableStyle,
				// assigned to singleXmlCell/@id and table/@id
				id: 1,
				// assigned to /xl/tables/table<tableId>.xml
				tableId: 1,
				// assigned to /xl/tables/tableSingleCells<tableSingleCellsId>.xml
				tableSingleCellsId: 1
			};
			
			// /docProps/app.xml
			let extendedFilePropertiesPart= pkg.extendedFilePropertiesPart();
			let extendedFilePropertiesPartXDoc= extendedFilePropertiesPart.getXDocument();
			let EP_Properties= extendedFilePropertiesPartXDoc.root;
			let VT_i4= EP_Properties.element(openXml.EP.HeadingPairs).element(openXml.VT.vector).elements(openXml.VT.variant).elementAt(1).element(openXml.VT.i4);
			VT_i4.setValue(String(queryArray.length));
			let VT_vector= EP_Properties.element(openXml.EP.TitlesOfParts).element(openXml.VT.vector);
			VT_vector.removeNodes();
			VT_vector.setAttributeValue(openXml.NoNamespace.size, String(queryArray.length));
			
			// /xl/workbook.xml
			let workbookPartXDoc= workbookPart.getXDocument();
			let S_sheets= workbookPartXDoc.root.element(openXml.S.sheets);
			S_sheets.removeNodes();
			
			// /xl/worksheets/sheet#.xml
			let worksheetPart= workbookPart.worksheetParts()[0];
			let worksheetPartXDoc= worksheetPart.getXDocument();
			pkg.deletePart(worksheetPart);
			workbookPart.getRelationshipsByRelationshipType(openXml.relationshipTypes.worksheet).forEach(function(r) {
				workbookPart.deleteRelationship(r.relationshipId);
			});
			///// END Layout tables /////
			
			queryArray.forEach(function(q, index) {
				let i= index + 1;
				let MapID= String(i);
				let MapName= 'Root_Map' + MapID;
				let SchemaID= 'Schema' + MapID;
				let SchemaRoot= 'Root';
				
				///// BEGIN Layout tables /////
				let Sheet= 'Sheet' + i;
				
				// /docProps/app.xml
				VT_vector.add(
					new XElement(openXml.VT.lpstr, Sheet)
				);
				
				let rId= 'rId' + (i + 1000);
				
				// /xl/workbook.xml
				S_sheets.add(
					new XElement(openXml.S.sheet,
						new XAttribute(openXml.NoNamespace.name, Sheet),
						new XAttribute(openXml.NoNamespace.sheetId, String(i)),
						new XAttribute(openXml.R.id, rId)
					)
				);
				
				// /xl/worksheets/sheet#.xml
				let sheetXml= 'sheet' + i + '.xml';
				
				// deep clone
				let newWorksheet= new XElement(worksheetPartXDoc.root);
				let sheetView= newWorksheet.element(openXml.S.sheetViews).element(openXml.S.sheetView);
				if (index == 0) {
					sheetView.setAttributeValue(openXml.NoNamespace.tabSelected, '1');
				} else {
					let tabSelected= sheetView.attribute(openXml.NoNamespace.tabSelected);
					if (tabSelected) {
						tabSelected.remove();
					}
				}
				let selection= sheetView.element(openXml.S.selection);
				if (selection) {
					selection.remove();
				}
				
				let newWorksheetPart= pkg.addPart('/xl/worksheets/' + sheetXml, openXml.contentTypes.worksheet, 'xml', newWorksheet);
				workbookPart.addRelationship(rId, openXml.relationshipTypes.worksheet, 'worksheets/' + sheetXml, 'Internal');
				
				let worksheetContext= {
					workbookContext: workbookContext,
					worksheetPart: newWorksheetPart,
					offset: {
						r: 1,
						c: 1
					}
				};
				///// END Layout tables /////
				
				let sequence= new XElement(openXml.XSD.sequence);
				
				if (q.c) {
					// with child relationships >> single parent record
					generateTableSingleCells(q, sequence, MapID, worksheetContext);
					q.c.forEach(function(q) {
						generateTable(q, sequence, MapID, worksheetContext);
					});
				} else {
					// without child relationships >> multiple records
					generateTable(q, sequence, MapID, worksheetContext);
				}
				
				MapInfo.add(
					new XElement(openXml.S.Schema,
						new XElement(openXml.XSD.schema,
							new XElement(openXml.XSD.element,
								new XElement(openXml.XSD.complexType, sequence),
								new XAttribute(openXml.NoNamespace.name, SchemaRoot)
							),
							new XAttribute('xmlns:' + xsdPrefix, openXml.xsdNs)
						),
						new XAttribute(openXml.NoNamespace.ID, SchemaID)
					)
				);
				
				MapInfo.add(
					new XElement(openXml.S.Map,
						new XAttribute(openXml.NoNamespace.ID, MapID),
						new XAttribute(openXml.NoNamespace.Name, MapName),
						new XAttribute(openXml.NoNamespace.RootElement, SchemaRoot),
						new XAttribute(openXml.NoNamespace.SchemaID, SchemaID),
						new XAttribute(openXml.NoNamespace.ShowImportExportValidationErrors, 'false'),
						new XAttribute(openXml.NoNamespace.AutoFit, 'true'),
						new XAttribute(openXml.NoNamespace.Append, 'false'),
						new XAttribute(openXml.NoNamespace.PreserveSortAFLayout, 'true'),
						new XAttribute(openXml.NoNamespace.PreserveFormat, 'true')
					)
				);
			});
			
			
			function generateSharedStrings(queryArray, pkg) {
				let sharedStrings= [];
				
				queryArray.forEach(function(q) {
					registerFieldNames(q);
				});
				
				let rId= 'rId4001'; // TODO
				let sst= new XElement(openXml.S.sst,
					new XAttribute(openXml.NoNamespace.uniqueCount, String(sharedStrings.length))
				);
				sharedStrings.forEach(function(ss) {
					sst.add(
						new XElement(openXml.S.si,
							new XElement(openXml.S.t, ss)
						)
					);
				});
				pkg.addPart('/xl/sharedStrings.xml', openXml.contentTypes.sharedStringTable, 'xml', sst);
				pkg.workbookPart().addRelationship(rId, openXml.relationshipTypes.sharedStringTable, 'sharedStrings.xml', 'Internal');
				
				return sharedStrings;
				
				
				function registerFieldNames(q) {
					if (!q.s || q.s.length == 0) {
						q.s= [{ "f":"ID", "t":"id" }];
					}
					
					q.s.forEach(function(s) {
						if (sharedStrings.indexOf(s.f) == -1) {
							sharedStrings.push(s.f);
						}
					});
					
					if (q.c) {
						q.c.forEach(function(c) {
							registerFieldNames(c);
						});
					}
				}
			}
			
			function generateTable(q, sequence/* Layout tables ->*/, mapId, worksheetContext) {
				let innerSequence= new XElement(openXml.XSD.sequence);
				
				///// BEGIN Layout tables /////
				// /xl/tables/table#.xml
				let workbookContext= worksheetContext.workbookContext;
				let worksheetPart= worksheetContext.worksheetPart;
				let offset= worksheetContext.offset;
				
				let name= 'Table' + workbookContext.id;
				let displayName= name;
				let r1= new spreadsheet.r(offset.r, offset.c);
				let r2= new spreadsheet.r(offset.r + 1, offset.c + q.s.length - 1);
				let ref= new spreadsheet.ref(r1, r2);
				
				let tableColumns= new XElement(openXml.S.tableColumns,
					new XAttribute(openXml.NoNamespace.count, String(q.s.length))
				);
				
				// /xl/worksheets/sheet#.xml
				let worksheet= worksheetPart.getXDocument();
				// TODO cannot assume A1 in the non zero offset cases.
				worksheet.element(openXml.S.dimension).setAttributeValue(openXml.NoNamespace._ref, 'A1:' + r2.toString());
				
				let sheetData= worksheet.element(openXml.S.sheetData);
				let spans= String(offset.c) + ':' + (offset.c + q.s.length - 1);
				// TODO Cannot simply append row. All row/@r must be sorted.
				let headRow= new XElement(openXml.S.row,
					new XAttribute(openXml.NoNamespace.r, String(offset.r)),
					new XAttribute(openXml.NoNamespace.spans, spans)
				);
				let dataRow= new XElement(openXml.S.row,
					new XAttribute(openXml.NoNamespace.r, String(offset.r + 1)),
					new XAttribute(openXml.NoNamespace.spans, spans)
				);
				sheetData.add(headRow);
				sheetData.add(dataRow);
				
				let rId= 'rId' + (workbookContext.tableId + 3000);
				let tableParts= worksheet.element(openXml.S.tableParts);
				if (tableParts) {
					let count= Number(tableParts.attribute(openXml.NoNamespace.count).value);
					tableParts.setAttributeValue(openXml.NoNamespace.count, String(count + 1));
				} else {
					tableParts= new XElement(openXml.S.tableParts,
						new XAttribute(openXml.NoNamespace.count, '1')
					);
					worksheet.add(tableParts);
				}
				tableParts.add(
					new XElement(openXml.S.tablePart,
						new XAttribute(openXml.R.id, rId)
					)
				);
				///// END Layout tables /////
				
				q.s.forEach(function(s/* Layout tables ->*/, index) {
					innerSequence.add(
						new XElement(openXml.XSD.element,
							new XAttribute(openXml.NoNamespace.name, s.f),
							new XAttribute(openXml.NoNamespace.type, xsdPrefix + ':' + dbTypeToXsdType(s.t))
						)
					);
					
					///// BEGIN Layout tables /////
					// /xl/worksheets/sheet#.xml
					let headRef= spreadsheet.r.toString(offset.r, offset.c + index);
//					let dataRef= spreadsheet.r.toString(offset.r + 1, offset.c + index);
					
					// TODO Cannot simply append c. All c/@r must be sorted.
					headRow.add(
						new XElement(openXml.S.c,
							new XElement(openXml.S.v, String(workbookContext.sharedStrings.indexOf(s.f))),
							new XAttribute(openXml.NoNamespace.r, headRef),
							new XAttribute(openXml.NoNamespace.t, 's')
						)
					);
					/*
					dataRow.add(
						new XElement(openXml.S.c,
							new XAttribute(openXml.NoNamespace.r, dataRef)
						)
					);*/
					
					// /xl/tables/table#.xml
					tableColumns.add(
						new XElement(openXml.S.tableColumn,
							new XElement(openXml.S.xmlColumnPr,
								new XAttribute(openXml.NoNamespace.mapId, mapId),
								new XAttribute(openXml.NoNamespace.xpath, '/Root/' + q.f + '/' + s.f),
								new XAttribute(openXml.NoNamespace.xmlDataType, dbTypeToXsdType(s.t))
							),
							new XAttribute(openXml.NoNamespace.id, String(index + 1)),
							new XAttribute(openXml.NoNamespace.uniqueName, s.f),
							new XAttribute(openXml.NoNamespace.name, s.f)
						)
					);
					///// END Layout tables /////
				});
				
				sequence.add(
					new XElement(openXml.XSD.element,
						new XElement(openXml.XSD.complexType, innerSequence),
						new XAttribute(openXml.NoNamespace.minOccurs, '0'),
						new XAttribute(openXml.NoNamespace.maxOccurs, 'unbounded'),
						new XAttribute(openXml.NoNamespace.name, q.f)
					)
				);
				
				///// BEGIN Layout tables /////
				let table= new XElement(openXml.S.table,
					new XElement(openXml.S.autoFilter,
						new XAttribute(openXml.NoNamespace._ref, ref.toString())
					),
					tableColumns,
					new XElement(openXml.S.tableStyleInfo,
						new XAttribute(openXml.NoNamespace.name, workbookContext.defaultTableStyle),
						new XAttribute(openXml.NoNamespace.showFirstColumn, '0'),
						new XAttribute(openXml.NoNamespace.showLastColumn, '0'),
						new XAttribute(openXml.NoNamespace.showRowStripes, '1'),
						new XAttribute(openXml.NoNamespace.showColumnStripes, '0')
					),
					new XAttribute(openXml.NoNamespace.id, String(workbookContext.id)),
					new XAttribute(openXml.NoNamespace.name, name),
					new XAttribute(openXml.NoNamespace.displayName, displayName),
					new XAttribute(openXml.NoNamespace._ref, ref.toString()),
					new XAttribute(openXml.NoNamespace.tableType, 'xml'),
					new XAttribute(openXml.NoNamespace.insertRow, '1'),
					new XAttribute(openXml.NoNamespace.totalsRowShown, '0')
				);
				
				let tableXml= 'table' + workbookContext.tableId + '.xml';
				pkg.addPart('/xl/tables/' + tableXml, openXml.contentTypes.tableDefinition, 'xml', table);
				worksheetPart.addRelationship(rId, openXml.relationshipTypes.tableDefinition, '../tables/' + tableXml, 'Internal');
				
				++workbookContext.id;
				++workbookContext.tableId;
				offset.r+= 2;
				///// END Layout tables /////
			}
			
			function generateTableSingleCells(q, sequence/* Layout tableSingleCells ->*/, mapId, worksheetContext) {
				let all= new XElement(openXml.XSD.all);
				
				///// BEGIN Layout tableSingleCells /////
				let workbookContext= worksheetContext.workbookContext;
				let worksheetPart= worksheetContext.worksheetPart;
				let offset= worksheetContext.offset;
				
				let singleXmlCells= new XElement(openXml.S.singleXmlCells);
				
				let rId= 'rId' + (workbookContext.tableSingleCellsId + 2000);
				let tableSingleCellsXml= 'tableSingleCells' + workbookContext.tableSingleCellsId + '.xml';
				pkg.addPart('/xl/tables/' + tableSingleCellsXml, openXml.contentTypes.singleCellTable, 'xml', singleXmlCells);
				worksheetPart.addRelationship(rId, openXml.relationshipTypes.singleCellTable, '../tables/' + tableSingleCellsXml, 'Internal');
				///// END Layout tableSingleCells /////
				
				q.s.forEach(function(s/* Layout tables ->*/, index) {
					all.add(
						new XElement(openXml.XSD.element,
							new XAttribute(openXml.NoNamespace.minOccurs, '0'),
							new XAttribute(openXml.NoNamespace.maxOccurs, '1'),
							new XAttribute(openXml.NoNamespace.name, s.f)
						)
					);
					
					///// BEGIN Layout tableSingleCells /////
					let r= spreadsheet.r.toString(offset.r, offset.c + 1);
					singleXmlCells.add(
						new XElement(openXml.S.singleXmlCell,
							new XElement(openXml.S.xmlCellPr,
								new XElement(openXml.S.xmlPr,
									new XAttribute(openXml.NoNamespace.mapId, mapId),
									new XAttribute(openXml.NoNamespace.xpath, '/Root/' + q.f + '/' + s.f),
									new XAttribute(openXml.NoNamespace.xmlDataType, dbTypeToXsdType(s.t))
								),
								new XAttribute(openXml.NoNamespace.id, '1'),
								new XAttribute(openXml.NoNamespace.uniqueName, s.f)
							),
							new XAttribute(openXml.NoNamespace.id, String(workbookContext.id)),
							new XAttribute(openXml.NoNamespace.r, r),
							new XAttribute(openXml.NoNamespace.connectionId, '0')
						)
					);
				
					let worksheet= worksheetPart.getXDocument();
					worksheet.element(openXml.S.dimension).setAttributeValue(openXml.NoNamespace._ref, 'A1:' + r);
					
					// /xl/worksheets/sheet#.xml
					let l= spreadsheet.r.toString(offset.r, offset.c);
					let spans= String(offset.c) + ':' + (offset.c + 1);
					// TODO cannot simply add newly created <row> or <c>.
					worksheet.element(openXml.S.sheetData).add(
						new XElement(openXml.S.row,
							new XElement(openXml.S.c,
								new XElement(openXml.S.v, String(workbookContext.sharedStrings.indexOf(s.f))),
								new XAttribute(openXml.NoNamespace.r, l),
								new XAttribute(openXml.NoNamespace.t, 's')
							),
							new XElement(openXml.S.c,
								new XAttribute(openXml.NoNamespace.r, r)
							),
							new XAttribute(openXml.NoNamespace.r, String(offset.r)),
							new XAttribute(openXml.NoNamespace.spans, spans)
						)
					);
					
					++workbookContext.id;
					++workbookContext.tableSingleCellsId;
					++offset.r;
					///// END Layout tableSingleCells /////
				});
				
				sequence.add(
					new XElement(openXml.XSD.element,
						new XElement(openXml.XSD.complexType, all),
						new XAttribute(openXml.NoNamespace.minOccurs, '0'),
						new XAttribute(openXml.NoNamespace.maxOccurs, '1'),
						new XAttribute(openXml.NoNamespace.name, q.f)
					)
				);
			}
		}
	};
	
	
	// TODO: Not confirmed.
	// Database type to XML Schema type mapping.
	// The dbType argument is SFDC type in the Salesforce case.
	function dbTypeToXsdType(dbType) {
		switch (dbType) {
		// address
		case 'address':
			return 'string';
		// boolean
		case 'boolean':
			return 'boolean';
		// currency date datetime double int
		case 'currency':
		case 'date':
		case 'datetime':
			return 'string';
		case 'double':
			return 'double';
		case 'int':
			return 'int';
		// email phone picklist string textarea url
		case 'email':
		case 'phone':
		case 'picklist':
		case 'string':
		case 'textarea':
		case 'url':
			return 'string';
		// id
		case 'id':
			return 'string';
		// reference
		case 'reference':
			return 'string';
		default:
			return 'string';
		}
	}
	
})();