if (!spreadsheet) {
	function spreadsheet() {}
}

(function() {
	
	let XElement= Ltxml.XElement;
	let XAttribute= Ltxml.XAttribute;
	
	
	/**
	 * /_rels/.rels
	 * /Relationships/@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
	 * -> @Target='xl/workbook.xml'
	 * pkg.workbookPart();
	 * 
	 * /xl/workbook.xml
	 * /workbook/sheets/sheet/@sheetId
	 * index = @sheetId - 1
	 *	 for queryArray[index]
	 * -> @r:id
	 * 
	 * [/xl/]_rels/workbook.xml.rels
	 * -> /Relationships/Relationship/@Id
	 *   -> @Target='worksheets/sheet#.xml
	 * 
	 * [/xl/]worksheets/sheet#.xml
	 * /worksheet/tableParts/tablePart/@r:id <- ignore them
	 * 
	 * [/xl/worksheets/]_rels/sheet#.xml.rels
	 * -> /Relationships/Relationship/@Id
	 *   -> @Target='../tables/table?.xml'
	 *   -> @Target='../tables/tableSingleCells?.xml'
	 * 
	 * [/xl/]tables/table?.xml
	 * /table/@id
	 * index = /table/@id - 1
	 *	 for (q|c)[index]
	 *   /table/tableColumns/tableColumn/@id
	 * index = /table/tableColumns/tableColumn/@id - 1
	 *	 for s[index]
	 *   /table/@ref='X#:Y#'
	 *	   columnIndex = aToN('X') + index
	 * 
	 * [/xl/]tables/tableSingleCells?.xml
	 * /singleXmlCells/singleXmlCell/@id
	 * index = /singleXmlCells/singleXmlCell/@id - 1
	 *	 for s[index]
	 *   /singleXmlCells/singleXmlCell/@r='X#'
	 */
	spreadsheet.Layout= function(pkg) {
		this.keyToTable= {};
		this.keyToSingleXmlCell= {};
		
		let sheetLayouts= {};
		this.sheetLayouts= sheetLayouts;
		
		// [table|sigleXmlCell]/@id => @sheetId
		let id2sheetId= {};
		this.id2sheetId= id2sheetId;
		
		// "<mapId>/<sobjectName|relationshipName>" => table info
		// "<mapId>/<sobjectName|relationshipName>/<fieldName>" => singleXmlCell info
		let keyToIdElement= {};
		this.keyToIdElement= keyToIdElement;
		
		let workbookPart= pkg.workbookPart();
		
		let customXmlMappingsPart= workbookPart.getPartByRelationshipType(openXml.relationshipTypes.customXmlMappings);
		let customXmlMappingsPartXDoc= customXmlMappingsPart.getXDocument();
		let MapInfo= customXmlMappingsPartXDoc.root;
		
		let workbookPartXDoc= workbookPart.getXDocument();
		workbookPartXDoc.root.element(openXml.S.sheets).elements(openXml.S.sheet).forEach(function(sheet) {
			let sheetId= Number(sheet.attribute(openXml.NoNamespace.sheetId).value);
			let rId= sheet.attribute(openXml.R.id).value;
			let worksheetPart= workbookPart.getPartById(rId);
			let worksheet= worksheetPart.getXDocument().root;
			let dimension= worksheet.element(openXml.S.dimension);
			let sheetData= worksheet.element(openXml.S.sheetData);
			
			// table/@id => table/@ref or sigleXmlCell/@id => singleXmlCell/@r
			let refs= {};
			// table/@id => table element
			let tables= {};
			
			worksheetPart.tableDefinitionParts().forEach(function(tableDefinitionPart) {
				let tableDefinitionPartXDoc= tableDefinitionPart.getXDocument();
				let table= tableDefinitionPartXDoc.root
				
				let id= Number(table.attribute(openXml.NoNamespace.id).value);
				id2sheetId[id]= sheetId;
				
				let ref= table.attribute(openXml.NoNamespace._ref).value;
//				console.log('[spreadsheet.Layout] table[@id="' + id + '"]/@ref="' + ref + '"');
				refs[id]= ref;
				tables[id]= table;
				
				table.element(openXml.S.tableColumns).elements(openXml.S.tableColumn).forEach(function(tableColumn, index) {
					let tableColumnId= Number(tableColumn.attribute(openXml.NoNamespace.id).value);
					
					let xmlColumnPr= tableColumn.element(openXml.S.xmlColumnPr);
					let mapId= xmlColumnPr.attribute(openXml.NoNamespace.mapId).value;
					let xpath= xmlColumnPr.attribute(openXml.NoNamespace.xpath).value;
					let xmlDataType= xmlColumnPr.attribute(openXml.NoNamespace.xmlDataType).value;
					
					let lastSeparator= xpath.lastIndexOf('/');
					let tableKey= mapId + xpath.substring('/Root'.length, lastSeparator);
					if (!keyToIdElement[tableKey]) {
						keyToIdElement[tableKey]= {
							id: id,
							element: table,
							tableColumnMap: {}
						};
					}
					
					let columnKey= xpath.substring(lastSeparator + 1);
					keyToIdElement[tableKey].tableColumnMap[columnKey]= {
						tableColumnId: tableColumnId,
						xmlDataType: xmlDataType,
						offset: index
					};
				});
			});
			
			worksheetPart.getPartsByContentType(openXml.contentTypes.singleCellTable).forEach(function(singleCellTablePart) {
				let singleCellTablePartXDoc= singleCellTablePart.getXDocument();
				singleCellTablePartXDoc.root.elements(openXml.S.singleXmlCell).forEach(function(singleXmlCell) {
					let id= Number(singleXmlCell.attribute(openXml.NoNamespace.id).value);
					id2sheetId[id]= sheetId;
					
					let r= singleXmlCell.attribute(openXml.NoNamespace.r).value;
					console.log('[spreadsheet.Layout] singleXmlCell[@id="' + id + '"]/@r="' + r + '"');
					refs[id]= r;
					
					let xmlCellPr= singleXmlCell.element(openXml.S.xmlCellPr);
					let xmlPr= xmlCellPr.element(openXml.S.xmlPr);
					let mapId= xmlPr.attribute(openXml.NoNamespace.mapId).value;
					let xpath= xmlPr.attribute(openXml.NoNamespace.xpath).value;
					let xmlDataType= xmlPr.attribute(openXml.NoNamespace.xmlDataType).value;
					
					let singleXmlCellKey= mapId + xpath.substring('/Root'.length);
					if (!keyToIdElement[singleXmlCellKey]) {
						keyToIdElement[singleXmlCellKey]= {
							id: id,
							element: singleXmlCell,
							xmlDataType: xmlDataType
						};
					}
				});
			});
			
			let row= {};
			sheetData.elements(openXml.S.row).forEach(function(e_row) {
				let row_r= Number(e_row.attribute(openXml.NoNamespace.r).value);
				let c= {};
				e_row.elements(openXml.S.c).forEach(function(e_c) {
					let c_r= e_c.attribute(openXml.NoNamespace.r).value;
					c[c_r]= e_c;
				});
				row[row_r]= {
					element: e_row,
					c: c
				};
			});
			
			sheetLayouts[sheetId]= {
				refs: refs, // @id => @ref|@r
				tables: tables,
				dimension: dimension,
				sheetData: {
					element: sheetData,
					row: row
				}
			};
		});
	};
		
	spreadsheet.Layout.prototype.setValue= function(sheetId, r, value, xmlDataType) {
		if (!isSetValueTarget(value, xmlDataType)) {
			return;
		}
		
		let sheetLayout= this.sheetLayouts[sheetId];
		let rObj= spreadsheet.r.parse(r);
		
		let dimensionRef= spreadsheet.ref.parse(sheetLayout.dimension.attribute(openXml.NoNamespace._ref).value);
		sheetLayout.dimension.setAttributeValue(openXml.NoNamespace._ref, dimensionRef.union(rObj).toString());
		
		let sheetData= sheetLayout.sheetData;
		
		let row= sheetData.row[rObj.row];
		if (row) {
			let spans= spreadsheet.spans.parse(row.element.attribute(openXml.NoNamespace.spans).value);
			spans= spans.union(rObj.column);
			row.element.setAttributeValue(openXml.NoNamespace.spans, spans.toString());
		} else {
			row= {
				element: new XElement(openXml.S.row,
									  new XAttribute(openXml.NoNamespace.r, String(rObj.row)),
									  new XAttribute(openXml.NoNamespace.spans, String(rObj.column) + ':' + String(rObj.column))
									 ),
				c: {}
			};
			addRow(sheetData.element, row.element);
			sheetData.row[rObj.row]= row;
		}
		
		let c= row.c[r];
		if (c) {
			c.removeNodes();
		} else {
			c= new XElement(openXml.S.c,
							new XAttribute(openXml.NoNamespace.r, r)
						   );
			addC(row.element, c);
			row.c[r]= c;
		}
		
		setValue(c, value, xmlDataType);
		
		
		function addRow(sheetData, row) {
			let r= Number(row.attribute(openXml.NoNamespace.r).value);
			let rows_= sheetData.elements(openXml.S.row);
			for (let i= 0, size= rows_.length; i < size; ++i) {
				let row_= rows_[i];
				let r_= Number(row_.attribute(openXml.NoNamespace.r).value);
				
				if (r == r_) {
					// TODO
					// Update @span
					return;
				} else if (r < r_) {
					// TODO
					// Insert r before r_.
					return;
				}
			}
			
			sheetData.add(row);
		}
		
		function addC(row, c) {
			// TODO
			row.add(c);
		}
	};
	
	spreadsheet.Layout.prototype.getTable= function(key) {
		if (this.keyToTable[key]) {
			return this.keyToTable[key];
		}
		
		let tableInfo= this.keyToIdElement[key];
		if (!tableInfo) {
			return null;
		}
		
		let id= tableInfo.id;
		let sheetId= this.id2sheetId[id];
		let sheetLayout= this.sheetLayouts[sheetId];
		if (!sheetLayout) {
			return null;
		}
		
		let ref= sheetLayout.refs[id];
		if (!ref) {
			return null;
		}
		
		let table= new spreadsheet.Table(this, tableInfo);
		this.keyToTable[key]= table;
		return table;
	};
	
	spreadsheet.Layout.prototype.getSingleXmlCell= function(key) {
		if (this.keyToSingleXmlCell[key]) {
			return this.keyToSingleXmlCell[key];
		}
		
		let singleXmlCellInfo= this.keyToIdElement[key];
		if (!singleXmlCellInfo) {
			return null;
		}
		
		let id= singleXmlCellInfo.id;
		let sheetId= this.id2sheetId[id];
		let sheetLayout= this.sheetLayouts[sheetId];
		if (!sheetLayout) {
			return null;
		}
		
		let r= sheetLayout.refs[id];
		if (!r) {
			return null;
		}
		
		let singleXmlCell= new spreadsheet.SingleXmlCell(this, singleXmlCellInfo);
		this.keyToSingleXmlCell[key]= singleXmlCell;
		return singleXmlCell;
	};
	
	
	/**
	 * table
	 */
	// Do not call this constructor directly.
	// Use spreadsheet.Layout.prototype.getTable and check the return value.
	spreadsheet.Table= function(layout, tableInfo) {
		this.keyToTableColumn= {};
		
		this.layout= layout;
		this.tableInfo= tableInfo;
		
		let id= tableInfo.id;
		let sheetId= layout.id2sheetId[id];
		let sheetLayout= layout.sheetLayouts[sheetId];
		
		this.sheetId= sheetId;
		this.ref= spreadsheet.ref.parse(sheetLayout.refs[id]);
		
		let table= tableInfo.element;
		this.tableRef= spreadsheet.ref.parse(table.attribute(openXml.NoNamespace._ref).value);
		this.autoFilter= table.element(openXml.S.autoFilter);
	};
	
	spreadsheet.Table.prototype.setValue= function(tableColumnId, objects, fieldName, xmlDataType) {
		if (!this.ref) {
			return;
		}
		
		let layout= this.layout;
		let sheetId= this.sheetId;
		let ref= this.ref;
		let tableRef= this.tableRef;
		let autoFilter= this.autoFilter;
		
		let table= this.tableInfo.element;
		let tableColumnOffset= this.tableInfo.tableColumnMap[fieldName].offset;
		
		objects.forEach(function(object, index) {
			let value= getValue(object, fieldName);
			let r= new spreadsheet.r(
				ref[0].row + 1 + index,
				ref[0].column + tableColumnOffset
			);
			layout.setValue(sheetId, r.toString(), value, xmlDataType);
			
			tableRef= tableRef.union(r);
			table.setAttributeValue(openXml.NoNamespace._ref, tableRef.toString());
			autoFilter.setAttributeValue(openXml.NoNamespace._ref, tableRef.toString());
		});
		
		
		// Retrieve the field value:
		// e.g. Contact.Name
		// e.g. Contact.Account.Name
		function getValue(object, fieldName) {
			let o= object;
			let f= fieldName;
//			console.log('+ o=' + o.attributes.type + ', f=' + f);
			for (let dot= f.indexOf('.'); dot != -1; dot= f.indexOf('.')) {
				o= o[f.substring(0, dot)];
				f= f.substring(dot + 1);
//				console.log('> o=' + o.attributes.type + ', f=' + f);
			}
			return o[f];
		}
	};
	
	spreadsheet.Table.prototype.getTableColumn= function(fieldName) {
//		console.log('[getTableColumn] fieldName=' + fieldName);
		if (this.keyToTableColumn[fieldName]) {
			return this.keyToTableColumn[fieldName];
		}
		
		let tableColumnInfo= this.tableInfo.tableColumnMap[fieldName];
		if (!tableColumnInfo) {
			return null;
		}
		
		let tableColumn= new spreadsheet.TableColumn(this, tableColumnInfo);
		this.keyToTableColumn[fieldName]= tableColumn;
		return tableColumn;
	};
	
	
	/**
	 * tableColumn
	 */
	// Do not call this constructor directly.
	// Use spreadsheet.Table.prototype.getTableColumn and check the return value.
	spreadsheet.TableColumn= function(table, tableColumnInfo) {
		this.table= table;
		this.tableColumnInfo= tableColumnInfo;
	};
	
	spreadsheet.TableColumn.prototype.setValue= function(objects, fieldName) {
		this.table.setValue(this.tableColumnInfo.tableColumnId, objects, fieldName, this.tableColumnInfo.xmlDataType);
	};
	
	
	/**
	 * singleXmlCell
	 */
	// Do not call this constructor directly.
	// Use spreadsheet.Layout.prototype.getSingleXmlCell and check the return value.
	spreadsheet.SingleXmlCell= function(layout, singleXmlCellInfo) {
		this.layout= layout;
		this.singleXmlCellInfo= singleXmlCellInfo;
		
		let id= singleXmlCellInfo.id;
		let sheetId= layout.id2sheetId[id];
		
		this.sheetId= sheetId;
		this.r= layout.sheetLayouts[sheetId].refs[id];
	};
	
	spreadsheet.SingleXmlCell.prototype.setValue= function(value) {
		this.layout.setValue(this.sheetId, this.r, value, this.singleXmlCellInfo.xmlDataType);
	};
	
	
	// TODO Add all necessary conditions.
	function isSetValueTarget(value, xmlDataType) {
		if (!value) return false;
		if (xmlDataType === 'string' && value.length === 0) {
			return false;
		}
		return true;
	}
	
	function setValue(e, value, xmlDataType) {
		switch (xmlDataType) {
		case 'int':
		case 'long':
			e.setAttributeValue(openXml.NoNamespace.t, 'n');
			e.add(new XElement(openXml.NoNamespace.v, value));
			break;
		default: //'string':
			e.setAttributeValue(openXml.NoNamespace.t, 'inlineStr');
			e.add(
				new XElement(openXml.S._is,
					new XElement(openXml.S.t, value)
				)
			);
			break;
		}
	}
	
	
	// A, B, ...,  Z, AA, AB, ..., AZ, BA, ...,  ZZ, AAA, ...
	// 1, 2, ..., 26, 27, 28, ..., 52, 53, ..., 702, 703, ...
	
	const alphabetToNumber= {
		A:  1, B:  2, C:  3, D:  4, E:  5, F:  6, G:  7, H:  8, I:  9, J: 10, K: 11, L: 12, M: 13,
		N: 14, O: 15, P: 16, Q: 17, R: 18, S: 19, T: 20, U: 21, V: 22, W: 23, X: 24, Y: 25, Z: 26
	};
	const numberToAlphabet= '*ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
	const alphabetsLength= numberToAlphabet.length - 1;
	
	/**
	 * Convert alphabet to number. It's straightforward.
	 */
	spreadsheet.aToN= function(alphabet) {
//		console.log('[spreadsheet.aToN] alphabet="' + alphabet + '"');
		let a= alphabet.split('');
		let base= 1;
		let result= 0;
		
		for (let i= a.length - 1; -1 < i; --i) {
			result+= alphabetToNumber[a[i]]*base;
			base*= alphabetsLength;
		}
		
		return result;
	};
	
/*	// Test
	console.log('  A =   1 : ' + (spreadsheet.aToN(  'A') ===   1));
	console.log('  Z =  26 : ' + (spreadsheet.aToN(  'Z') ===  26));
	console.log(' AA =  27 : ' + (spreadsheet.aToN( 'AA') ===  27));
	console.log(' AZ =  52 : ' + (spreadsheet.aToN( 'AZ') ===  52));
	console.log(' BA =  53 : ' + (spreadsheet.aToN( 'BA') ===  53));
	console.log(' ZZ = 702 : ' + (spreadsheet.aToN( 'ZZ') === 702));
	console.log('AAA = 703 : ' + (spreadsheet.aToN('AAA') === 703));
	*/
	
	/**
	 * Convert number to alphabet. It's a bit complicated.
	 */
	spreadsheet.nToA= function(number) {
//		console.log('[spreadsheet.nToA] number=' + number);
		if (number == 0) {
			return '';
		}
		
		let q= number/alphabetsLength;
		let r= number%alphabetsLength;
		if (r == 0) {
			--q;
			r= alphabetsLength;
		}
		
		return spreadsheet.nToA((number - r)/alphabetsLength) + numberToAlphabet[r];
	};
	
/*	// Test
	console.log('  1 =   A : ' + (spreadsheet.nToA(  1) ===   'A'));
	console.log(' 26 =   Z : ' + (spreadsheet.nToA( 26) ===   'Z'));
	console.log(' 27 =  AA : ' + (spreadsheet.nToA( 27) ===  'AA'));
	console.log(' 52 =  AZ : ' + (spreadsheet.nToA( 52) ===  'AZ'));
	console.log(' 53 =  BA : ' + (spreadsheet.nToA( 53) ===  'BA'));
	console.log('702 =  ZZ : ' + (spreadsheet.nToA(702) ===  'ZZ'));
	console.log('703 = AAA : ' + (spreadsheet.nToA(703) === 'AAA'));
	*/
	
	
	/**
	 * r attribute (e.g. 'A1')
	 */
	spreadsheet.r= function(row, column) {
//		console.log('[spreadsheet.r] row=' + row + ', column=' + column);
		this.row= row;
		this.column= column;
	};
	
	spreadsheet.r.toString= function(row, column) {
		return spreadsheet.nToA(column) + row;
	};
	
	spreadsheet.r.parse= function(r/* e.g. 'A1' */) {
//		console.log('[spreadsheet.r.parse] r="' + r + '"');
		return new spreadsheet.r(
			Number(r.match(/[1-9][0-9]*/)[0]),
			spreadsheet.aToN(r.match(/[A-Z]+/)[0])
		);
	};
	
	spreadsheet.r.prototype.toString= function() {
//		console.log('[spreadsheet.r.prototype.toString] r=(row=' + this.row + ', column=' + this.column + ')');
		return spreadsheet.r.toString(this.row, this.column);
	};
	
	
	/**
	 * ref attribute (e.g. 'A1:B2')
	 */
	spreadsheet.ref= function(r0, r1) {
		this[0]= r0;
		this[1]= r1;
	};
	
	spreadsheet.ref.parse= function(ref/* e.g. 'A1:B2' */) {
//		console.log('[spreadsheet.ref.parse] ref="' + ref + '"');
		let rPair= ref.split(':');
		return new spreadsheet.ref(
			spreadsheet.r.parse(rPair[0]),
			spreadsheet.r.parse(rPair[1])
		);
	};
	
	spreadsheet.ref.prototype.union= function(r) {
//		console.log('[spreadsheet.ref.prototype.union] this=' + this.toString() + ', r=' + r.toString());
		return new spreadsheet.ref(
			new spreadsheet.r(
				r.row < this[0].row ? r.row : this[0].row,
				r.column < this[0].column ? r.column : this[0].column
			),
			new spreadsheet.r(
				this[1].row < r.row ? r.row : this[1].row,
				this[1].column < r.column ? r.column : this[1].column
			)
		);
	};
	
	spreadsheet.ref.prototype.toString= function() {
		return this[0].toString() + ':' + this[1].toString();
	};
	
	
	/**
	 * spans attribute (e.g. '1:2')
	 */
	spreadsheet.spans= function(column0, column1) {
		this[0]= column0;
		this[1]= column1;
	};
	
	spreadsheet.spans.parse= function(spans/* e.g. '1:2' */) {
		let cPair= spans.split(':');
		return new spreadsheet.spans(
			Number(cPair[0]),
			Number(cPair[1])
		);
	};
	
	spreadsheet.spans.prototype.union= function(column) {
		return new spreadsheet.spans(
			column < this[0] ? column : this[0],
			this[1] < column ? column : this[1]
		);
	};
	
	spreadsheet.spans.prototype.toString= function() {
		return String(this[0]) + ':' + this[1];
	};
	
})();
