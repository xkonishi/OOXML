if (!ReportFactory) {
	function ReportFactory() {}
}

(function() {
	
	const XElement= Ltxml.XElement;
	const XAttribute= Ltxml.XAttribute;
	
	
	ReportFactory.Excel= function(template, queryArray) {
		this.pkg= new openXml.OpenXmlPackage(template);
		this.queryArray= queryArray;
		
		this.layout= new spreadsheet.Layout(this.pkg);
	};
	
	// BEGIN adapter interfaces
	ReportFactory.Excel.prototype.merge= function(data) {
		this.create(data);
	}
	
	ReportFactory.Excel.prototype.save= function(reportName) {
		this.pkg.saveToBlobAsync(function(blob) {
			saveAs(blob, reportName + '.xlsx');
		});
	};
	// END adapter interfaces
	
	ReportFactory.Excel.prototype.create= function(data) {
//		alert('queryArray.length=' + queryArray.length + ', data.length=' + data.length);
		
		let layout= this.layout;
		
		this.queryArray.forEach(function(q, index) {
			console.log('[create] query=' + JSON.stringify(q));
			
			let mapId= String(index + 1);
			let objects= data[index];
			
			if (q.c) {
				// with child relationships >> single parent record
				let parent= objects[0];
				if (parent) {
					fillTableSingleCells(mapId, q, layout, parent);
					
					q.c.forEach(function(q) {
						if (parent[q.f]) {
							let children= parent[q.f].records;
							fillTable(mapId, q, layout, children);
						}
					});
				}
			} else {
				// without child relationships >> multiple records
				fillTable(mapId, q, layout, objects);
			}
		});
		
		return this.pkg.saveToBase64();
		
		
		function fillTable(mapId, q, layout, objects) {
			let tableKey= mapId + '/' + q.f;
			let table= layout.getTable(tableKey);
//			console.log('[ReportFactoryExcel#fillTable] tableKey=' + tableKey + ', table=' + table);
			if (table) {
				q.s.forEach(function(s) {
					let tableColumn= table.getTableColumn(s.f);
//					console.log('[ReportFactoryExcel#fillTable] s.f=' + s.f + ', tableColumn=' + tableColumn);
					if (tableColumn) {
						tableColumn.setValue(objects, s.f);
					}
				});
			}
		}
		
		function fillTableSingleCells(mapId, q, layout, object) {
			q.s.forEach(function(s) {
				let singleXmlCellKey= mapId + '/' + q.f + '/' + s.f;
				let singleXmlCell= layout.getSingleXmlCell(singleXmlCellKey);
				if (singleXmlCell) {
					singleXmlCell.setValue(object[s.f]);
				}
			});
		}
	};
	
})();