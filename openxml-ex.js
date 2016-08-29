(function(){
	//スプレッドシート要素名
    var S = openXml.S;
    //名前空間なしの要素名
    var NN = openXml.NoNamespace;

    //属性
    var XAttribute = Ltxml.XAttribute;
    //エレメント
    var XElement = Ltxml.XElement;

    /************************ openXml.Excel **************************/
    openXml.Excel = function() {};

    /**
    * セルエレメントを作成する 【"string","int","long"以外のデータ型は、case文の追加が必要です】
    * @param value          セル値
    * @param type           データ型
    * @param r_attr         r属性値（セル指定文字：例）"A1"）
    * @return XElementオブジェクト
    */
    var newCellElement = openXml.Excel.prototype.newCellElement = function(value, type, r_attr) {
        var cellElement;

        if (value) {
            switch (type) {
                case "string":
                    cellElement = new XElement(S.c, new XAttribute(NN.r, r_attr), new XAttribute(NN.t, "inlineStr"),
                                               new XElement(S._is,  new XElement(S.t, value)));
                    break;
                case "int":
                case "long":
                    cellElement = new XElement(S.c, new XAttribute(NN.r, r_attr), new XAttribute(NN.t, "n"),
                                               new XElement(NN.v, value));
                    break;
                default:
                    cellElement = new XElement(S.c, new XAttribute(NN.r, r_attr), new XAttribute(NN.s, "1"));
                    break;
            }
        }
        else {
            cellElement = new XElement(S.c, new XAttribute(NN.r, r_attr), new XAttribute(NN.s, "1"));
        }

        return cellElement;
    }

    /**
    * 行エレメントを作成する
    * @param rowdata        行データ
    * @param rownum         行番号
    * @param tableCNames    テーブル列名
    * @param sheetCNames    シート列名
    * @param dataTypes      データ型
    * @return XElementオブジェクト
    */
    var newRowElement = openXml.Excel.prototype.newRowElement = function(rowdata, rownum, tableCNames, sheetCNames, dataTypes) {
        var rowElement =  new XElement(S.row, new XAttribute(NN.r, rownum));

        for (var i=0; i<tableCNames.length; i++) {
            var name = tableCNames[i];
            var r_attr = sheetCNames[i] + rownum;

            var cellElement =  newCellElement(rowdata[name], dataTypes[name], r_attr);
            rowElement.add(cellElement);
        }

        return rowElement;
    }

    /**
    * テーブル範囲の設定文字列より、先頭行・最終行・左位置・右位置を取得する
    * 例） "C3:F4"-> 先頭行：3、最終行：4、左位置："C"、右位置："F"
    * @param tablePart      テーブルパートオブジェクト
    * @return テーブル範囲オブジェクト
    */
    var getTableRange = openXml.Excel.prototype.getTableRange = function(tablePart) {
        var range = {};

        var tbXDoc = tablePart.getXDocument();
        var refs = tbXDoc.root.attribute(NN._ref).value;

        range.top = Number(refs.match( /\d+/g)[0]);
        range.bottom = Number(refs.match( /\d+/g)[1]);
        range.left = refs.match(/[A-Z]+/g)[0];
        range.right = refs.match(/[A-Z]+/g)[1];

        return range;
    }

    /**
    * データ件数よりテーブル範囲の設定文字列を変更する
    * @param tablePart      テーブルパートオブジェクト
    * @param datacount      データ件数
    * @return 変更後のテーブル範囲オブジェクト
    */
    var setTableRange = openXml.Excel.prototype.setTableRange = function(tablePart, datacount) {
        var range = getTableRange(tablePart);
        range.bottom = range.top + datacount;

        var refs = range.left + range.top + ":" + range.right + range.bottom;

        //table1.xmlの要素(root,autoFilter)の属性(ref)を変更
        var tbXDoc = tablePart.getXDocument();
        tbXDoc.root.setAttributeValue(NN._ref, refs);
        tbXDoc.root.element(S.autoFilter).setAttributeValue(NN._ref, refs);

        return range;
    }

    /**
    * テーブル範囲のシートデータをクリアする
    * @param worksheetPart  ワークシートパートオブジェクト
    * @param tablePart      テーブルパートオブジェクト
    * @return なし
    */
    var clearSheetData = openXml.Excel.prototype.clearSheetData = function(worksheetPart, tablePart) {
        var wsXDoc = worksheetPart.getXDocument();
        var sheetDataElement = wsXDoc.root.element(S.sheetData);
        var rowElements = sheetDataElement.elements(S.row);

        var range = getTableRange(tablePart);
        var from = range.top + 1;
        var to = range.bottom;

        for (var i=0; i<rowElements.count(); i++) {
            var el = rowElements.elementAt(i);
            var r = Number(el.attribute(NN.r).value);
            if (from <= r && r <=to) {
                el.remove();
            }
        }
    }

    /**
    * テーブルの全列名を取得する
    * @param tablePart      テーブルパートオブジェクト
    * @return 列名の配列
    */
    var getTableColumnNames = openXml.Excel.prototype.getTableColumnNames = function(tablePart) {
        var tableCNames = [];

        var tbXDoc = tablePart.getXDocument();
        var elements = tbXDoc.root.element(S.tableColumns).elements(S.tableColumn);

        for (var i=0; i<elements.count(); i++) {
            var el = elements.elementAt(i);
            tableCNames.push(el.attribute(NN.uniqueName).value);
        }

        return tableCNames;
    }

    /**
    * テーブル列に対応するシートの全列名を取得する
    * @param worksheetPart  ワークシートパートオブジェクト
    * @param tablePart      テーブルパートオブジェクト
    * @return 列名の配列
    */
    var getSheetColumnNames = openXml.Excel.prototype.getSheetColumnNames = function(worksheetPart, tablePart) {
        var sheetCNames = [];

        var wsXDoc = worksheetPart.getXDocument();
        var sheetDataElement = wsXDoc.root.element(S.sheetData);
        var range = getTableRange(tablePart);

        var rowElement = sheetDataElement.elements(S.row).first(function(row) {
            return row.attribute(NN.r).value === String(range.top);
        });

        var elements = rowElement.elements(S.c);
        for (var i=0; i<elements.count(); i++) {
            var el = elements.elementAt(i);
            var val = el.attribute(NN.r).value;
            sheetCNames.push(val.match(/[A-Z]/)[0]);
        }

        return sheetCNames;
    }

    /**
    * スキーマ情報より、テーブル列のデータ型を取得する
    * @param xmlmapPart     マップパートオブジェクト
    * @return データ型オブジェクト（テーブル列名毎）
    */
    var getDataTypes = openXml.Excel.prototype.getDataTypes = function(xmlmapPart) {
        var dataTypes = {};

        var mapXDoc = xmlmapPart.getXDocument();
        var schemaElement = mapXDoc.root.element(new Ltxml.XName("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "Schema"));
        var elements = getMapElements(schemaElement);

        for (var i=0; i<elements.count(); i++) {
            var el = elements.elementAt(i);
            var type = el.attribute(NN.type).value;
            var name = el.attribute(NN.name).value;
            dataTypes[name] = type.match(/[a-z]*$/)[0];
        }

        return dataTypes;
    }

    /**
    * スキーマ情報より、type属性を持った全エレメントを取得する（内部関数）
    * @param current        マップパートオブジェクト
    * @return type属性を持った全エレメント
    */
    function getMapElements(current) {
        var elements = current.elements();
        if (elements.count() > 0) {
            var el = elements.elementAt(0);
            if (el.attribute(NN.type)) {
                return elements;
            }
            else {
                return getMapElements(el);
            }
        }
        else {
            return null;
        }
    }

    /**
    * 差し込みデータより、テーブルに行エレメントを追加する
    * @param mergedata      差し込みデータ
    * @param worksheetPart  ワークシートパートオブジェクト
    * @param tablePart      テーブルパートオブジェクト
    * @param xmlmapPart     マップパートオブジェクト
    * @return なし
    */
    var mergeSheetData = openXml.Excel.prototype.mergeSheetData = function(mergedata, worksheetPart, tablePart, xmlmapPart) {
        var wsXDoc = worksheetPart.getXDocument();
        var sheetDataElement = wsXDoc.root.element(S.sheetData);

        var range = getTableRange(tablePart);
        var from = range.top + 1;
        var to = range.bottom;

        var tableCNames = getTableColumnNames(tablePart);
        var sheetCNames = getSheetColumnNames(worksheetPart, tablePart);
        var dataTypes = getDataTypes(xmlmapPart);

        for (var i=0; i<=(to-from); i++) {
			var rownum = i + from;
            var rowElement = newRowElement(mergedata[i], rownum, tableCNames, sheetCNames, dataTypes);
            sheetDataElement.add(rowElement);
        }
    }

    /*********** OpenXmlPart ***********/

    openXml.OpenXmlPart.prototype.tableSingleCellParts = function () {
        return this.getPartsByRelationshipType(openXml.relationshipTypes.singleCellTable);
    };
}());
