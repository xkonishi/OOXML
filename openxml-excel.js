(function(){

    //パッケージオブジェクト
    let pkg;

    //ワークシート［xl/worksheets/sheet1.xml］XML文書
    let wsXDoc;
    //テーブル［xl/tables/table1.xml］XML文書
    let tbXDoc;
    //マップ［xl/xmlMaps.xml］XML文書
    let mapXDoc;

    /************************ openXml.Excel **************************/

    /**
    * コンストラクタ
    * @param [String] officedoc		Officeファイル（Base64形式）
    */
    openXml.Excel = function(officedoc) {

        //パッケージオブジェクト
        pkg = new openXml.OpenXmlPackage(officedoc);

        //ワークシート［xl/worksheets/sheet1.xml］XML文書
        let worksheetPart = pkg.workbookPart().worksheetParts()[0];
        wsXDoc = worksheetPart.getXDocument();

        //テーブル［xl/tables/table1.xml］XML文書
        let tablePart = worksheetPart.tableDefinitionParts()[0];
        tbXDoc = tablePart.getXDocument();

        //マップ［xl/xmlMaps.xml］XML文書
        let xmlmapPart = pkg.getPartByUri('/xl/xmlMaps.xml');
        mapXDoc = xmlmapPart.getXDocument();
    };

    /**
    * 差し込みデータの挿入
    * @param [Array] mergedata		差し込みデータ
    */
    openXml.Excel.prototype.merge = function(mergedata) {

        let ref = tbXDoc.root.attribute(openXml.NoNamespace._ref).value;
        let range = toTableRange(ref);

        let elTCs = tbXDoc.root.element(openXml.S.tableColumns);
        let columns = openXml.Util.findElements(elTCs, openXml.S.tableColumn);

        let colInfo = [];
        columns.forEach(function(col, index, ar) {
            let obj = {};
            obj.name = col.attribute(openXml.NoNamespace.uniqueName).value;
            obj.type = col.element(openXml.S.xmlColumnPr).attribute(openXml.NoNamespace.xmlDataType).value;
            colInfo[index] = obj;
        });

        let elSD = wsXDoc.root.element(openXml.S.sheetData);
        let rows = openXml.Util.findElements(elSD, openXml.S.row);

        if (rows.length >= 2) {
            //ヘッダ行の取得＆先頭データ行の削除
            //（テーブルはヘッダ行＋1行のデータ行のみで、テーブル以下にデータが存在しない前提）
            let head = rows[rows.length-2];
            rows[rows.length-1].remove();

            //行番号
            let rownum = range.bottom;

            //ヘッダ行をコピーし、新規行を作成
            let newrow = new Ltxml.XElement(head);
            newrow.setAttributeValue(openXml.NoNamespace.r, rownum);

            //先頭データの取得
            let data = mergedata[0];

            //先頭データの設定
            let cs = openXml.Util.findElements(newrow, openXml.S.c);
            cs.forEach(function(c, index, ar) {
                let info = colInfo[index];
                let value = data[info.name];
                let type = info.type;
                let r_attr = c.attribute(openXml.NoNamespace.r).value.match(/[A-Z]+/) + rownum;
                c.parent.add(newCellElement(value, type, r_attr));
                c.remove();
            });

            //新規行を追加
            head.parent.add(newrow);

            //２行目以降を追加
            for (let i=1; i<mergedata.length; i++) {

                //行番号
                let rownum = range.bottom + i;

                //ヘッダ行をコピーし、新規行を作成
                let newrow = new Ltxml.XElement(head);
                newrow.setAttributeValue(openXml.NoNamespace.r, rownum);

                //データの取得
                let data = mergedata[i];

                //データの設定
                let cs = openXml.Util.findElements(newrow, openXml.S.c);
                cs.forEach(function(c, index, ar) {
                    let info = colInfo[index];
                    let value = data[info.name];
                    let type = info.type;
                    let r_attr = c.attribute(openXml.NoNamespace.r).value.match(/[A-Z]+/) + rownum;
                    c.parent.add(newCellElement(value, type, r_attr));
                    c.remove();
                });

                //新規行を追加
                head.parent.add(newrow);
            }
        }







        //■■■■■■■■■■■　以下は修正予定　■■■■■■■■■■■

/*
        //シートデータのクリア
        clearSheetData(worksheetPart, tablePart);

        //テーブルの表示範囲を設定
        setTableRange(tablePart, mergedata.length);

        //差し込みデータの挿入
        mergeSheetData(mergedata, worksheetPart, tablePart, xmlmapPart);
        */
    };

    /**
    * レポートファイルの出力
    * @param [String] reportName		レポート名
    */
    openXml.Excel.prototype.save = function(reportName) {
        pkg.saveToBlobAsync(function (blob) {
            saveAs(blob, reportName+'.xlsx');
        });
    };

    /************************ inner functions **************************/

    /**
    * テーブル範囲の設定文字列より、先頭行・最終行・左位置・右位置を取得する
    * 例） "C3:F4"-> 先頭行：3、最終行：4、左位置："C"、右位置："F"
    * @param ref            ref属性の値
    * @return テーブル範囲オブジェクト
    */
    function toTableRange(ref) {
        let range = {};
        range.top = Number(ref.match( /\d+/g)[0]);
        range.bottom = Number(ref.match( /\d+/g)[1]);
        range.left = ref.match(/[A-Z]+/g)[0];
        range.right = ref.match(/[A-Z]+/g)[1];

        return range;
    }

    /**
    * セルエレメントを作成する 【"string","int","long"以外のデータ型は、case文の追加が必要です】
    * @param value          セル値
    * @param type           データ型
    * @param r_attr         r属性値（セル指定文字：例）"A1"）
    * @return XElementオブジェクト
    */
    function newCellElement(value, type, r_attr) {
        var cellElement;

        if (value) {
            switch (type) {
                case "string":
                    cellElement = new XElement(openXml.S.c, new XAttribute(openXml.NoNamespace.r, r_attr), new XAttribute(openXml.NoNamespace.t, "inlineStr"),
                                        new XElement(openXml.S._is,  new XElement(openXml.S.t, value)));
                    break;
                case "int":
                case "long":
                    cellElement = new XElement(openXml.S.c, new XAttribute(openXml.NoNamespace.r, r_attr), new XAttribute(openXml.NoNamespace.t, "n"),
                                        new XElement(openXml.NoNamespace.v, value));
                    break;
                default:
                    cellElement = new XElement(openXml.S.c, new XAttribute(openXml.NoNamespace.r, r_attr), new XAttribute(openXml.NoNamespace.s, "1"));
                    break;
            }
        }
        else {
            cellElement = new XElement(openXml.S.c, new XAttribute(openXml.NoNamespace.r, r_attr), new XAttribute(openXml.NoNamespace.s, "1"));
        }

        return cellElement;
    }
    //■■■■■■■■■■■　以下は修正予定　■■■■■■■■■■■

    /**
    * セルエレメントを作成する 【"string","int","long"以外のデータ型は、case文の追加が必要です】
    * @param value          セル値
    * @param type           データ型
    * @param r_attr         r属性値（セル指定文字：例）"A1"）
    * @return Ltxml.XElementオブジェクト
    */
    function newCellElement(value, type, r_attr) {
        var cellElement;

        if (value) {
            switch (type) {
                case "string":
                    cellElement = new Ltxml.XElement(openXml.S.c, new Ltxml.XAttribute(openXml.NoNamespace.r, r_attr), new Ltxml.XAttribute(openXml.NoNamespace.t, "inlineStr"),
                                               new Ltxml.XElement(openXml.S._is,  new Ltxml.XElement(openXml.S.t, value)));
                    break;
                case "int":
                case "long":
                    cellElement = new Ltxml.XElement(openXml.S.c, new Ltxml.XAttribute(openXml.NoNamespace.r, r_attr), new Ltxml.XAttribute(openXml.NoNamespace.t, "n"),
                                               new Ltxml.XElement(openXml.NoNamespace.v, value));
                    break;
                default:
                    cellElement = new Ltxml.XElement(openXml.S.c, new Ltxml.XAttribute(openXml.NoNamespace.r, r_attr), new Ltxml.XAttribute(openXml.NoNamespace.s, "1"));
                    break;
            }
        }
        else {
            cellElement = new Ltxml.XElement(openXml.S.c, new Ltxml.XAttribute(openXml.NoNamespace.r, r_attr), new Ltxml.XAttribute(openXml.NoNamespace.s, "1"));
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
    * @return Ltxml.XElementオブジェクト
    */
    function newRowElement(rowdata, rownum, tableCNames, sheetCNames, dataTypes) {
        var rowElement =  new Ltxml.XElement(openXml.S.row, new Ltxml.XAttribute(openXml.NoNamespace.r, rownum));

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
    /*
    function getTableRange(tablePart) {
        var range = {};

        var tbXDoc = tablePart.getXDocument();
        var refs = tbXDoc.root.attribute(openXml.NoNamespace._ref).value;

        range.top = Number(refs.match( /\d+/g)[0]);
        range.bottom = Number(refs.match( /\d+/g)[1]);
        range.left = refs.match(/[A-Z]+/g)[0];
        range.right = refs.match(/[A-Z]+/g)[1];

        return range;
    }*/

    /**
    * データ件数よりテーブル範囲の設定文字列を変更する
    * @param tablePart      テーブルパートオブジェクト
    * @param datacount      データ件数
    * @return 変更後のテーブル範囲オブジェクト
    */
    function setTableRange(tablePart, datacount) {
        var range = getTableRange(tablePart);
        range.bottom = range.top + datacount;

        var refs = range.left + range.top + ":" + range.right + range.bottom;

        //table1.xmlの要素(root,autoFilter)の属性(ref)を変更
        var tbXDoc = tablePart.getXDocument();
        tbXDoc.root.setAttributeValue(openXml.NoNamespace._ref, refs);
        tbXDoc.root.element(openXml.S.autoFilter).setAttributeValue(openXml.NoNamespace._ref, refs);

        return range;
    }

    /**
    * テーブル範囲のシートデータをクリアする
    * @param worksheetPart  ワークシートパートオブジェクト
    * @param tablePart      テーブルパートオブジェクト
    * @return なし
    */
    function clearSheetData(worksheetPart, tablePart) {
        var wsXDoc = worksheetPart.getXDocument();
        var sheetDataElement = wsXDoc.root.element(openXml.S.sheetData);
        var rowElements = sheetDataElement.elements(openXml.S.row);

        var range = getTableRange(tablePart);
        var from = range.top + 1;
        var to = range.bottom;

        for (var i=0; i<rowElements.count(); i++) {
            var el = rowElements.elementAt(i);
            var r = Number(el.attribute(openXml.NoNamespace.r).value);
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
    function getTableColumnNames(tablePart) {
        var tableCNames = [];

        var tbXDoc = tablePart.getXDocument();
        var elements = tbXDoc.root.element(openXml.S.tableColumns).elements(openXml.S.tableColumn);

        for (var i=0; i<elements.count(); i++) {
            var el = elements.elementAt(i);
            tableCNames.push(el.attribute(openXml.NoNamespace.uniqueName).value);
        }

        return tableCNames;
    }

    /**
    * テーブル列に対応するシートの全列名を取得する
    * @param worksheetPart  ワークシートパートオブジェクト
    * @param tablePart      テーブルパートオブジェクト
    * @return 列名の配列
    */
    function getSheetColumnNames(worksheetPart, tablePart) {
        var sheetCNames = [];

        var wsXDoc = worksheetPart.getXDocument();
        var sheetDataElement = wsXDoc.root.element(openXml.S.sheetData);
        var range = getTableRange(tablePart);

        var rowElement = sheetDataElement.elements(openXml.S.row).first(function(row) {
            return row.attribute(openXml.NoNamespace.r).value === String(range.top);
        });

        var elements = rowElement.elements(openXml.S.c);
        for (var i=0; i<elements.count(); i++) {
            var el = elements.elementAt(i);
            var val = el.attribute(openXml.NoNamespace.r).value;
            sheetCNames.push(val.match(/[A-Z]/)[0]);
        }

        return sheetCNames;
    }

    /**
    * スキーマ情報より、テーブル列のデータ型を取得する
    * @param xmlmapPart     マップパートオブジェクト
    * @return データ型オブジェクト（テーブル列名毎）
    */
    function getDataTypes(xmlmapPart) {
        var dataTypes = {};

        var mapXDoc = xmlmapPart.getXDocument();
        var schemaElement = mapXDoc.root.element(new Ltxml.XName("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "Schema"));
        var elements = getMapElements(schemaElement);

        for (var i=0; i<elements.count(); i++) {
            var el = elements.elementAt(i);
            var type = el.attribute(openXml.NoNamespace.type).value;
            var name = el.attribute(openXml.NoNamespace.name).value;
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
            if (el.attribute(openXml.NoNamespace.type)) {
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
    function mergeSheetData(mergedata, worksheetPart, tablePart, xmlmapPart) {
        var wsXDoc = worksheetPart.getXDocument();
        var sheetDataElement = wsXDoc.root.element(openXml.S.sheetData);

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

}());