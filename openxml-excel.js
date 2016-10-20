(function(){

    //パッケージオブジェクト
    let pkg;

    //ワークシート［xl/worksheets/sheet1.xml］XML文書
    let wsXDoc;
    //テーブル［xl/tables/table1.xml］XML文書
    let tbXDoc;

    /************************ openXml.Excel **************************/

    /**
    * コンストラクタ
    * @param [String] officedoc		Officeファイル（Base64形式）
    */
    openXml.Excel = function(officedoc) {

        //パッケージオブジェクト
        pkg = new openXml.OpenXmlPackage(officedoc);

        //ワークシート［xl/worksheets/sheet1.xml］XML文書
        let worksheetPart = pkg.workbookPart().worksheetParts()[0];//検証プログラムのため、先頭シート固定とする
        wsXDoc = worksheetPart.getXDocument();

        //テーブル［xl/tables/table1.xml］XML文書
        let tablePart = worksheetPart.tableDefinitionParts()[0];//検証プログラムのため、先頭テーブル固定とする
        tbXDoc = tablePart.getXDocument();
    };

    /**
    * 差し込みデータの挿入
    * @param [Array] mergedata		差し込みデータ
    */
    openXml.Excel.prototype.merge = function(mergedata) {

        //テーブル範囲の取得
        let ref = tbXDoc.root.attribute(openXml.NoNamespace._ref).value;
        let range = refToRange(ref);

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

            //２行目以降を追加
            for (let i=0; i<mergedata.length; i++) {

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
    function refToRange(ref) {
        let ar = ref.split(':');
        let range = {
            top:    Number(ar[0].match(/[0-9]/)),
            bottom: Number(ar[1].match(/[0-9]/)),
            left:   ar[0].match(/[A-Z]/),
            right:  ar[1].match(/[A-Z]/)
        };
        return range;
    }

    /**
    * refToRangeで取得したテーブル範囲を文字列に戻す
    * @param range          テーブル範囲オブジェクト
    * @return ref属性の値
    */
    function rangeToRef(range) {
        let ref = range.left
                + range.top
                + ':'
                + range.right
                + range.bottom;
        return ref;
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
}());