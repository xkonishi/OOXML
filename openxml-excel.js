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

        //テーブル情報の取得
        let colInfo = [];
        tbXDoc.root.descendants(openXml.S.tableColumn).forEach(function(col, index, ar) {
            colInfo[index] = {
                name: col.attribute(openXml.NoNamespace.uniqueName).value,
                type: col.element(openXml.S.xmlColumnPr).attribute(openXml.NoNamespace.xmlDataType).value
            };
        });

        // let elSD = wsXDoc.root.element(openXml.S.sheetData);
        // let rows = openXml.Util.findElements(elSD, openXml.S.row);

        //シートの行データを取得
        let rows = wsXDoc.root.element(openXml.S.sheetData).descendants(openXml.S.row);

        //本処理は、以下の条件以外は動作保証しない
        //・シート内のテーブルは１つ
        //・テーブルはヘッダ行＋1行の空行
        //・テーブル以降のセルに余計な値を設定しない（ヘッダ行をコピーしてデータ行を作成しているため）
        if (rows.count() >= 2) {
            //テーブルの空行を削除（最終行も１減らす）
            rows.last().remove();
            range.bottom -= 1;

            //テーブルのヘッダ行を取得
            let head = rows.elementAt(rows.count()-2);

            //データ件数分データ行を追加する
            for (let i=0; i<mergedata.length; i++) {

                //テーブルの最終行番号の更新
                range.bottom += 1;

                //ヘッダ行をコピーし、新規行を作成
                let newrow = new Ltxml.XElement(head);
                newrow.setAttributeValue(openXml.NoNamespace.r, range.bottom);

                //データの設定
                newrow.descendants(openXml.S.c).forEach(function(c, index, ar) {
                    let info = colInfo[index];
                    let value = mergedata[i][info.name];
                    let type = info.type;
                    let r_attr = c.attribute(openXml.NoNamespace.r).value.match(/[A-Z]+/) + range.bottom;
                    c.parent.add(newCellElement(value, type, r_attr));
                    c.remove();
                });

                //新規行を追加
                head.parent.add(newrow);
            }
        }

        //テーブル範囲の更新
        ref = rangeToRef(range);
        tbXDoc.root.setAttributeValue(openXml.NoNamespace._ref, ref);

        //テーブル範囲の更新（フィルタ）
        tbXDoc.root.element(openXml.S.autoFilter).setAttributeValue(openXml.NoNamespace._ref, ref);
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
        //セル属性の初期値（空白セル用）
        let attr = [new Ltxml.XAttribute(openXml.NoNamespace.r, r_attr), new Ltxml.XAttribute(openXml.NoNamespace.s, "1")];
        //データ用エレメント
        let dataEl;

        if (value) {
            switch (type) {
                case "string":
                    attr[1] = new Ltxml.XAttribute(openXml.NoNamespace.t, "inlineStr");
                    dataEl = new Ltxml.XElement(openXml.S._is,  new Ltxml.XElement(openXml.S.t, value));
                    break;
                case "int":
                case "long":
                    attr[1] = new Ltxml.XAttribute(openXml.NoNamespace.t, "n");
                    dataEl = new Ltxml.XElement(openXml.NoNamespace.v, value);
                    break;
                default:
                    break;
            }
        }
        return new Ltxml.XElement(openXml.S.c, attr[0], attr[1], dataEl);
    }
}());