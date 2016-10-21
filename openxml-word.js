(function(){

    //正規表現パターン（差し込みフィールド用）
    const REGEXP_MERGE = /[A-Za-z\.]+/g;

    //パッケージオブジェクト
    let pkg;

    //メインドキュメント［word/document.xml］XML文書
    let mnXDoc;

    /************************ openXml.Word **************************/

    /**
    * コンストラクタ
    * @param [String] officedoc		Officeファイル（Base64形式）
    */
    openXml.Word = function(officedoc) {

        //パッケージオブジェクト
        pkg = new openXml.OpenXmlPackage(officedoc);

        //メインドキュメント［word/document.xml］XML文書
        mnXDoc = pkg.mainDocumentPart().getXDocument();
    };

    /**
    * 差し込みデータの挿入
    * 注）本処理は、以下条件でのみ動作可能
    * 　１．差し込みフィールド名は《SFDCオブジェクト名_項目名》、子オブジェクトは《SFDCオブジェクト名_子オブジェクト名_項目名》
    * 　２．子オブジェクト項目は、表内にのみ配置可能
    * 　３．差し込みフィールドの装飾（太字、色、etc.）なし
    * 　４．親オブジェクトのデータは差し込み不可
    * @param [Object] mergedata		差し込みデータ
    */
    openXml.Word.prototype.merge = function(mergedata) {

        //bodyタグ
        let body = mnXDoc.root.element(openXml.W.body);

        //表データの差し込み
        body.elements(openXml.W.tbl).forEach(function(tbl, index, ar) {

            //以下は、２行目以降で使用する（データ・子オブジェクト名・項目名）
            let data;
            let objname;
            let colnames = [];

            //1行目の作成
            tbl.descendants(openXml.W.fldSimple).forEach(function(fld, index, ar) {

                //フィールド情報の取得
                let info = getMergeFieldInfo(fld)
                if (info.childname && info.colname) {

                    if (mergedata[info.childname] && mergedata[info.childname].records.length > 0) {
                        let val = mergedata[info.childname].records[0][info.colname];
                        if (val) {
                            //差し込みフィールドをテキストに置き換え
                            fld.replaceWith(new Ltxml.XElement(openXml.W.r, new Ltxml.XElement(openXml.W.t, val)));
                        }
                    }

                    //2行目以降で使用するため、データ・子オブジェクト名・項目名を保存しておく
                    if (!data) data = mergedata[info.childname];
                    if (!objname) objname = info.childname;
                    colnames[index] = info.colname;
                }
            });

            //2行目以降の作成
            if (colnames.length > 0) {
                //１行目を取得
                let tr = tbl.elements(openXml.W.tr).last();

                //レコード数分、１行目のコピーより行追加する
                for (let i=1; i<data.records.length; i++) {
                    let newtr = new Ltxml.XElement(tr);

                    //各列のtタグを編集する
                    newtr.descendants(openXml.W.t).forEach(function(t, index, ar) {
                        let colname = colnames[index];

                        let val = data.records[i][colname];
                        if (val) {
                            //テキスト値を置き換え
                            t.replaceWith(new Ltxml.XElement(openXml.W.t, val));
                        }
                    });
                    tr.parent.add(newtr);
                }
            }
        });

        //表以外のデータの差し込み
        body.descendants(openXml.W.fldSimple).forEach(function(fld, index, ar) {

            //フィールド情報の取得
            let info = getMergeFieldInfo(fld)
            if (info.colname) {

                if (mergedata[info.colname]) {
                    let val = mergedata[info.colname];
                    if (val) {
                        //差し込みフィールドをテキストに置き換え
                        fld.replaceWith(new Ltxml.XElement(openXml.W.r, new Ltxml.XElement(openXml.W.t, val)));
                    }
                }
            }
        });
    };

    /**
    * レポートファイルの出力
    * @param [String] reportName		レポート名
    */
    openXml.Word.prototype.save = function(reportName) {
        pkg.saveToBlobAsync(function (blob) {
            saveAs(blob, reportName+'.docx');
        });
    };

    /************************ inner functions **************************/

    /**
    * 差し込みフィールド名を“_”で分割し、フィールド情報（SFDCオブジェクト名・項目名）を作成する
    * @param fld            fldSimpleエレメント
    * @return フィールド情報
    */
    function getMergeFieldInfo(fld) {

        //差し込みフィールド名の取得（tタグで分割されている場合があるので結合する）
        //例）通常：《Account_Name》、子オブジェクト：《Account_Contacts_Name》
        let fieldName = '';
        fld.descendants(openXml.W.t).forEach(function(t, index, ar) {
            fieldName += t.value;
        });

        //フィールド名を“_”で分割
        let array = fieldName.match(REGEXP_MERGE);

        //フィールド情報の作成
        let fieldinfo = {};
        if (array.length === 2) {
            //通常
            fieldinfo = {
                objname: array[0],
                colname: array[1]
            };
        }
        else if (array.length === 3) {
            //子オブジェクト
            fieldinfo = {
                objname: array[0],
                childname: array[1],
                colname: array[2]
            };
        }
        return fieldinfo;
    }
}());