(function(){

    /************************ openXml.Word **************************/
    const REGEXP_SOBJ = /[A-Za-z\.]+/g;

    let pkg;
    let mnPart;
    let mnXDoc;

    openXml.Word = function(officedoc) {
        pkg = new openXml.OpenXmlPackage(officedoc);
        mnPart = pkg.mainDocumentPart();
        mnXDoc = mnPart.getXDocument();
    };

    openXml.Word.prototype.merge = function(mergedata) {
        let data = mergedata[0];
        let body = mnXDoc.root.element(openXml.W.body);

        let flds = openXml.Util.findElements(body, openXml.W.fldSimple, openXml.W.tbl);
        flds.forEach(function(fld, index, ar) {
            let fieldName = fld.element(openXml.W.r).element(openXml.W.t).value;
            let sobjInfo = fieldName.match(REGEXP_SOBJ);

            if (sobjInfo.length === 2) {
                let colname = sobjInfo[1];

                if (data[colname]) {
                    let val = data[colname];
                    if (val) {
                        //テキスト挿入、および差し込みフィールドの削除
                        fld.parent.add(new Ltxml.XElement(openXml.W.r, new Ltxml.XElement(openXml.W.t, val)));
                        fld.remove();
                    }
                }
            }
        });

        let tbls = body.elements(openXml.W.tbl);
        tbls.forEach(function(tbl, index, ar) {
            let exeptElements = [openXml.W.tblPr, openXml.W.tblGrid];
            let flds = openXml.Util.findElements(tbl, openXml.W.fldSimple, exeptElements);

            let objname;
            let colnames = [];

            //１行目のデータ設定
            flds.forEach(function(fld, index, ar) {
                let fieldName = fld.element(openXml.W.r).element(openXml.W.t).value;
                let sobjInfo = fieldName.match(REGEXP_SOBJ);

                if (sobjInfo.length === 3) {
                    if (!objname) objname = sobjInfo[1];
                    let colname = sobjInfo[2];
                    colnames[index] = colname;

                    if (data[objname] && data[objname].records.length > 0) {
                        let val = data[objname].records[0][colname];
                        if (val) {
                            //テキスト挿入、および差し込みフィールドの削除
                            fld.parent.add(new Ltxml.XElement(openXml.W.r, new Ltxml.XElement(openXml.W.t, val)));
                            fld.remove();
                        }
                    }
                }
            });

            //２行目以降を追加
            if (flds.length > 0) {
                let tr = flds[0].parent.parent.parent;//tbl
                for (let i=1; i<data[objname].records.length; i++) {
                    let trnew = new Ltxml.XElement(tr);

                    let ts = openXml.Util.findElements(trnew, openXml.W.t, openXml.W.tcPr);
                    ts.forEach(function(t, index, ar) {
                        let colname = colnames[index];

                        let val = data[objname].records[i][colname];
                        if (val) {
                            //テキスト置き換え
                            t.parent.add(new Ltxml.XElement(openXml.W.t, val));
                            t.remove();
                        }
                    });
                    tr.parent.add(trnew);
                }
            }
        });
    };

    openXml.Word.prototype.save = function(reportName) {
        pkg.saveToBlobAsync(function (blob) {
            saveAs(blob, reportName+'.docx');
        });
    };

    /************************ inner functions **************************/

}());