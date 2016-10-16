﻿(function(){

    /************************ openXml.Word **************************/
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
            let sobjInfo = fieldName.match(/[A-Za-z\.]+/g);

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

            //１行目のデータ設定
            flds.forEach(function(fld, index, ar) {
                let fieldName = fld.element(openXml.W.r).element(openXml.W.t).value;
                let sobjInfo = fieldName.match(/[A-Za-z\.]+/g);

                if (sobjInfo.length === 3) {
                    let objname = sobjInfo[1];
                    let colname = sobjInfo[2];

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
        });




        /*
        let mergeFields = _getMergeFields();

        let data = mergedata[0];
        Object.keys(data).forEach(function(key, index, ar) {
            if ((typeof data[key]) !== 'object') {
                let val = data[key];
                if (mergeFields[key]) {
                    let el = mergeFields[key].el;
                    if (el) {
                        el.element(openXml.W.fldSimple).remove();
                        el.add(new Ltxml.XElement(openXml.W.r, new Ltxml.XElement(openXml.W.t, val)));
                    }
                }
            }
            else {
                let sfdc_childobj = key;
                if (mergeFields[sfdc_childobj]) {
                    let obj = data[sfdc_childobj];

                    //2行目以降を追加
                    let tr = mergeFields[sfdc_childobj].el;
                    let trs = [tr];
                    for (let i=1; i<obj.records.length; i++) {
                        let trnew = new Ltxml.XElement(tr);
                        tr.parent.add(trnew);
                        trs.push(trnew);
                    }

                    for (let i=0; i<obj.records.length; i++) {
                        let data = obj.records[i];
                        let tr = trs[i];

                        Object.keys(data).forEach(function(key, index, ar) {
                            if ((typeof data[key]) !== 'object') {
                                let val = data[key];
                                if (mergeFields[sfdc_childobj][key]) {
                                    if (i === 0) {
                                        let el = mergeFields[sfdc_childobj][key].el;
                                        if (el) {
                                            //差し込みフィールドの削除、およびテキスト挿入
                                            el.element(openXml.W.fldSimple).remove();
                                            el.add(new Ltxml.XElement(openXml.W.r, new Ltxml.XElement(openXml.W.t, val)));
                                        }
                                    }
                                    else {

                                    }
                                }
                            }
                        });
                    }
                }
            }
        });*/
    };

    openXml.Word.prototype.save = function(reportName) {
        pkg.saveToBlobAsync(function (blob) {
            saveAs(blob, reportName+'.docx');
        });
    };

    /************************ inner functions **************************/

    function _getMergeFields() {
        let mergeFields = {};

        let body = mnXDoc.root.element(openXml.W.body);

        let flds = openXml.Util.findElements(body, openXml.W.fldSimple, openXml.W.tbl);
        flds.forEach(function(fld, index, ar) {
            let obj = {}
            obj.el = fld.parent;
            obj.fieldName = fld.element(openXml.W.r).element(openXml.W.t).value;
            obj.sfdcInfo = obj.fieldName.match(/[A-Za-z\.]+/g);

            let sfdc_colname = obj.sfdcInfo[1];
            mergeFields[sfdc_colname] = obj;
        });

        let tbls = body.elements(openXml.W.tbl);
        tbls.forEach(function(tbl, index, ar) {
            let exeptElements = [openXml.W.tblPr, openXml.W.tblGrid];
            let flds = openXml.Util.findElements(tbl, openXml.W.fldSimple, exeptElements);

            flds.forEach(function(fld, index, ar) {
                let obj = {}
                obj.el = fld.parent;
                obj.fieldName = fld.element(openXml.W.r).element(openXml.W.t).value;
                obj.sfdcInfo = obj.fieldName.match(/[A-Za-z]+/g);

                sfdc_childobj = obj.sfdcInfo[1];
                if (!mergeFields[sfdc_childobj]) {
                    mergeFields[sfdc_childobj] = {};
                }

                let sfdc_colname = obj.sfdcInfo[2];
                mergeFields[sfdc_childobj].el = fld.parent.parent.parent;
                mergeFields[sfdc_childobj][sfdc_colname] = obj;
            });
        });
        return mergeFields;
    }
}());