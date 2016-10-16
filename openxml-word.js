(function(){

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
        let mergeFields = getMergeFields();

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
                    for (let i=1; i<obj.records.length; i++) {
                        tr.parent.add(new Ltxml.XElement(tr));
                    }

                    for (let i=0; i<obj.records.length; i++) {
                        let data = obj.records[i];

                        Object.keys(data).forEach(function(key, index, ar) {
                            if ((typeof data[key]) !== 'object') {
                                let val = data[key];
                                if (mergeFields[sfdc_childobj][key]) {
                                    let el = mergeFields[sfdc_childobj][key].el;
                                    if (el) {
                                        //差し込みフィールドの削除、およびテキスト挿入
                                        el.element(openXml.W.fldSimple).remove();
                                        el.add(new Ltxml.XElement(openXml.W.r, new Ltxml.XElement(openXml.W.t, val)));
                                    }
                                }
                            }
                        });
                        break;//とりあえず
                    }
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

    function getMergeFields() {
        let mergeFields = {};

        let body = mnXDoc.root.element(openXml.W.body);

        let ps = body.elements(openXml.W.p);
        ps.forEach(function(p, index, ar) {

            let fld = p.element(openXml.W.fldSimple);
            if (fld) {
                let obj = {}
                obj.el = p;
                obj.fieldName = fld.element(openXml.W.r).element(openXml.W.t).value;
                obj.sfdcInfo = obj.fieldName.match(/[A-Za-z\.]+/g);

                let sfdc_colname = obj.sfdcInfo[1];
                mergeFields[sfdc_colname] = obj;
            }
        });

        let tbls = body.elements(openXml.W.tbl);
        tbls.forEach(function(tbl, index, ar) {

            let trs = tbl.elements(openXml.W.tr);
            trs.forEach(function(tr, index, ar) {

                let sfdc_childobj = '';

                let tcs = tr.elements(openXml.W.tc);
                for (let i=0; i<tcs.count(); i++) {
                    let tc = tcs.elementAt(i);

                    let p = tc.element(openXml.W.p);
                    if (p) {
                        let fld = p.element(openXml.W.fldSimple);
                        if (fld) {
                            let obj = {}
                            obj.el = p;
                            obj.fieldName = fld.element(openXml.W.r).element(openXml.W.t).value;
                            obj.sfdcInfo = obj.fieldName.match(/[A-Za-z]+/g);

                            sfdc_childobj = obj.sfdcInfo[1];
                            if (!mergeFields[sfdc_childobj]) {
                                mergeFields[sfdc_childobj] = {};
                            }

                            let sfdc_colname = obj.sfdcInfo[2];
                            mergeFields[sfdc_childobj].el = tr;
                            mergeFields[sfdc_childobj][sfdc_colname] = obj;
                        }
                    }
                }

                if (sfdc_childobj) {
                    return true;//break
                }
            });
        });

        return mergeFields;
    }
}());