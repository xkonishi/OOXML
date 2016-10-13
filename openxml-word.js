(function(){

    /************************ openXml.Word **************************/
    let pkg;
    let mnPart;
    let mnXDoc;
    let mergeFields;

    openXml.Word = function(officedoc) {
        pkg = new openXml.OpenXmlPackage(officedoc);
        mnPart = pkg.mainDocumentPart();
        mnXDoc = mnPart.getXDocument();
    };

    openXml.Word.prototype.merge = function(mergedata) {
        mergeFields = getMergeFields();
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
                obj.sfdcInfo = obj.fieldName.match(/[A-Za-z]+/g);

                let sfdc_colname = obj.sfdcInfo[1];
                mergeFields[sfdc_colname] = obj;
            }
        });

        let tbls = body.elements(openXml.W.tbl);
        tbls.forEach(function(tbl, index, ar) {

            let trs = tbl.elements(openXml.W.tr);
            if (trs.count() > 0) {
                let tr = trs.elementAt(0);

                let tcs = tr.elements(openXml.W.tc);
                tcs.forEach(function(tc, index, ar) {

                    let p = tc.element(openXml.W.p);
                    if (p) {
                        let fld = p.element(openXml.W.fldSimple);
                        if (fld) {
                            let obj = {}
                            obj.el = p;
                            obj.fieldName = fld.element(openXml.W.r).element(openXml.W.t).value;
                            obj.sfdcInfo = obj.fieldName.match(/[A-Za-z]+/g);

                            let sfdc_colname = obj.sfdcInfo[2];
                            row[sfdc_colname] = obj;

                            sfdc_childobj = obj.sfdcInfo[1];
                        }
                    }
                });
            }






            trs.forEach(function(tr, index, ar) {

                let row = {};
                let sfdc_childobj = '';

                let tcs = tr.elements(openXml.W.tc);
                tcs.forEach(function(tc, index, ar) {

                    let p = tc.element(openXml.W.p);
                    if (p) {
                        let fld = p.element(openXml.W.fldSimple);
                        if (fld) {
                            let obj = {}
                            obj.el = p;
                            obj.fieldName = fld.element(openXml.W.r).element(openXml.W.t).value;
                            obj.sfdcInfo = obj.fieldName.match(/[A-Za-z]+/g);

                            let sfdc_colname = obj.sfdcInfo[2];
                            row[sfdc_colname] = obj;

                            sfdc_childobj = obj.sfdcInfo[1];
                        }
                    }
                });

                if (sfdc_childobj) {
                    if (!mergeFields[sfdc_childobj]) {
                        mergeFields[sfdc_childobj] = [];
                    }
                    mergeFields[sfdc_childobj].push(row);
                }
            });
        });

        return mergeFields;
    }
}());