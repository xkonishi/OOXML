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