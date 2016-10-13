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

    openXml.Word.prototype.getMergeField = function() {
        let mergeFields = [];

        const body = mnXDoc.root.element(openXml.W.body);
        const els = body.nodesArray;

        for (let i=0; i<els.length; i++) {
            let el = els[i];
            if (el.name.localName === 'p') {
                let fld = el.element(openXml.W.fldSimple);
                if (fld) {
                    let obj = {}
                    obj.el = el;
                    obj.fieldName = fld.element(openXml.W.r).element(openXml.W.t).value;
                    mergeFields.push(obj);
                }
            }
        }
        return mergeFields;
    };

    openXml.Word.prototype.save = function(reportName) {
        pkg.saveToBlobAsync(function (blob) {
            saveAs(blob, reportName+'.docx');
        });
    };

}());