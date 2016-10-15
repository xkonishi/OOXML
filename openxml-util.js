(function(){

    /************************ openXml.Util **************************/
    openXml.Util = {};

    openXml.Util.findElements = function(rootEl, targeTtag, exceptTags) {
        let results = [];

        let except = (Array.isArray(exceptTags)) ? exceptTags : [exceptTags];
        _findElements(rootEl, targeTtag, except, results);

        return results;
    }

    /************************ inner functions **************************/

    function _findElements(rootEl, targeTtag, except, results) {

        if (rootEl.nodeType === 'Element') {
            rootEl.nodesArray.forEach(function(el, index, ar) {
                if (el.nodeType === 'Element') {
                    let xname = new Ltxml.XName(el.name.namespaceName, el.name.localName);
                    if (except.indexOf(xname) === -1) {
                        if (xname === targeTtag) {
                            results.push(el);
                        }
                        else {
                            _findElements(el, targeTtag, except, results);
                        }
                    }
                }
            });
        }
    }
}());