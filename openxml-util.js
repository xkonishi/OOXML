(function(){

    /************************ openXml.Util **************************/
    openXml.Util = {};

    /**
    * エレメントの検索
    * @param rootEl         検索開始ルートエレメント
    * @param targeTtag      検索対象タグ（Ltxml.XName型）
    * @param exceptTags     検索除外タグ（１つ、または配列で複数指定可能）
    * @return 検索結果
    */
    openXml.Util.findElements = function(rootEl, targeTtag, exceptTags) {
        let results = [];

        let except = (Array.isArray(exceptTags)) ? exceptTags : [exceptTags];
        _findElements(rootEl, targeTtag, except, results);

        return results;
    }

    /************************ inner functions **************************/

    /**
    * エレメントの検索
    * @param rootEl         検索開始ルートエレメント
    * @param targeTtag      検索対象タグ（Ltxml.XName型）
    * @param exceptTags     検索除外タグ（１つ、または配列で複数指定可能）
    * @param results        検索結果
    */
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