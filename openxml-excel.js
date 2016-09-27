(function(){
	//スプレッドシート要素名
    const S = openXml.S;
    //名前空間なしの要素名
    const NN = openXml.NoNamespace;

    //属性
    const XAttribute = Ltxml.XAttribute;
    //エレメント
    const XElement = Ltxml.XElement;
	
	let pkg;

    /************************ openXml.Excel **************************/
    openXml.Excel = function(officedoc) {
		pkg = new openXml.OpenXmlPackage(officedoc);
	};
	
}());