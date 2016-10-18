(function(){

    //パッケージオブジェクト
    let pkg;
    //ワークブックパーツ（xl/workbook.xml）
    let workbookPart;
    //ワークシートパーツ（xl/worksheets/sheet[n].xml）
    let worksheetPart;
    //テーブルパーツ（xl/tables/table[n].xml）
    let tablePart;
    //マップパーツ（xl/xmlMaps.xml）
    let xmlmapPart;

    /************************ openXml.Excel **************************/

    /**
    * コンストラクタ
    * @param [String] officedoc		Officeファイル（Base64形式）
    */
    openXml.Excel = function(officedoc) {
        pkg = new openXml.OpenXmlPackage(officedoc);
        workbookPart = pkg.workbookPart();
        worksheetPart = workbookPart.worksheetParts()[0];
        tablePart = worksheetPart.tableDefinitionParts()[0];
        xmlmapPart = pkg.getPartByUri('/xl/xmlMaps.xml');
    };

    /**
    * 差し込みデータの挿入
    * @param [Object] mergedata		差し込みデータ
    */
    openXml.Word.prototype.merge = function(mergedata) {
    };

    /**
    * レポートファイルの出力
    * @param [String] reportName		レポート名
    */
    openXml.Word.prototype.save = function(reportName) {
        pkg.saveToBlobAsync(function (blob) {
            saveAs(blob, reportName+'.xlsx');
        });
    };

    /************************ inner functions **************************/

}());