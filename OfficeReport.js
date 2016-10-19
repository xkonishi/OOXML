 $(document).ready(function() {

    //出力ボタンのクリック
    $('table[id$="reportlist"]').find('input[type=button]').click(function() {
        //画面のブロック
        $.blockUI();

        var base = '/services/apexrest';
        if (sfdc.namespacePrefix) {
            base += '/' + sfdc.namespacePrefix;
        }

        //Webサービスの実行
        $.when(
            //Officeファイル（Base64形式）
            $.ajax(base+'/officefile/'+this.id+'?userId='+sfdc.userId,
                   {
                       beforeSend: function(xhr) {
                           xhr.setRequestHeader('Authorization', 'Bearer '+sfdc.sessionId);
                       },
                       success: function(response) {
                           console.log('OK!!:'+response);
                       },
                       error: function(jqXHR, textStatus, errorThrown) {
                           console.log(textStatus);
                       }
                   }),
            //差し込みデータ
            $.ajax(base+'/officedata/'+this.id+'?userId='+sfdc.userId,
                   {
                       beforeSend: function(xhr) {
                           xhr.setRequestHeader('Authorization', 'Bearer '+sfdc.sessionId);
                       },
                       success: function(response) {
                           console.log('OK!!:');
                       },
                       error: function(jqXHR, textStatus, errorThrown) {
                           console.log(textStatus);
                       }
                   }),
            //レポート名
            this.name,
            //ファイルタイプ
            $(this).parent().prev().find('span').text()
        )
        .done(function(fileResponse, dataResponse, reportName, fileType){
            try {
                //レポート出力
                saveOffice(fileResponse[0], dataResponse[0], reportName, fileType);
            }
            catch (e) {
                alert('Merge Error!!');
            }
        })
        .fail(function() {
            alert('Ajax Error!!');
        })
        .always(function() {
            //画面のブロック解除
            $.unblockUI();
        });
    });

    /**
    * Officeファイル（Base64形式）と差し込みデータより、レポートファイルを出力する
    * @parameter	[String] officedoc		Officeファイル（Base64形式）
    * @parameter	[Object] mergedata		差し込みデータ
    * @parameter	[String] reportName		レポート名
    * @parameter	[String] fileType		ファイルタイプ[Excel/Word]
    */
    function saveOffice(officedoc, mergedata, reportName, fileType) {

        if (fileType === 'Excel') {
            let excel = new openXml.Excel(officedoc);
            excel.merge(mergedata);
            excel.save(reportName);
        }
        else if (fileType === 'Word') {
            let word = new openXml.Word(officedoc);
            word.merge(mergedata[0]);
            word.save(reportName);
        }


/*
        //Base64形式のOfficeファイルを読み込み
        const pkg = new openXml.OpenXmlPackage(officedoc);
        const main = pkg.mainDocumentPart();
        const mnXDoc = main.getXDocument();

        //各パーツの取得
        var workbookPart = pkg.workbookPart();
        var worksheetPart = workbookPart.worksheetParts()[0];
        var tablePart = worksheetPart.tableDefinitionParts()[0];
        var xmlmapPart = pkg.getPartByUri('/xl/xmlMaps.xml');

        //Excel操作オブジェクト
        var excel = new openXml.Excel();

        //シートデータのクリア
        excel.clearSheetData(worksheetPart, tablePart);

        //テーブルの表示範囲を設定
        excel.setTableRange(tablePart, mergedata.length);

        //差し込みデータの挿入
        excel.mergeSheetData(mergedata, worksheetPart, tablePart, xmlmapPart);

        //レポートファイルの出力
        pkg.saveToBlobAsync(function (blob) {
            saveAs(blob, reportName+'.xlsx');
        });
*/
    };
});