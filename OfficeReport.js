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
            word.merge(mergedata);
            word.save(reportName);
        }
    };
});