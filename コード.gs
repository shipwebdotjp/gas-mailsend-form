//使用者リストの列順番
var userSheetIndex = [
    "isSend", "group", "name", "gender", "mail", "message", "add1", "add2", "attachment1"
];


let settings = [];

//スプレッドシート起動時に実行（メニューの登録）
function onOpen() {
    SpreadsheetApp
        .getActiveSpreadsheet()
        .addMenu('メール', [
            { name: '一斉送信', functionName: 'createSidebar' },
            { name: '添付ファイル作成', functionName: 'createAttachment' },
        ]);
}

//使用者シートのカラムインデックスを返す row: カラム名
function getUserSheetIndex(row) {
    return userSheetIndex.indexOf(row)
}

//グループを返す
function getGroups() {
    var mySheet = SpreadsheetApp.getActiveSheet();
    var selectedRow = mySheet.getRange(2, 2, mySheet.getLastRow() - 1, 1).getValues(); //グループ列全体を取得
    var groups = ["全て"];
    for (var i = 0; i < selectedRow.length; i++) {
        if (selectedRow[i][0] != "") {
            groups.push(selectedRow[i][0]);
        }
    }
    return JSON.stringify(Array.from(new Set(groups)));
}

//サイドバーを作成する
function createSidebar() {
    var htmlOutput = HtmlService.createTemplateFromFile('mailform').evaluate();
    htmlOutput.setTitle('一斉送信');
    SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

//get settings
function getSettings() {
    const settingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("設定");
    return {
        "attachmentFolderId": settingSheet.getRange("B1").getValue(),
        "templateDocumentId": settingSheet.getRange("B2").getValue(),
        "senderName": settingSheet.getRange("B3").getValue(),
        "senderMail": settingSheet.getRange("B4").getValue(),
    }
}

//メールの送信
function processSendEmail(form) {
    settings = getSettings();
    const attachmentFolder = DriveApp.getFolderById(settings["attachmentFolderId"]);

    if (form["group"].length == 0) {
        return { "error": true, "message": "グループを選択してください。" }
    }
    if (form["subject"] == "") {
        return { "error": true, "message": "件名を入力してください。" }
    }
    if (form["body"] == "") {
        return { "error": true, "message": "本文を入力してください。" }
    }

    let attachmentBlobs = [];
    if (form["attachment"] != null) {
        let attachment_size_total = 0;
        for (var i = 0; i < form["attachment"].length; i++) {
            var reader_result = form["attachment"][i]["data"];  //アップロードされたデータ
            var file_name = form["attachment"][i]["name"];  //ファイル名
            var result_split = reader_result.split(','); //base64エンコードされたファイルをコンマで分割
            var content_type = result_split[0].split(';')[0].replace('data:', ''); //ファイルのMINEタイプ
            var row_data = result_split[1]; //ファイルの実データ
            var data = Utilities.base64Decode(row_data); //base64エンコードされた文字列からバイナリへデコード
            const new_blob = Utilities.newBlob(data, content_type, file_name);
            attachmentBlobs.push(new_blob);
            attachment_size_total += new_blob.getBytes().length;
            //console.log(attachmentBlob.getName() + " ContentType: " + attachmentBlob.getContentType() + " size: "+attachmentBlob.getBytes().length);
        }
        //所有附件的容量超过25MB



    }
    var mySheet = SpreadsheetApp.getActiveSheet();  //アクティブシートを取得
    if (!mySheet) {
        return { "error": true, "message": "シートがありません。" }
    }

    //送信元メールアドレスが、メールエイリアスとして登録されているかチェック
    if (settings["senderMail"] != "") {
        let aliases = GmailApp.getAliases();
        if (aliases.indexOf(settings["senderMail"]) === -1 && Session.getActiveUser().getUserLoginId() != settings["senderMail"]) {
            return { "error": true, "message": '送信元メールアドレス' + settings["senderMail"] + 'はメールエイリアスとして設定されていません。' }
        }
    }

    console.log("Sheet:" + mySheet.getName() + " Group:" + form["group"] + " Subject:" + form["subject"] + " Body:" + form["body"]);

    var regexp = /^[A-Za-z0-9]{1}[A-Za-z0-9_.-]*@{1}[A-Za-z0-9_.-]{1,}\.[A-Za-z0-9]{1,}$/; //メールアドレスの正規表現
    var selectedRow = mySheet.getRange(2, 1, mySheet.getLastRow() - 1, userSheetIndex.length).getValues(); //アクティブな行全体を取得
    var sendque = []; //送信予定リスト
    var sended = [];  //送信済みリスト
    var errors = [];  //エラーリスト
    for (var i = 0; i < selectedRow.length; i++) {
        if ((form["group"].indexOf("全て") > -1 || form["group"].indexOf(selectedRow[i][getUserSheetIndex("group")]) > -1) && selectedRow[i][getUserSheetIndex("isSend")] == "1" && selectedRow[i][getUserSheetIndex("mail")] != "") {
            userData = {
                "name": selectedRow[i][getUserSheetIndex("name")] + selectedRow[i][getUserSheetIndex("gender")],
                "gender": selectedRow[i][getUserSheetIndex("gender")],
                "mail": selectedRow[i][getUserSheetIndex("mail")].trim(),
                "message": selectedRow[i][getUserSheetIndex("message")],
                "add1": selectedRow[i][getUserSheetIndex("add1")],
                "add2": selectedRow[i][getUserSheetIndex("add2")],
                "attachment1": selectedRow[i][getUserSheetIndex("attachment1")],
            }
            var subject = replaceString(form["subject"], userData);
            var body = replaceString(form["body"], userData);
            var htmlbody = replaceStringHTML(form["htmlbody"], userData);
            const attachment = getGoogleDriveFile(userData["attachment1"], attachmentFolder);

            if (regexp.test(userData["mail"])) {
                var message = {
                    to: selectedRow[i][getUserSheetIndex("name")] + " <" + userData["mail"] + ">",
                    subject: subject,
                    body: body,
                    htmlBody: htmlbody,
                    attachments: attachment ? [attachment.getBlob()] : [],
                }
                sendque.push(message);
            } else {
                errors.push(userData["name"] + " <" + userData["mail"] + ">"); //メールアドレスの形式が正しくない場合はエラーリストへ追加
            }
        }
    }

    if (errors.length > 0) {  //正しくないメールアドレスが含まれていた場合
        console.log('次のメールアドレスが正しくありません：' + errors.join(','));
        return { "error": true, "message": '次のメールアドレスが正しくありません：' + errors.join(',') }
    }

    if (sendque.length > MailApp.getRemainingDailyQuota()) { //送信予定数が残りの一日の送信可能数を超える場合
        console.log('一日の送信可能数を超えるため送信できません。本日残り可能送信数：' + MailApp.getRemainingDailyQuota());
        return { "error": true, "message": '一日の送信可能数を超えるため送信できません。本日残り可能送信数：' + MailApp.getRemainingDailyQuota() }
    }

    for (var i = 0; i < sendque.length; i++) { //送信先リストからメール送信
        if (settings["senderMail"] != "") {
            sendque[i]["from"] = settings["senderMail"];
        }
        if (settings["senderName"] != "") {
            sendque[i]["name"] = settings["senderName"];
        }
        if (attachmentBlobs.length > 0) {
            //merge attachment
            sendque[i]["attachments"] = sendque[i]["attachments"].concat(attachmentBlobs);
        }
        try {
            MailApp.sendEmail(sendque[i]);
            sended.push(sendque[i]["to"]);
        } catch (e) {
            console.log(e.message);
            errors.push(sendque[i]["to"]);
        }

    }

    if (sended.length > 0) {
        console.log("メールを送信しました。" + sended.join(','));
        return { "error": false, "message": "メールを送信しました。", "sended": sended, "errors": errors, "quota": MailApp.getRemainingDailyQuota() }
    } else {
        return { "error": true, "message": "該当する送信先がありません。", "errors": errors }
    }
}

function createAttachment() {
    settings = getSettings();
    const attachmentFolder = DriveApp.getFolderById(settings["attachmentFolderId"]);
    const templateDocument = getGoogleDriveFile(settings["templateDocumentId"], attachmentFolder);
    //get active sheet
    const mySheet = SpreadsheetApp.getActiveSheet();
    //get active rows from active sheet
    const selectedRow = mySheet.getRange(2, 1, mySheet.getLastRow() - 1, userSheetIndex.length).getValues();
    //loop each selected rows
    for (let i = 0; i < selectedRow.length; i++) {
        if (selectedRow[i][getUserSheetIndex("isSend")] == "1") {
            //make copy of template document
            const copyOftemplateDocument = templateDocument.makeCopy(templateDocument.getName() + "_copy");
            //get document
            const doc = DocumentApp.openById(copyOftemplateDocument.getId());
            //replace list
            const replaceLists = {
                "name": selectedRow[i][getUserSheetIndex("name")] + selectedRow[i][getUserSheetIndex("gender")],
                "gender": selectedRow[i][getUserSheetIndex("gender")],
                "mail": selectedRow[i][getUserSheetIndex("mail")].trim(),
                "message": selectedRow[i][getUserSheetIndex("message")],
                "add1": selectedRow[i][getUserSheetIndex("add1")],
                "add2": selectedRow[i][getUserSheetIndex("add2")],
                "attachment1": selectedRow[i][getUserSheetIndex("attachment1")],
            }
            //replace content
            const replacedDoc = replaceContent(doc, replaceLists);
            //get as PDF
            const pdfBlob = replacedDoc.getAs('application/pdf');
            //set file name 
            const pdfFileName = replaceString(templateDocument.getName(), replaceLists) + ".pdf";
            pdfBlob.setName(pdfFileName);
            //save blob to attachment folder
            const pdfFile = attachmentFolder.createFile(pdfBlob);
            //set filename to sheet
            mySheet.getRange(i + 2, getUserSheetIndex("attachment1") + 1).setValue(pdfFile.getId());
            //delete replaced document
            DriveApp.getFileById(replacedDoc.getId()).setTrashed(true);
        }
    }
}

//replace content of document
function replaceContent(doc, replaceLists) {
    //get body and replace string
    const body = doc.getBody();
    Object.keys(replaceLists).forEach(key => {
        body.replaceText(`\{${key}\}`, replaceLists[key]);
    })
    //save and close
    doc.saveAndClose();
    return doc;
}

//replace string
function replaceString(str, replaceLists) {
    Object.keys(replaceLists).forEach(key => {
        str = str.replace(new RegExp(`\{${key}\}`, "gi"), replaceLists[key]);
    })
    return str;
}

//replace string to HTML
function replaceStringHTML(str, replaceLists) {
    Object.keys(replaceLists).forEach(key => {
        str = str.replace(new RegExp(`\{${key}\}`, "gi"), nl2br(replaceLists[key]));
    })
    return str;
}

//recognize Google drive file ID or URL or filename and return file ID
function getGoogleDriveFile(str, attachmentFolder) {
    if (fileId = str.match(/([-\w]{25,}(?!.*[-\w]{25,}))/)) {
        return DriveApp.getFileById(fileId[1]);
    } else if (str != "") {
        //get file by file name
        return getFileByName(str, attachmentFolder);
    }
    return null
}

// 特定のフォルダからファイル名でファイルを検索
function getFileByName(file_name, folder) {
    var files = DriveApp.getFolderById(folder.getId()).getFilesByName(file_name);
    while (files.hasNext()) {
        // 一つ目のファイルを返す（複数存在した場合は考慮しない）
        return files.next();
    }
    return null;
}

// nl2br
function nl2br(str) {
    str = str.toString();
    return str.replace(/([^>\r\n]?)(\r\n|\n\r|\r|\n)/g, '$1<br>$2');
}

