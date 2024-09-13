(function () {
    // エラーログファイルの作成
    var eLog = new File("~/Desktop/error_log.txt");
    eLog.open("w");
    eLog.writeln("エラーログ: " + new Date());
    
    // フォルダ選択ダイアログを表示
    var sourceFolder = Folder.selectDialog("選択するフォルダを選んでください");
    if (sourceFolder === null) {
        alert("フォルダが選択されていません。");
        eLog.writeln("フォルダが選択されていません。");
        eLog.close();
        return;
    }

    var files = sourceFolder.getFiles("*.indd");

    if (files.length === 0) {
        alert("選択したフォルダに .indd ファイルが見つかりません。");
        eLog.writeln("選択したフォルダに .indd ファイルが見つかりません。");
        eLog.close();
        return;
    }

    // 処理中ダイアログの作成
    var dialogLoad = new Window("palette", "処理中");
    dialogLoad.add("statictext", undefined, "処理中...");
    dialogLoad.show();

    // デスクトップに Links フォルダを作成
    var desktopFolder = Folder("~/Desktop/Links");
    if (!desktopFolder.exists) desktopFolder.create();

    for (var i = 0; i < files.length; i++) {
        try {
            processDocument(files[i], desktopFolder);
        } catch (e) {
            eLog.writeln("ドキュメントの処理中にエラーが発生しました: " + files[i].name + " - エラー: " + e.message);
        }
    }

    dialogLoad.close();
    alert("処理が完了しました。");
    eLog.close();

    function processDocument(docFile, linksFolder) {
        var doc;
        try {
            doc = app.open(docFile);
        } catch (e) {
            eLog.writeln("ドキュメントのオープン中にエラーが発生しました: " + docFile.name + " - エラー: " + e.message);
            return;
        }
        
        collectLinks(doc, linksFolder);
        doc.close(SaveOptions.NO);
    }

    function collectLinks(doc, linksFolder) {
        var links = doc.links;
        if (links.length === 0) {
            return;
        }
        var collectedLinks = {};

        // リンクの処理
        for (var i = 0; i < links.length; i++) {
            var link = links[i];
            var sourceFile = new File(link.filePath);
            if (sourceFile.exists) {
                var destFile = new File(linksFolder + "/" + link.name);
                copyFile(sourceFile, destFile);
                collectedLinks[link.name] = true;
                collectSubLinks(destFile.fsName, linksFolder, collectedLinks, sourceFile.path);
            } else {
                eLog.writeln("リンクが見つかりません: " + link.filePath);
            }
        }
    }

    function collectSubLinks(filePath, linksFolder, collectedLinks, basePath) {
        var subLinks = getLinks(filePath);
        for (var j = 0; j < subLinks.length; j++) {
            var subLinkName = subLinks[j];
            if (!collectedLinks[subLinkName]) {
                var subLinkFile = new File(basePath + "/" + subLinkName);
                if (subLinkFile.exists) {
                    var subDestFile = new File(linksFolder + "/" + subLinkName);
                    copyFile(subLinkFile, subDestFile);
                    collectedLinks[subLinkName] = true;
                    collectSubLinks(subDestFile.fsName, linksFolder, collectedLinks, subLinkFile.path);
                }
            }
        }
    }
    // リンクのコピーを行う関数
    function copyFile(moto, saki) {
        var d1 = app.documents.add();
        var p = d1.spreads[0].place(File(moto));
        p[0].itemLink.copyLink(File(saki));
        d1.close(SaveOptions.NO);
    }
    // AdobeXMPScriptの初期化とリンク取得関数の定義
    function getLinks(fls) {
        var prop = "Manifest";
        var ns = "http://ns.adobe.com/xap/1.0/mm/";
        if (typeof xmpLib === 'undefined') {
            var xmpLib = new ExternalObject('lib:AdobeXMPScript');
        }
        var xmpFile = new XMPFile(fls, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_READ);
        var xmpPackets = xmpFile.getXMP();
        var xmp = new XMPMeta(xmpPackets.serialize());
        var str = "";
        var result = [];
        for (var i = 1; i <= xmp.countArrayItems(ns, prop); i++) {
            str = xmp.getProperty(ns, prop + "[" + i + "]/stMfs:reference/stRef:filePath").toString();
            result.push(str.slice(str.lastIndexOf("/") + 1));
        }
        return result;
    }
    
})();
