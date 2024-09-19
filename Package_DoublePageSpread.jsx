main();

function main() {
  var folder = Folder.selectDialog("フォルダを選択してください");
  if (!folder) {
    alert("フォルダが選択されていません");
    return;
  }

  var dialog = new Window("dialog", "タイトルを入力してください");
  var inputField = dialog.add("edittext", undefined, "");
  inputField.characters = 30; // フィールドの幅を設定
  var buttonGroup = dialog.add("group");
  buttonGroup.add("button", undefined, "OK", { name: "ok" });
  buttonGroup.add("button", undefined, "キャンセル", { name: "cancel" });
  if (dialog.show() == 1) {
    var userInput = inputField.text;

    // 処理中のダイアログを作成
    var dialogLoad = new Window("palette", "処理中");
    dialogLoad.add("statictext", undefined, "処理中...");
    dialogLoad.show();
    // フォルダ内の .indd ファイルを取得
    var files = folder.getFiles("*.indd");
    if (files.length === 0) {
      alert("フォルダ内に .indd ファイルがありません");
      dialogLoad.close();
      return;
    }

    // スプレッドを分割してリンクを収集
    spledSplit(folder, userInput, files);
    linkSyusyu(folder);

    dialogLoad.close();
    alert("処理が完了しました");
  } else {
    alert("処理がキャンセルされました");
  }
}

function spledSplit(folder, userInput, files) {
    for (var f = 0; f < files.length; f++) {
      var doc = app.open(files[f], false);
      if (doc.saved === false || doc.modified === true) {
        alert("ドキュメントを保存してから実行してください: " + files[f].name);
        continue;
      }
      // ドキュメントページの移動を許可しない
      doc.documentPreferences.allowPageShuffle = false;
      var spr_obj = doc.spreads;
      for (var i = 0, iL = spr_obj.length; i < iL; i++) {
        spr_obj[i].allowPageShuffle = false;
        var pag_obj = spr_obj[i].pages;
        var spr_name = pag_obj.length > 1 ? "0" + pag_obj[0].name + "-0" + pag_obj[pag_obj.length - 1].name : pag_obj[0].name;
        spr_obj[i].insertLabel('sp_id', "" + i);
        spr_obj[i].insertLabel('sp_name', spr_name);
        spr_obj[i].insertLabel('start_p_num', pag_obj[0].name.replace(pag_obj[0].appliedSection.name, ""));
      }
      var org_doc_path = doc.fullName;
      var iiL = spr_obj.length;
      var new_fd_path = folder.fsName + "/00_SplitData/";
      Folder(new_fd_path).create();
      // バックアップを作成
      doc.close(SaveOptions.NO, new File(org_doc_path + ".bk"));
      // 再度バックアップを開く
      for (var is = 0; is < iiL; is++) {
        var doc2 = app.open(File(org_doc_path + ".bk"), false);
        var spr_obj2 = doc2.spreads;
        // 他のスプレッドを削除
        for (var ii = iiL - 1; ii >= 0; ii--) {
          if (spr_obj2[ii].extractLabel('sp_id') !== "" + is) {
            spr_obj2[ii].remove();
          }
        }
        // ページ番号のスタートを設定
        doc2.sections[0].continueNumbering = false;
        doc2.sections[0].pageNumberStart = doc2.spreads[0].extractLabel('start_p_num') * 1;
        // パッケージ機能を使用して保存
        var packageFolder = new Folder(new_fd_path + userInput + "_" + spr_obj2[0].extractLabel('sp_name'));
        packageFolder.create();
        doc2.packageForPrint(packageFolder, true, true, false, false, false, false, false);
        //保存先,Fonts,Links,指示書,更新,非表示レイヤー,プリフライトエラー,レポート
        doc2.close(SaveOptions.NO, new File(packageFolder.fsName + "/" + userInput + "_" + spr_obj2[0].extractLabel('sp_name') + ".indd"));
      }      
    }
    // packageFolder.fsName 内の .bk ファイルを削除
    File(org_doc_path + ".bk").remove();
    function removeBkFiles(folder) {
      var files = folder.getFiles();
      for (var i = 0; i < files.length; i++) {
        if (files[i] instanceof Folder) {
          removeBkFiles(files[i]);
        } else if (files[i].name.match(/\.bk$/)) {
          files[i].remove();
        }
      }
    }
    var splitDataFolder = new Folder(folder.fsName + "/00_SplitData/");
    if (splitDataFolder.exists) {
      removeBkFiles(splitDataFolder);
    }
};


function linkSyusyu(folder) {
  processFolders(new Folder(folder.fsName + "/00_SplitData/"));
}

function processFolders(folder) {
  var subFolders = folder.getFiles(function (file) {
    return file instanceof Folder;
  });
  for (var i = 0; i < subFolders.length; i++) {
    var inddFiles = subFolders[i].getFiles("*.indd");
    var linksFolder = new Folder(subFolders[i].fsName + "/Links");
    // 既にLinksフォルダが存在する場合は作成しない
    if (!linksFolder.exists) {
      linksFolder.create(); // Linksフォルダを作成
    }
    for (var j = 0; j < inddFiles.length; j++) {
      processDocument(inddFiles[j], linksFolder);
    }
  }
}

function processDocument(docFile, linksFolder) {
  var doc;
  try {
    doc = app.open(docFile, false);
  } catch (e) {
    alert("ドキュメントのオープン中にエラーが発生しました: " + docFile.name + " - エラー: " + e.message);
    return;
  }
  collectLinks(doc, linksFolder);
  doc.close(SaveOptions.NO);
}

function collectLinks(doc, linksFolder) {
  var links = doc.links;
  if (!links || links.length === 0) {
    return;
  }
  var collectedLinks = {};
  for (var i = 0; i < links.length; i++) {
    var link = links[i];
    var sourceFile = new File(link.filePath);
    if (sourceFile.exists) {
      collectSubLinksFromFile(sourceFile, linksFolder, collectedLinks);
    } else {
      alert("リンクが見つかりません: " + link.filePath);
    }
  }
}

function collectSubLinksFromFile(file, linksFolder, collectedLinks) {
  var subLinks = getLinks(file.fsName);
  for (var j = 0; j < subLinks.length; j++) {
    var subLinkName = subLinks[j];
    if (!collectedLinks[subLinkName]) {
      var subLinkFile = new File(file.path + "/" + subLinkName);
      if (subLinkFile.exists) {
        var subDestFile = new File(linksFolder.fsName + "/" + subLinkName);
        try {
          copyFile(subLinkFile.fsName, subDestFile.fsName);
          collectedLinks[subLinkName] = true;
          // Recursive call to collect further sub-links
          collectSubLinksFromFile(subLinkFile, linksFolder, collectedLinks);
        } catch (e) {
          alert("ファイルのコピー中にエラーが発生しました: " + e.message);
        }
      }
    }
  }
}

//リンクを移動させる
function copyFile(moto, saki) {
  var destFile = new File(saki);
  if (destFile.exists) {
    return;
  }
  var d1 = app.documents.add();
  try {
    var p = d1.spreads[0].place(File(moto));
    p[0].itemLink.copyLink(destFile);
  } catch (e) {
    alert("ファイルのコピー中にエラーが発生しました: " + e.message);
  } finally {
    d1.close(SaveOptions.NO);
  }
}

function getLinks(filePath) {
  var prop = "Manifest";
  var ns = "http://ns.adobe.com/xap/1.0/mm/";
  var xmpFile = new XMPFile(filePath, XMPConst.UNKNOWN, XMPConst.OPEN_FOR_READ);
  var xmpPackets = xmpFile.getXMP();
  var xmp = new XMPMeta(xmpPackets.serialize());
  var result = [];
  for (var i = 1; i <= xmp.countArrayItems(ns, prop); i++) {
    var str = xmp.getProperty(ns, prop + "[" + i + "]/stMfs:reference/stRef:filePath").toString();
    result.push(str.slice(str.lastIndexOf("/") + 1));
  }
  return result;
}

function daialogTojiru() {
  alert("処理が完了しました");
}
