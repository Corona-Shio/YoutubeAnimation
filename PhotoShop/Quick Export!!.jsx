// PhotoShop Script

activeDocument.suspendHistory("Quick Export", "main()");

function main() {
    var fileName = getFileNameExcludedFileExtension(activeDocument);
    var layerName = activeDocument.activeLayer.name;
    ExportLayer(fileName, layerName);
}

function ExportLayer(fileName, layerName){
    //現在のドキュメントの保存階層取得
    var fullPath = app.activeDocument.path;
    fileObj = new File(fullPath + "/" + fileName + "_" + layerName+ ".png");
    //PNGファイル作成　～PNGを保存する設定を宣言～
    pngOpt = new PNGSaveOptions();
    //1.インターレースを有効にするのか　true：する　false：しない
    pngOpt.interlaced = false;
    //2.圧縮設定。0～9の数値？　デフォルトは0　9になるほど圧縮される
    pngOpt.compression = 0;
    //3.上記の設定で現在のドキュメントを保存する
    activeDocument.saveAs(fileObj, pngOpt, true, Extension.LOWERCASE);
    //　※fileObjにpngOptの設定で同名で保存。拡張子は小文字で。
    //　※trueの部分をfalseにすると、別名保存になる。
    //　　⇒実行すると保存ポップアップが表れます。
}

//拡張子を除いたファイル名の取得
function getFileNameExcludedFileExtension(doc){
    var number = doc.name.lastIndexOf(".");
    var fileName = doc.name.slice(0,number);
    return fileName;
}