activeDocument.suspendHistory("PNGの高速書き出しスクリプト", "main()");

var folderPath;

function main() {

    setting();

    if(uDlg.optionSelected.value){
        if (activeDocument.activeLayer.typename == "LayerSet") {
            var childLayers = activeDocument.activeLayer.layers;
            toInvisibleLayers(childLayers)
            folderPath = createFolder(activeDocument);
            showOnlyLayer(childLayers)
            toVisibleLayers(childLayers)
        } else {
            alert("フォルダを選択してください")
        }
    } else {
        toInvisibleLayers(activeDocument.layers);
        folderPath = createFolder(activeDocument);
        showOnlyLayer(activeDocument.layers);
        toVisibleLayers(activeDocument.layers);
    }
}

function toVisibleLayers(layers) {
    for (var i = 0; i < layers.length; i++) {
        layers[i].visible = true;
    }
}

function toInvisibleLayers(layers) {
    for (var i = 0; i < layers.length; i++) {
        layers[i].visible = false;
    }
}

function showOnlyLayer (layers) {
    for (var i = 0; i < layers.length; i++) {
        layers[i].visible = true;
        ExportLayer(layers[i].name, i+1);
        layers[i].visible = false;
    }
}

function ExportLayer(layerName, number){

    var path;

    if(uDlg.optionPrefix.value){
        //Number_レイヤー名.png
        path = folderPath  + "/" + number + "_" + layerName + ".png";
    } else {
        //レイヤー名.png
        path = folderPath  + "/" + layerName + ".png";
    }


    //現在のドキュメントの保存階層取得
    fileObj = new File(path);
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

function createFolder(doc){
    var fullPath = doc.path;
    var fileName = getOnlyFileName(doc);
    var folderPath = fullPath + "/" + fileName;

    var folderObj = new Folder(folderPath);
    folderObj.create();

    return folderPath;
}

//拡張子を除いたファイル名の取得
function getOnlyFileName(doc){
    var number = doc.name.lastIndexOf(".");
    var fileName = doc.name.slice(0,number);
    return fileName;
}



function setting(){

    var wWidth = 250;
    var posLeft = 20;
    var posWidth = 100
    var padding = 15;
    var margin = 25;
    var posHeight = 20;
    var btnMargin = 30;
    

    var okBtnWidth = 90;
    var okBtnHeight = 25;

    uDlg = new Window('dialog','Quick Export as PNG',[0,0,wWidth,padding+margin+btnMargin+okBtnHeight+padding+5]);
    uDlg.optionSelected = uDlg.add("checkbox",[posLeft,padding,0,0], "選択したフォルダを対象にする");
    uDlg.optionPrefix = uDlg.add("checkbox",[posLeft,padding+margin,0,0], "ファイル名の頭に連番を振る");
    uDlg.okBtn = uDlg.add("button",[
         (wWidth-okBtnWidth)/2
        ,padding+margin+btnMargin
        ,(wWidth-okBtnWidth)/2+okBtnWidth
        ,padding+margin+btnMargin+okBtnHeight]
        , "実行", { name:"OK"});

    uDlg.optionPrefix.value = true;
    uDlg.center();
    uDlg.show();
}