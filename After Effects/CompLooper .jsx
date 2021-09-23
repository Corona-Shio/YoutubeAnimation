// After Effects Script

var frame = prompt("何フレームでループさせる？", "数字を入力してください");
var frameRate = app.project.activeItem.frameRate; // 全て同じフレームレートと仮定する

var targetLayer = app.project.activeItem.selectedLayers[0];

if(targetLayer.source instanceof CompItem){
    main();
} else {
    alert("コンポジションレイヤーを選択してください");
}



function main(){

    // アクティブなコンポジションレイヤーをコンポジションとして取得
    var childComp = app.project.activeItem.selectedLayers[0].source

    adjustPropertyForLoop(childComp);
    loopCompLayer(targetLayer);


}

function loopCompLayer(compLayer){

    var countLayer = compLayer.source.layers.length;

    if(compLayer.collapseTransformation){
        compLayer.collapseTransformation = true;
    }
    
    if(hasKey(compLayer.timeRemap)){
        removeAllKeyFrames(compLayer.timeRemap);
    }

    if(compLayer.canSetTimeRemapEnabled){
        compLayer.timeRemapEnabled = true;
    }
    
    // タイムリマップのキーフレーム設定
    compLayer.timeRemap.setValueAtTime(countLayer*frame/frameRate, [0])
    compLayer.timeRemap.setValueAtTime((countLayer*frame/frameRate)-(1/frameRate), [(countLayer*frame/frameRate)-(1/frameRate)])
    
    compLayer.timeRemap.expression = "loopOut()"
}

// コンプの中身を指定フレームごとにズラす
function adjustPropertyForLoop(comp){
    for(var i=1; i<=comp.layers.length; i++){

        if(hasKey(comp.layer(i).transform.opacity)){
            removeAllKeyFrames(comp.layer(i).transform.opacity);
        }

        comp.layer(i).startTime = 0;
        comp.layer(i).transform.opacity.setValueAtTime(0,[100]);
        comp.layer(i).transform.opacity.setValueAtTime(frame/frameRate,[0]);

        comp.layer(i).transform.opacity.setInterpolationTypeAtKey(1, KeyframeInterpolationType.HOLD)
        comp.layer(i).transform.opacity.setInterpolationTypeAtKey(2, KeyframeInterpolationType.HOLD)

        comp.layer(i).startTime = (i-1)*frame/frameRate;
    }
    comp.duration = frame/frameRate*comp.layers.length
}

function hasKey(property){
    if(property.numKeys>=1){
        return true;
    } else{
        return false;
    }
}

function removeAllKeyFrames(property){
    for(i=property.numKeys; property.numKeys>0 ;i--){
        property.removeKey(i);
    }
}