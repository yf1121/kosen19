//id_findChildrenStyles.jsx　2017.01.27（z-）
//選択テキストに適用されている段落スタイルがベースになっている段落スタイル、
//「次のスタイル」に使用している段落スタイルを探して名前だけお知らせするやつ

var doc = app.activeDocument;
var sel = doc.selection[0];
try{
    var aps = sel.paragraphs[0].appliedParagraphStyle;
    }
catch(e){
    alert("あなたは段落スタイルが適用されたテキストを選択してから実行します"); exit();
    }

var prstAry = doc.allParagraphStyles;
var basedAry = [], nextAry = [];

for (var i = 2; i < prstAry.length; i++){
    if(prstAry[i].basedOn.reflect.name == "ParagraphStyle"){
        if(prstAry[i].basedOn.id == aps.id){
            basedAry.push(nameFunc(prstAry[i]));
            }
        if(prstAry[i].nextStyle.id == aps.id){
            nextAry.push(nameFunc(prstAry[i]));
            }
        }
    }

var result = "基準にされている：\n" + (basedAry.length ? basedAry.join("\n") : "なし") 
               + "\n次のスタイルにされている：\n" + (nextAry.length ? nextAry.join("\n") : "なし");
alert(result);

function nameFunc(obj){
    var str = obj.name;
    var par = obj;
    while(par.parent != doc){
        par = par.parent;
        str += " (" + par.name + ")";
        }
    return str;
    }
