﻿filename = File.openDialog("新規組版：一般企画のCSVファイルを選択してください...")  //ファイル選択
fileObj = new File(filename);
flag = fileObj.open("r");  //読み込みモードで開く
if (flag == true)
{
    docObj = app.activeDocument;  //アクティブなドキュメントを選択
    categStyle = docObj.paragraphStyles.add({name:"企画分類"});
    categStyle.appliedFont = "源ノ角ゴシック";
    categStyle.fontStyle = "Medium";
    categStyle.pointSize = "8Q";  //プロパティ（フォントサイズ）設定
    
    var separator = ",\""  //区切り文字の指定
    var pageno = 1
    pageObj = docObj.pages.add(LocationOptions.AFTER, docObj.pages[0]);  //新規ページを追加
    pageObj = app.activeDocument.pages.add();  //新規ページを追加し、アクティブなページに選択
    line = fileObj.readln();  //1行読み込む
    while (!fileObj.eof){  //行がなくなるまで繰り返す
        for (i=0; i<6; i++){
            for (j=0; j<3; j++)
            {
                if (fileObj.eof){
                    break;
                    break; //空値ならば抜ける
                }else{
                    line = fileObj.readln();  //1行読み込む
                    var data = line.split(separator);  //行を「,"」で区切って配列化
                    if (valueOf(data[3]) != valueOf(pageno)){  //改ページか確認
                        pageno = parseInt(data[3])
                        i = 0;
                        j = 0;  //改ページなら新規ページを追加
                        pageObj = docObj.pages.add(LocationOptions.AFTER, docObj.pages[pageno]);  //新規ページを追加し、アクティブなページに選択
                        pageObj = app.activeDocument.pages.add();  //新規ページを追加し、アクティブなページに選択
                    }
                    //企画宣伝背景
                    selFile = new File("/C/Users/Yudai/OneDrive - 筑波大学/学園祭実行委員会/広報宣伝局2019/パンフレット/design資料/企画紹介エリア別（新幹線）.ai");  //入れる画像ファイル指定
                    imgObj = pageObj.textFrames.add();
                    imgObj.name = data[4] + data[1].slice(0, -1) + "-Img";  //レイヤー名の設定
                    x = j * 47 + 3;  //横の間隔
                    y = i * 31 +3;  //高さの間隔
                    x01 = x;  //左からの座標, 横の間隔
                    y01 = y;  //上からの座標
                    x02 = x01 + 46;  //横46㎜
                    y02 = y01 + 31.5;  //縦31㎜
                    x01 = x01 + "mm";
                    y01 = y01 + "mm";
                    x02 = x02 + "mm";
                    y02 = y02 + "mm";
                    imgObj.visibleBounds = [y01,x01,y02,x02];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    imgObj.contentType = ContentType.graphicType;  //フレームの種類をグラフィックに明示的に設定
                    imgObj.place(selFile);  //入れる画像
                    imgObj.fit(FitOptions.frameToContent);  //内容にフレームサイズを合わせる
                    //企画分類
                    txtObj = pageObj.textFrames.add();  //フレーム追加
                    txtObj.name = data[4] + data[1].slice(0, -1) + "-Category"  //レイヤー名の設定
                    x11 = x + 1  //左からの座標
                    y11 = y + 2.75  //上からの座標
                    x12 = x11 + 30;  //横30㎜
                    y12 = y11 + 2;  //縦2㎜
                    x11 = x11 + "mm";
                    y11 = y11 + "mm";
                    x12 = x12 + "mm";
                    y12 = y12 + "mm";
                    txtObj.visibleBounds = [y11,x11,y12,x12];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    txtObj.contents = data[18].slice(0, -1);  //テキストフレームに流し込む
                    txtObj.parentStory.appliedParagraphStyle = "企画分類"

                    //企画名
                    txtObj = pageObj.textFrames.add();  //フレーム追加
                    txtObj.name = data[4] + data[1].slice(0, -1) + "-Title";  //レイヤー名の設定
                    x21 = x + 1.5;  //左からの座標
                    y21 = y + 5.625;  //上からの座標
                    x22 = x21 + 30.25;  //横30.25㎜
                    y22 = y21 + 6.25;  //縦6.25㎜
                    x21 = x21 + "mm";
                    y21 = y21 + "mm";
                    x22 = x22 + "mm";
                    y22 = y22 + "mm";
                    txtObj.visibleBounds = [y21,x21,y22,x22];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    txtObj.contents = data[10].slice(0, -1);  //テキストフレームに流し込む
                    
                    //企画団体名
                    txtObj = pageObj.textFrames.add();  //フレーム追加
                    txtObj.name = data[4] + data[1].slice(0, -1) + "-Name"  //レイヤー名の設定
                    x1 = x + 1.5  //左からの座標
                    y1 = y + 12.5  //上からの座標
                    x2 = x1 + 42;  //横42㎜
                    y2 = y1 + 1.75;  //縦1.75㎜
                    x1 = x1 + "mm";
                    y1 = y1 + "mm";
                    x2 = x2 + "mm";
                    y2 = y2 + "mm";
                    txtObj.visibleBounds = [y1,x1,y2,x2];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    txtObj.contents = data[11].slice(0, -1);  //テキストフレームに流し込む
                    
                }
            }
        }
    }
fileObj.close();
}else{
alert("ファイルが開けませんでした。");
docObj = app.activeDocument; 
pageObj = docObj.pages[1];
var l = pageObj.textFrames.itemByName("21-Name");  //21-Nameという名前のテキストフレーム
var m = pageObj.textFrames.itemByName("4-Name");  //4-Nameという名前のテキストフレーム
var lC = l.contents
var mC = m.contents
l.contents =mC  //入れ替える
m.contents = lC
}
