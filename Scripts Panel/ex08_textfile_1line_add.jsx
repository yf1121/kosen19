selFile = new File("/C/Users/Yudai/OneDrive - 筑波大学/学園祭実行委員会/広報宣伝局2019/パンフレット/design資料/企画紹介エリア別（新幹線）.ai");  //入れる画像ファイル指定
filename = File.openDialog("新規組版：一般企画のCSVファイルを選択してください...")  //ファイル選択
fileObj = new File(filename);
flag = fileObj.open("r");  //読み込みモードで開く
if (flag == true)
{
    docObj = app.activeDocument;  //アクティブなドキュメントを選択
    var separator = ",\""  //区切り文字の指定
    pageObj = docObj.pages.add();  //新規ページを追加し、アクティブなページに選択
    pageObj = app.activeDocument.pages.add();  //新規ページを追加し、アクティブなページに選択
    var pageno = 0
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
                    if (data[1] != pageno){  //改ページか確認
                        pageObj = app.activeDocument.pages.add();  //新規ページを追加し、アクティブなページに選択
                        pageno = data[3]
                        break;
                        break;  //改ページなら新規ページを追加して抜ける
                    }
                    //企画宣伝背景
                    imgObj = docObj.textFrames.add();
                    imgObj.name = data[1] + data[4] + "-Img";  //レイヤー名の設定
                    x = j * 47 + 3;  //横の間隔
                    y = i * 31 +3;  //高さの間隔
                    x1 = x;  //左からの座標, 横の間隔
                    y1 = y;  //上からの座標
                    x2 = x1 + 46;  //横46㎜
                    y2 = y1 + 31.5;  //縦31㎜
                    x1 = x1 + "mm";
                    y1 = y1 + "mm";
                    x2 = x2 + "mm";
                    y2 = y2 + "mm";
                    imgObj.visibleBounds = [y1,x1,y2,x2];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    imgObj.contentType = ContentType.graphicType;  //フレームの種類をグラフィックに明示的に設定
                    imgObj.place(selFile);  //入れる画像
                    imgObj.fit(FitOptions.frameToContent);  //内容にフレームサイズを合わせる
                    //企画分類
                    txtObj = pageObj.textFrames.add();  //フレーム追加
                    txtObj.name = data[1] + data[4] + "-Category"  //レイヤー名の設定
                    x1 = x  //左からの座標
                    y1 = y  //上からの座標
                    x2 = x1 + 46;  //横46㎜
                    y2 = y1 + 31.5;  //縦31㎜
                    x1 = x1 + "mm";
                    y1 = y1 + "mm";
                    x2 = x2 + "mm";
                    y2 = y2 + "mm";
                    txtObj.visibleBounds = [y1,x1,y2,x2];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    txtObj.contents = data[18].slice(0, -1);  //テキストフレームに流し込む
                    txtObj.applyParagraphStyle("企画分類", false);
                    //企画名
                    txtObj = pageObj.textFrames.add();  //フレーム追加
                    txtObj.name = data[1] + data[4] + "-Title";  //レイヤー名の設定
                    x1 = x;  //左からの座標
                    y1 = y;  //上からの座標
                    x2 = x1 + 46;  //横46㎜
                    y2 = y1 + 31.5;  //縦31㎜
                    x1 = x1 + "mm";
                    y1 = y1 + "mm";
                    x2 = x2 + "mm";
                    y2 = y2 + "mm";
                    txtObj.visibleBounds = [y1,x1,y2,x2];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    txtObj.contents = data[10].slice(0, -1);  //テキストフレームに流し込む
                    txtObj.applyParagraphStyle("企画紹介名", false);
                    //企画団体名
                    txtObj = pageObj.textFrames.add();  //フレーム追加
                    txtObj.name = data[1] + data[4] + "-Name"  //レイヤー名の設定
                    x1 = x  //左からの座標
                    y1 = y  //上からの座標
                    x2 = x1 + 46;  //横46㎜
                    y2 = y1 + 31.5;  //縦31㎜
                    x1 = x1 + "mm";
                    y1 = y1 + "mm";
                    x2 = x2 + "mm";
                    y2 = y2 + "mm";
                    txtObj.visibleBounds = [y1,x1,y2,x2];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    txtObj.contents = data[11].slice(0, -1);  //テキストフレームに流し込む
                    txtObj.applyParagraphStyle("企画団体名", false);
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
