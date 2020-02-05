///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////CSVファイル選択//////
filename = File.openDialog("新規組版：一般企画のCSVファイルを選択してください...")  //ファイル選択
fileObj = new File(filename);
flag = fileObj.open("r");  //読み込みモードで開く
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////画像フォルダ選択//////
foldername = Folder.selectDialog("使用する画像が保存されているフォルダを選択してください...　新規組版：背景画像等の保存フォルダ選択");  //ファイル選択
folderObj = new File(foldername);
imgfold = folderObj.fsName;  //フォルダパスを取得
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////画像フォルダ選択//////
foldername = Folder.selectDialog("タグアイコンのaiファイルが保存されているフォルダを選択してください...　新規組版：背景画像等の保存フォルダ選択");  //ファイル選択
folderObj = new File(foldername);
tagfold = folderObj.fsName;  //フォルダパスを取得
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////ファイルとフォルダが選択されているか確認//////
if (flag == true && imgfold.substr (-4) != "null" && tagfold.substr (-4) != "null")  //CSVファイルが開けかつ指定フォルダの末尾4文字がnullでなければ、処理を始める
{
    docObj = app.activeDocument;  //アクティブなドキュメントを選択
    //////////////////////////////////カラースウォッチの設定////////////////////////////////////////////////////////////
    setColor = [0, 0, 0, 0];
    ColorWhite = docObj.colors.add({model:ColorModel.process,space:ColorSpace.cmyk,colorValue:setColor});
    setColor = [0, 0, 0, 100];
    ColorK = docObj.colors.add({model:ColorModel.process,space:ColorSpace.cmyk,colorValue:setColor});
    //////////////////////////////////段落スタイルの存在確認と新規作成/////////////////////////////////////
    var Style1 = docObj.paragraphStyles.item("企画名");
    if (Style1==null){
        Style1 = docObj.paragraphStyles.add({name:"企画名"});
        Style1.appliedFont = "源ノ角ゴシック";  //フォントファミリ設定
        Style1.fontStyle = "Bold";  //フォントスタイル設定
        Style1.fillColor = ColorWhite;  //文字色を紙色に設定
        Style1.pointSize = "11Q";  //プロパティ（フォントサイズ）設定
        Style1.leading = "13H";  //行送り
        Style1.justification = 1818915700;  //プロパティ（段落揃え）を均等最終行左揃えに設定
    }
    var Style2 = docObj.paragraphStyles.item("企画団体名");
    if (Style2==null){
        Style2 = docObj.paragraphStyles.add({name:"企画団体名"});
        Style2.basedOn = Style1  //基準スタイルの設定
        Style2.pointSize = "7Q";  //プロパティ（フォントサイズ）設定
        Style2.leading = "10H";  //プロパティ（行送り）設定
    }
    var Style3 = docObj.paragraphStyles.item("企画分類");
    if (Style3==null){
        Style3 = docObj.paragraphStyles.add({name:"企画分類"});
        Style3.basedOn = Style2  //基準スタイルの設定
        Style3.fontStyle = "Medium";  //フォントスタイル設定
        Style3.pointSize = "8Q";  //プロパティ（フォントサイズ）設定
        Style3.leading = "8H";  //プロパティ（行送り）設定
        Style3.fillColor = ColorK;  //文字色を黒色に設定
        Style3.justification = 1667591796;  //プロパティ（段落揃え）を中央揃えに設定
    }
    var Style4 = docObj.paragraphStyles.item("企画場所");
    if (Style4==null){
        Style4 = docObj.paragraphStyles.add({name:"企画場所"});
        Style4.basedOn = Style3  //基準スタイルの設定
        Style4.appliedFont = "FOT-筑紫A丸ゴシック Std";  //フォントファミリ設定
        Style4.fontStyle = "B";  //フォントスタイル設定
        Style4.horizontalScale = 70;  //プロパティ（文字水平比率）設定
        Style4.leading = "8.5H";  //プロパティ（行送り）設定
    }
    var Style5 = docObj.paragraphStyles.item("企画宣伝文");
    if (Style5==null){
        Style5 = docObj.paragraphStyles.add({name:"企画宣伝文"});
        Style5.basedOn = Style3  //基準スタイルの設定
        Style5.pointSize = "9Q";  //プロパティ（フォントサイズ）設定
        Style5.leading = "12H";  //プロパティ（行送り）設定
        Style5.justification = 1818915700;  //プロパティ（段落揃え）を均等最終行左揃えに設定
    }
    //////////////////////////////////段落スタイル設定ここまで////////////////////////////////////////////////////
    //////////////////////////////////ここから本編////////////////////////////////////////////////////////////////////////
    var separator = ",\""  //区切り文字の指定
    var pageno = 1
    var tentno = 1
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
                    ////////////////改ページ・改行処理//////////////////////////////////////////////////////
                    if (data[3].slice(0, -1) != pageno){  //改ページか確認
                        pageno = parseInt(data[3].slice(0, -1))  //ページ番号の更新
                        i = 0;
                        j = 0;  //改ページなら新規ページを追加
                        tentno = parseInt(data[4].slice(0, -3))  //テントアルファベットの更新
                        pageObj = docObj.pages.add(LocationOptions.AFTER, docObj.pages[pageno]);  //新規ページを追加し、アクティブなページに選択
                        pageObj = app.activeDocument.pages.add();  //新規ページを追加し、アクティブなページに選択
                    }else if (data[4].slice(0, -3).toString() != tentno){
                        tentno = data[4].slice(0, -3)  //テントアルファベットの更新
                        if (j == 0){
                            i = i + 0.2
                        }else{
                            i = i + 1.2
                            j = 0
                        }
                    }
                    //###企画宣伝背景###########################
                    imgpath = imgfold.replace(/\\/g, '/') + "/!!!backShinkansen.ai"
                    selFile = new File(imgpath);  //入れる画像ファイル指定
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
                    //###企画分類##############################
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
                    txtObj.parentStory.appliedParagraphStyle = "企画分類"  //段落スタイルの適用

                    //###企画名###############################
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
                    txtObj.parentStory.appliedParagraphStyle = "企画名"  //段落スタイルの適用
                    
                    //###企画団体名############################
                    txtObj = pageObj.textFrames.add();  //フレーム追加
                    txtObj.name = data[4] + data[1].slice(0, -1) + "-Name"  //レイヤー名の設定
                    x31 = x + 1.5  //左からの座標
                    y31 = y + 12.5  //上からの座標
                    x32 = x31 + 42;  //横42㎜
                    y32 = y31 + 1.75;  //縦1.75㎜
                    x31 = x31 + "mm";
                    y31 = y31 + "mm";
                    x32 = x32 + "mm";
                    y32 = y32 + "mm";
                    txtObj.visibleBounds = [y31,x31,y32,x32];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    txtObj.contents = data[11].slice(0, -1);  //テキストフレームに流し込む
                    txtObj.parentStory.appliedParagraphStyle = "企画団体名"  //段落スタイルの適用
                    //###企画場所##############################
                    txtObj = pageObj.textFrames.add();  //フレーム追加
                    txtObj.name = data[4] + data[1].slice(0, -1) + "-Place"  //レイヤー名の設定
                    x41 = x + 35.2  //左からの座標
                    y41 = y + 6.55  //上からの座標
                    x42 = x41 + 8.9;  //横8.9㎜
                    y42 = y41 + 3.2;  //縦3.2㎜
                    x41 = x41 + "mm";
                    y41 = y41 + "mm";
                    x42 = x42 + "mm";
                    y42 = y42 + "mm";
                    txtObj.visibleBounds = [y41,x41,y42,x42];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    txtObj.contents = data[5].slice(0, -1);  //テキストフレームに流し込む
                    txtObj.parentStory.appliedParagraphStyle = "企画場所"  //段落スタイルの適用
                    //###企画分類アイコン#######################
                    imgpath = imgfold.replace(/\\/g, '/') + "/" + data[34].slice(0, -1)  //パスの指定//企画分類アイコンデータの取得
                    selFile = new File(imgpath);  //入れる画像ファイル指定
                    imgObj = pageObj.textFrames.add();
                    imgObj.name = data[4] + data[1].slice(0, -1) + "-Icon";  //レイヤー名の設定
                    x51 = x + 2.8;  //左からの座標, 横の間隔
                    y51 = y +20.2;  //上からの座標
                    x52 = x51 + 8;  //横8㎜
                    y52 = y51 + 7.2;  //縦7.2㎜
                    x51 = x51 + "mm";
                    y51 = y51 + "mm";
                    x52 = x52 + "mm";
                    y52 = y52 + "mm";
                    imgObj.visibleBounds = [y51,x51,y52,x52];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    imgObj.contentType = ContentType.graphicType;  //フレームの種類をグラフィックに明示的に設定
                    imgObj.place(selFile);  //入れる画像
                    imgObj.fit(FitOptions.frameToContent);  //内容にフレームサイズを合わせる
                    //###企画宣伝文##############################
                    txtObj = pageObj.textFrames.add();  //フレーム追加
                    txtObj.name = data[4] + data[1].slice(0, -1) + "-PR"  //レイヤー名の設定
                    x41 = x + 20.5  //左からの座標
                    y41 = y + 19.7  //上からの座標
                    x42 = x41 + 22.5;  //横22.5㎜
                    y42 = y41 + 8.5;  //縦8.5㎜
                    x41 = x41 + "mm";
                    y41 = y41 + "mm";
                    x42 = x42 + "mm";
                    y42 = y42 + "mm";
                    txtObj.visibleBounds = [y41,x41,y42,x42];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                    if (data[15].slice(0, -1) == "〇"){  //もし学術枠なら、末尾にリンクをはる
                        data[19] = data[19].slice(0, -1) + "(→p." + data[31].slice(0, -1) + ")"  //リンクをはる
                        txtObj.contents = data[19];  //テキストフレームに流し込む
                    }else{
                        txtObj.contents = data[19].slice(0, -1);  //テキストフレームに流し込む
                    }
                    txtObj.parentStory.appliedParagraphStyle = "企画宣伝文"  //段落スタイルの適用
                    //###タグアイコン###############################
                    m = 0  //繰り返し変数の初期化
                    for (k=0; k<3; k++){
                        for (l=0; l<2; l++){
                            if (data[23+m].slice(0, -1) == ""){break;break;}
                            imgpath = tagfold.replace(/\\/g, '/') + "/" + data[23+m].slice(0, -1)  //パスの指定//タグアイコンデータの取得
                            selFile = new File(imgpath);  //入れる画像ファイル指定
                            imgObj = pageObj.textFrames.add();
                            imgObj.name = data[4] + data[1].slice(0, -1) + "-Tag";  //レイヤー名の設定
                            x51 = x + 4.45 * k + 11.375;  //左からの座標, 横の間隔
                            y51 = y + 4.45 * l + 15.12;  //上からの座標
                            x52 = x51 + 3.86;  //横3.86㎜
                            y52 = y51 + 3.86;  //縦3.86㎜
                            x51 = x51 + "mm";
                            y51 = y51 + "mm";
                            x52 = x52 + "mm";
                            y52 = y52 + "mm";
                            imgObj.visibleBounds = [y51,x51,y52,x52];  //新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
                            imgObj.contentType = ContentType.graphicType;  //フレームの種類をグラフィックに明示的に設定
                            imgObj.place(selFile);  //入れる画像
                            imgObj.fit(FitOptions.frameToContent);  //内容にフレームサイズを合わせる
                            m = m + 1  //繰り返し変数の更新
                        }
                    }
                    //###日付アイコン###############################
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
