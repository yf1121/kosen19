pageObj = app.activeDocument  //アクティブなドキュメントを選択
txtObj = pageObj.textFrames.add();  //ページ追加
//新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
txtObj.visibleBounds = ["2cm","2cm","3cm","20cm"];
txtObj.contentType = ContentType.textType;  //フレームの種類をテキストに明示的に設定
txtObj.contents = "InDesign JavaScript Sample";  //入れる文字列
txtObj.characters[0].pointSize = "24Q";  //先頭文字のフォントサイズを指定