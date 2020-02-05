selFile = new File("/C/others/xampp/htdocs/dashboard/images/apple-logo.png");  //入れる画像ファイル指定
docObj = app.activeDocument;  //アクティブなドキュメントを選択
txtObj = docObj.textFrames.add();  //フレーム追加
//新規フレーム[x,y,z,w] => 上からx、左からyに、縦z-x、横w-yの長方形を作る
txtObj.visibleBounds = ["2cm","2cm","5cm","5cm"];
txtObj.contentType = ContentType.graphicType;  //フレームの種類をグラフィックに明示的に設定
txtObj.place(selFile);  //入れる画像
txtObj.fit(FitOptions.proportionally);  //フレームに内容を収める