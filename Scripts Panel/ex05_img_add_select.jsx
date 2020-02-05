selFile = File.openDialog("画像ファイルを選択してください");
docObj = app.activeDocument;  //アクティブなドキュメントを選択
imgObj = docObj.textFrames.add();  //フレーム追加
imgObj.contentType = ContentType.graphicType;  //フレームの種類をグラフィックに明示的に設定
imgObj.place(selFile);  //入れる画像
imgObj.fit(FitOptions.frameToContent);  //内容にフレームサイズを合わせる