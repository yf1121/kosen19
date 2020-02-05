pageObj = app.activeDocument;  //アクティブなドキュメントを選択
txtObj = pageObj.textFrames.add();  //フレーム追加
txtObj.visibleBounds = ["4cm","5cm","8cm","10cm"];
textFile = new File("/C/others/xampp/htdocs/jitsui/panf-out.csv");
txtObj.place(textFile);