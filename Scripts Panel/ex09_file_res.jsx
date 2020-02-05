pageObj = app.activeDocument;  //アクティブなドキュメントを選択
txtObj = pageObj.textFrames.add();  //フレーム追加
txtObj.visibleBounds = ["4cm","5cm","8cm","10cm"];
fileObj = new File("/C/others/xampp/htdocs/jitsui/panf-out.csv");
flag = fileObj.open("r");
if (flag == true)
{
text = fileObj.readch();
alert(text.charCodeAt(0));
fileObj.close();
}else{
alert("ファイルが開けませんでした");
}