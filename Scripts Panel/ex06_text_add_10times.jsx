selFile = new File("/C/Users/Yudai/OneDrive - 筑波大学/学園祭実行委員会/広報宣伝局2019/パンフレット/design資料/企画紹介エリア別（新幹線）.ai");  //入れる画像ファイル指定
docObj = app.activeDocument;  //アクティブなドキュメントを選択
for (i=0; i<6; i++)
{
    for (j=0; j<3; j++)
    {
    imgObj = docObj.textFrames.add();
    x1 = j * 47 + 3;  //左からの座標, 横の間隔
    y1 = i * 31 +3;  //上からの座標
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
    }
}