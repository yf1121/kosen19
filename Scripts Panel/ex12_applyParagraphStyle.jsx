//
// InDesign ExtendScript 
// 段落スタイルを新規作成して適用する.
//


// 1) ドキュメント作成
var myDoc=app.documents.add();
var myPage=myDoc.pages.item(0);


// 2) テキストフレームを作成して内容を配置
var tf=createTextFrame(myDoc,myPage);
tf.contents = "吾輩（わがはい）は猫である。\r名前はまだ無い。";


// 3) 最初の段落のみスタイルを作成しそれを適用
var myStyle = myDoc.paragraphStyles.add({name:"myStyle"});  //新規スタイル作成
myStyle.pointSize = "14pt";
tf.paragraphs[0].applyParagraphStyle(myStyle, true);


// 4) 確認
alert( tf.paragraphs[0].contents )
alert( tf.paragraphs[1].contents )



// 便利メソッド

function createTextFrame(myDoc,myPage){
	var myTextFrame=myPage.textFrames.add();
	var myBounds = myGetBounds(myDoc,myPage);
	myTextFrame.geometricBounds=myBounds;

	return myTextFrame;
}

function myGetBounds(myDoc, myPage){
	with(myDoc.documentPreferences){
		var myPageHeight = pageHeight;
		var myPageWidth = pageWidth;
	}
	with(myPage.marginPreferences){
		var myTop = top;
		var myLeft = left;
		var myRight = right;
		var myBottom = bottom;
	}
	myRight = myPageWidth - myRight;
	myBottom = myPageHeight- myBottom;
	return [myTop, myLeft, myBottom, myRight];
}