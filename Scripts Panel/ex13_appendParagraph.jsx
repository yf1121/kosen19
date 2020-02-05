// 段落構築用テキストフレームを生成
function createBufferedTextFrame(){
	var mytf  = app.activeDocument.textFrames.add();
	return mytf;
}
// マージ
function mergeTextFrame( tf0 , tf1 ){
	tf0.nextTextFrame = tf1;
	tf1.remove();
}

// 改行を入れる
function appendBr( tf ){
	var mytf = createBufferedTextFrame()
	mytf.contents = "\r";
	mergeTextFrame( tf,mytf );
}

function applyPStyle( tf,pstylename ){
	var myPStyle=null;
	try{ 
		var myDoc = app.activeDocument;
		myPStyle=myDoc.paragraphStyles.item( pstylename );
		myPStyle.name;
	}
	catch(myError){
		var myDoc = app.activeDocument;
		myPStyle = myDoc.paragraphStyles.add({name:pstylename});
	}

	tf.paragraphs.everyItem().applyParagraphStyle(myPStyle, true);

	/*
	for( var i=0; i<tf.paragraphs.length; i++){
		tf.paragraphs.item(i).applyParagraphStyle(myPStyle, true);
		alert(tf.paragraphs.item(i));
	}
	*/
}


function appendText(tf, text, pstylename){
	var btf          = createBufferedTextFrame();
	btf.contents = text;
	appendBr( btf );
	applyPStyle(btf, pstylename);
	mergeTextFrame( tf, btf );
}


// ここより上が汎用の関数



// 以下は、アプリケーションコード


function createTextFrame(doc,page){
	var tf             = page.textFrames.add();
	tf.geometricBounds = ["20mm", "20mm", "40mm", "80mm"];
	return tf;
}


var myDoc  = app.documents.add();
//var myDoc  = app.activeDocument;
var myPage = myDoc.pages.item(0);
var tf     = createTextFrame(myDoc,myPage);

// 1)
appendText(tf, "Hello World.","h1");

// 2)
appendText(tf, "nin-nin","body");