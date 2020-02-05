foldername = Folder.selectDialog("使用する画像が保存されているフォルダを選択してください");
folderObj = new File(foldername);
fsname = folderObj.fsName;
alert(fsname);