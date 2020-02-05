var p=1;
pageObj = app.activeDocument.pages[p]
alert(pageObj.name);
var page=app.activeDocument.pageItemDefaults
alert(page.appliedTextObjectStyle.name);
page=app.activeDocument.pages.add(
    LocationOptions.AFTER,
    app.activeDocument.pages[p]
);
