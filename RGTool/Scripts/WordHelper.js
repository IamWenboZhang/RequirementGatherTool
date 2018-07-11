function search() {
    Word.run(function (ctx) {

        // Queue a command to search the document for the string "Contoso".
        // Create a proxy search results collection object.
        var mustresults = ctx.document.body.search("must", { matchCase: true });
        var shouldresults = ctx.document.body.search("should", { matchCase: true });
        var mayresults = ctx.document.body.search("may", { matchCase: true });
        var alwaysresults = ctx.document.body.search("always", { matchCase: true });
        var mustnotsresults = ctx.document.body.search("MUST not", { matchCase: true });
        var shouldnotresults = ctx.document.body.search("SHOULD not", { matchCase: true });
        var allresults = [mustresults, shouldresults, mayresults, alwaysresults, mustnotsresults, shouldnotresults];
       
        ctx.load(mustresults);
        ctx.load(shouldresults);
        ctx.load(mayresults);
        ctx.load(alwaysresults);
        ctx.load(mustnotsresults);
        ctx.load(shouldnotresults);
        
        return ctx.sync().then(function () {
            for (var i = 0; i < allresults.length; i++) {
                for (var j = 0; j < allresults[i].items.length; j++) {
                    if (allresults[i].items.length > 0) {
                        allresults[i].items[j].font.color = "#FF0000"    // Change color to Red
                        allresults[i].items[j].font.highlightColor = "#FFFF00";
                        allresults[i].items[j].font.bold = true;
                        var cc = allresults[i].items[j].insertContentControl();
                        cc.tag = "WrongKeyWord";  // This value is used in another part of this sample.
                        cc.title = "Wrong RFC2119 keyword";
                    }

                }
            }
            showNotification("搜索成功", "小写的RFC2119关键词已被标识出来.");
        })
            .then(ctx.sync)
            .then(function () {
                
            })
            .catch(function (error) {
                //handleError(error);
                showNotification(error.code, error.message);
            })
    });
}


function getallsections() {
    Word.run(function (context) {
        var sections = context.document.sections;
        //sections.load("items");        
        context.load(sections, 'body/style');
        var contents = new Array();
        context.sync()
            .then(function () {
                if (sections.isNullObject) {
                    return;
                }
                for (var i = 0; i < sections.items.length; i++) {
                    sections.items[i].body.load("text");
                }
                return context.sync()
                    .then(function () {
                        for (var i = 0; i < sections.items.length; i++) {
                            contents[i] = sections.items[i].body.text;
                            console.log(contents[i]);
                        }
                    })
                    .catch(function (error) {

                    })
            })
            .catch(function (error) {
                console.log(error);
            });
    });
}

function gettables() {
    Word.run(function (context) {
        var tables = context.document.body.tables;
        context.load(tables, "items");
        context.sync()
            .then(function () {
                var tableitems = new Array();
                if (tables.isNullObject) {
                    return;
                }
                for (var i = 0; i < tables.items.length; i++) {
                    context.load(tables.items[i], "rows");
                    context.load(tables.items[i], "rowCount");
                    context.load(tables.items[i], "tables");
                }
                return context.sync()
                    .then(function () {
                        for (var i = 0; i < tables.items.length; i++) {
                            var rows = tables.items[i].rows;
                            rows.items[i].cells.items[i].value;
                            var rowcount = tables.items[i].rowCount;
                            var t = tables.items[i].tables;
                        }
                    })
                    .catch(function (error) {
                        console.log(error);
                    })
            })
    });
}

function list() {
    Word.run(function (context) {
        var lists = context.document.body.lists;
        lists.load("items");
        var listitems = new Array();
        context.sync()
            .then(function () {
                if (lists.items.isNullObject) {
                    return;
                }
                for (var i = 0; i < lists.items.length; i++) {
                    listitems[i] = lists.items[i];
                }
            })
            .catch(function (error) {
                console.log(error);
            })
    });
}

function getbodytext() {
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text property of the proxy body object.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the Word document body.
        body.insertText('This is text inserted after loading the body.text property',
            Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });
    })
}

function getcurrentselection() {
    Word.run(function (context) {
        var s = context.document.getSelection
    });
}

function getParagraph() {
    Word.run(function (context) {
        var p = context.document.body.paragraphs;
        //context.load(p);
        p.load("items");
        context.sync()
            .then(function () {
                if (p.isNullObject) {
                    return;
                }
                var ptext = new Array();
                for (var i = 0; i < p.items.length; i++) {
                    p.items[i].load("text");
                    ptext[i] = p.items[i].text;
                }
                return context.sync()
                    .then(function () {
                        for (var i = 0; i < ptext.length; i++) {
                            console.log(ptext[i]);
                        }
                    })
                    .catch(function (error) {
                        console.log(error);
                    })
            })
            .catch(function (error) {
                console.log(error);
            })
      
    });
}

function displaySelectedText() {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
        function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                showNotification('The selected text is:', '"' + result.value + '"');
            } else {
                showNotification('Error:', result.error.message);
            }
        });
}

//$$(Helper function for treating errors, $loc_script_taskpane_home_js_comment34$)$$
function errorHandler(error) {
    // $$(Always be sure to catch any accumulated errors that bubble up from the Word.run execution., $loc_script_taskpane_home_js_comment35$)$$
    showNotification("Error:", error);
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
}

// Helper function for displaying notifications
function showNotification(header, content) {
    $("#notification-header").text(header);
    $("#notification-body").text(content);
    messageBanner.showBanner();
    messageBanner.toggleExpansion();
}

function createexcel(filepath, ooxml) {
    var datas;
    console.log("进入函数");
    $.ajax({
        type: "post",
        url: "/Home/CreateExcel",
        data: {
            "filepath": filepath,
            "ooxml": ooxml
        },
        async: false,
        cache: false,
        datatype: 'json',
        success: function (data) {
            datas = data;
            console.log("成功");
        },
        error: function (data) {
            console.log("失败" + data);
            datas = data;
        }
    });
    console.log(datas);
}

function testtables() {
    Word.run(function (context) {
        var tables = context.document.body.tables;
        context.load(tables, "items");
        context.sync()
            .then(function () {
                var tableitems = tables.items;
                if (tableitems.isNullObject) {
                    console.log("Cannot find table of content.");
                }
                var tableranges = new Array();
                for (var i = 0; i < tableitems.length; i++) {
                    tableranges[i] = tableitems[i].getRange();
                    context.load(tableranges[i]);
                }
               
                context.sync()
                    .then(function () {
                        for (var i = 0; i < tableranges.length; i++) {
                            var text = tableranges[i].text
                            console.log(text);
                        }
                    })
               
            })
            .catch(function (error) {
                console.log(error);
            })
          });
}

function getcatlog() {
    Word.run(function (context) {
        var docproperties = context.document.properties;
        context.load(docproperties);
        context.sync()
            .then(function () {
                var category = docproperties.category;
                console.log(category);
            })
    });
}

function getmenu(callback) {
    var menu = new Array();
    var startindex = -1;
    var endindex = -1;
    Word.run(function (context) {
        var paragraphs = context.document.body.paragraphs;
        //sections.load("items");        
        context.load(paragraphs, 'body/style');
        var contents = new Array();
        context.sync()
            .then(function () {
                if (paragraphs.isNullObject) {
                    return;
                }
                for (var i = 0; i < paragraphs.items.length; i++) {
                    context.load(paragraphs.items[i], "text");
                }
                context.sync()
                    .then(function () {
                        for (var i = 0; i < paragraphs.items.length; i++) {
                            var text = paragraphs.items[i].text;
                            console.log(text);
                            if (text == "Table of Contents" && startindex == -1) {
                                startindex = i + 1;
                            }
                            if (text == "Introduction" && endindex == -1) {
                                endindex = i - 1;
                                break;
                            }
                        }
                        if (startindex != -1 && endindex != -1 && startindex != endindex) {
                            console.log("Start index is " + startindex + "     End index is " + endindex);
                            for (i = startindex; i <= endindex; i++) {
                                menu[i - startindex] = paragraphs.items[i].text;
                            }
                            //return menu;
                            callback(menu);
                        }
                    })
                    .catch(function (error) {
                        console.log(error);
                    })
            })
            .catch(function (error) {
                console.log(error);
            });
    });
}

function parsedocument(menu) {
    var sections = [];
    var sectionIds = [];
    var sectionNames = [];
    for (var i = 0; i < menu.length; i++) {
        if (menu[i].length > 0) {
            var tmp = menu[i].split("\t");
            sectionIds.push(tmp[0]);
            sectionNames.push(tmp[1]);
            sections[tmp[0]] = tmp[1]
        }
    }
};


