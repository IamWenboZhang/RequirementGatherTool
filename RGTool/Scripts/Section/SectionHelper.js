var sectionHelper = {
    menuList: [],
    menustartindex: -1,
    menuendindex: -1,
    //getmenu: function () {
    //    var callback = function (paragraphs) {
    //        if (paragraphs.length > 0) {
    //            var menu = new Array();
    //            for (var i = 0; i < paragraphs.length; i++) {
    //                if (paragraphs[i] == "Table of Contents" && sectionHelper.menustartindex == -1) {
    //                    sectionHelper.menustartindex = i + 1;
    //                }
    //                if (paragraphs[i] == "Introduction" && sectionHelper.menuendindex == -1) {
    //                    sectionHelper.menuendindex = i - 1;
    //                    break;
    //                }
    //            }
    //            if (sectionHelper.menustartindex != -1 && sectionHelper.menuendindex != -1 && sectionHelper.menustartindex != sectionHelper.menuendindex) {
    //                for (i = sectionHelper.menustartindex; i <= sectionHelper.menuendindex; i++) {
    //                    menu[i - sectionHelper.menustartindex] = paragraphs[i];
    //                }
    //                sectionHelper.menuList = sectionHelper.parsemenu(menu);
    //                var sections = sectionHelper.getsectioncontents(sectionHelper.menuList, paragraphs)
    //                return sections;
    //            }
    //        }
    //    }
    //    var p = getParagraphs(callback);
    //    return p;
    //},
    splitsectioncontent: function (paragraphs, configdata) {
        var menu = sectionHelper.getmenu(paragraphs);
        var result = [];
        var sectiontitles = sectionHelper.parsemenu(menu);
        var sectionitems = sectionHelper.getsectioncontents(sectiontitles, paragraphs, configdata.StartSection, configdata.EndSection);
        var index = 0; 
        for (var i = 0; i < sectionitems.length; i++) {
            var sentences = getsentences(sectionitems[i].Content);
            for (var j = 0; j < sentences.length; j++) {
                //result.push("[in " + sectionitems[i].ID + " " + sectionitems[i].Name + "] " + sentences[j]);
                index++;
                var _content = "[in " + sectionitems[i].Name + "] " + sentences[j];
                var ExcelItem = { "ID": index, "RequirementID": "R" + index, "SectionID": sectionitems[i].ID, "Itemcontent": _content, "InformaltiveorNormaltive": "Normaltive", "ServerorClient": "Server" };
                result.push(ExcelItem);
            }
        }
        var jsonstr = JSON.stringify(result);
        var configJson = JSON.stringify(configdata);
        insertExcel(jsonstr, configJson);
    },

    getmenu: function (paragraphs) {
        if (paragraphs.length > 0) {
            var menu = [];
            for (var i = 0; i < paragraphs.length; i++) {
                if (paragraphs[i] == "Table of Contents" && sectionHelper.menustartindex == -1) {
                    sectionHelper.menustartindex = i + 1;
                }
                if (paragraphs[i] == "Introduction" && sectionHelper.menuendindex == -1) {
                    sectionHelper.menuendindex = i - 1;
                    break;
                }
            }
            if (sectionHelper.menustartindex != -1 && sectionHelper.menuendindex != -1 && sectionHelper.menustartindex != sectionHelper.menuendindex) {
                for (i = sectionHelper.menustartindex; i <= sectionHelper.menuendindex; i++) {
                    menu[i - sectionHelper.menustartindex] = paragraphs[i];
                }
                return menu;
            }
        }
    },

    parsemenu: function (menu) {
        var sectiontitles = [];
        for (var i = 0; i < menu.length; i++) {
            if (menu[i].length > 0) {
                var tmp = menu[i].split("\t");
                var sectionItem = { ID: tmp[0], Name: tmp[1] };
                sectiontitles.push(sectionItem);
            }
        }
        return sectiontitles;
    },

    getsectioncontents: function (sectiontitles, paragraphs, startsectionID, endsectionID) {
        var config = getconfig();
        var result = [];
        var needThisSection = false;
        var loopstartindex = sectionHelper.menuendindex;
        if (paragraphs.length > 0 && sectiontitles.length > 0) {
            for (var j = 0; j < sectiontitles.length; j++) {
                var startindex = -1;
                var endindex = -1;
                for (var i = loopstartindex; i < paragraphs.length; i++) {
                    if (paragraphs[i] == "Complex Types") {
                        console.log(i);
                    }
                    if (paragraphs[i] == sectiontitles[j].Name && startindex == -1) {
                        startindex = i;
                    }
                    if (j == sectiontitles.length - 1) {
                        endindex = paragraphs.length - 1;
                    }
                    else if (paragraphs[i] == sectiontitles[j + 1].Name && endindex == -1) {
                        endindex = i;
                        loopstartindex = endindex;
                        break;
                    }
                }
                if (startindex != -1 && endindex != -1 && startindex != endindex) {
                    if (sectiontitles[j].ID == startsectionID && !needThisSection) {
                        needThisSection = true;
                    }
                    if (needThisSection) {
                        var contentstr = "";
                        var contents = [];
                        for (var k = startindex + 1; k < endindex; k++) {
                            contentstr += paragraphs[k];
                            contents.push(paragraphs[k]);
                        }
                        //var section = { ID: sectionItems[j].ID, Name: sectionItems[j].Name, Content: contents };
                        var sectionitem = { ID: sectiontitles[j].ID, Name: sectiontitles[j].Name, Content: contentstr };
                        result.push(sectionitem);
                        //sectionHelper.sections.push(section);
                        if (sectiontitles[j].ID == endsectionID) {
                            needThisSection = false;
                            break;
                        }
                    }
                }
            }
            //return sectionHelper.sections;
            return result;
        }
    }
};