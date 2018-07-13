var sectionHelper = {
    sectionList: [],
    menustartindex: -1,
    menuendindex: -1,
    autogeather: function () {
        var callback = function (paragraphs) {
            if (paragraphs.length > 0) {
                var menu = new Array();
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
                    sectionHelper.sectionList = sectionHelper.parsemenu(menu);
                    var sections = sectionHelper.getsectioncontents(sectionHelper.sectionList, paragraphs)
                    return sections;
                }
            }
        }

        var p = getParagraphs(callback);
        return p;
    },

    parsemenu: function (menu) {
        var sectionItems = [];
        var sectionIds = [];
        var sectionNames = [];
        for (var i = 0; i < menu.length; i++) {
            if (menu[i].length > 0) {
                var tmp = menu[i].split("\t");
                var sectionItem = { ID: tmp[0], Name: tmp[1] };
                sectionItems.push(sectionItem);
            }
        }
        return sectionItems;
    },
    getsectioncontents: function (sectionItems, paragraphs) {
        var config = getconfig();
        var sections = [];
        if (paragraphs.length > 0 && sectionItems.length > 0) {
            for (var j = 0; j < sectionItems.length; j++) {
                var startindex = -1;
                var endindex = -1;
                for (var i = sectionHelper.menuendindex; i < paragraphs.length; i++) {
                    if (paragraphs[i] == sectionItems[j].Name && startindex == -1) {
                        startindex = i;
                    }
                    if (j == sectionItems.length - 1) {
                        endindex = paragraphs.length - 1;
                    }
                    else if (paragraphs[i] == sectionItems[j + 1].Name && endindex == -1) {
                        endindex = i;
                        break;
                    }
                }
                if (startindex != -1 && endindex != -1 && startindex != endindex) {
                    var contents = [];
                    for (var k = startindex + 1; k < endindex; k++) {
                        contents.push(paragraphs[k]);
                    }
                    var section = { ID: sectionItems[j].ID, Name: sectionItems[j].Name, Content: contents };
                    sections.push(section);
                }
            }
            return sections;
        }
    }
};