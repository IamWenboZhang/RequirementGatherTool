function getsentences(content) {
    var startindex = 0;
    var sentence = "";
    var sentences = [];
   
    for (var i = startindex; i < content.length; i++) {
        if (i == content.length - 1) {
            sentences.push(sentence);
            startindex = i + 1;
            sentence = "";
        }
        else {
            if ((content[i] == '.' || content[i] == ';' || content[i] == '!' || content[i] == '?') && content[i+1] == ' ') {
                sentences.push(sentence);
                startindex = i + 1;
                sentence = "";
            }
            else {
                sentence += content[i];
            }
        }
    }
    return sentences;
}

//function autogather() {
//    getconfig(function (data) {
//        var jsonobj = jQuery.parseJSON(data);
//        getParagraphs(sectionHelper.splitsectioncontent, jsonobj);
//    });   
//}
function autogather(data) {
    var jsonobj = jQuery.parseJSON(data);
    getParagraphs(sectionHelper.splitsectioncontent, jsonobj);
}
