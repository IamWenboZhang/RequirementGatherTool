function insertExcel(jsonStr,configJson) {
    jsonStr = escape(jsonStr);
    configdata = escape(configJson);
    $.ajax({
        type: "post",
        url: "/Excel/CreateExcel",
        data: {
            "templateName": "MS-XXXX_RequirementSpecification.xlsx",
            "JsonContent": jsonStr,
            "Configdata": configJson,
        },
        async: false,
        cache: false,
        datatype: "json",
        success: function (data) {
            console.log(data);
        },
        error: function (error) {
            console.log(error);
        }
    });
}
