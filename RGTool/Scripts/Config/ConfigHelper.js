function setconfig() {
    $.ajax({
        type: "post",
        url: "/Config/SetConfig",
        data: {
            "shortname": document.getElementById("ip_shortname").value,
            "version": document.getElementById("ip_version").value,
            "type": document.getElementById("select_tdtype").value,
            "startsection": document.getElementById("ip_startsection").value,
            "endsection": document.getElementById("ip_endsection").value
        },
        async: true,
        cache: false,
        datatype: "json",
        success: function (data) {
            console.log(data);
            autogather(data);
        },
        error: function (error){
            console.log(error);
        }
    });
}

function getconfig(callback) {
    $.ajax({
        type: "get",
        url: "Config/GetConfig",
        async: false,
        cache: false,
        success: function (data) {
            callback(data);
        },
        error: function (error) {
            console.log("getconfig"+error);
        }
    });
}

function parseconfig(configdata) {

}