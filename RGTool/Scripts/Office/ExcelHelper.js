function insertExcel() {
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