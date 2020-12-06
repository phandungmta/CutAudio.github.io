function handleFile(e) {
    //Get the files from Upload control
    var files = e.target.files;
    var i, f, j;


    //Loop through files

    for (j = 0, f = files[j]; j != files.length; ++j) {
        var reader = new FileReader();
        var name = f.name;
        reader.onload = function (e) {
            var data = e.target.result;
            console.log(data)


            var workbook = XLSX.read(data, { type: 'binary' });

            var sheet_name_list = workbook.SheetNames;
            sheet_name_list.forEach(function (y) { /* iterate through sheets */
                //Convert the cell value to Json
                var roa = XLSX.utils.sheet_to_json(workbook.Sheets[y]);
                if (roa.length > 0) {
                    result = roa;
                }
            });

            if (result.length > 0) {
                var f = Number(result[0]['id'])
                document.getElementById("numberStart").value = f + ""
                document.getElementById("sub").innerHTML = result[0]['text']
            }
            var Table = document.getElementById("bodyTable");
            Table.innerHTML = ""


            //Get the first column first cell value
            for (i = 0; i < result.length; i++) {


                n = Number(result[i]['id'])

                var x =
                    '<tr>' +
                    '<td id="i' + n + '" >' + result[i]['id'] + '</td>' +
                    '<td id="t' + n + '">' + result[i]['text'] + '</td>' +
                    '<td id="s' + n + '">' + '</td>' +
                    '<td id="e' + n + '">' + '</td>' +
                    '<td> <button id="' + n + '" onclick="listenTest(this.id)">nghe thử </button>'
                '</tr>'

                $('#bodyTable').append(x);

            }

        };
        reader.readAsArrayBuffer(f);
    }
}
function makeTable(array) {
    var table = document.createElement('table');
    for (var i = 0; i < array.length; i++) {
        var row = document.createElement('tr');
        for (var j = 0; j < array[i].length; j++) {
            var cell = document.createElement('td');
            cell.textContent = array[i][j];
            row.appendChild(cell);
        }
        table.appendChild(row);
    }
    return table;
}

//Change event to dropdownlist
$(document).ready(function () {
    $('#files').change(handleFile);
});



var url = localStorage.getItem("excel");
var oReq = new XMLHttpRequest();
oReq.open("GET", url, true);
oReq.responseType = "arraybuffer";

oReq.onload = function (e) {
    var arraybuffer = oReq.response;

    /* convert data to binary string */
    var data = new Uint8Array(arraybuffer);
    var arr = new Array();
    for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
    var bstr = arr.join("");

    /* Call XLSX */
    var workbook = XLSX.read(bstr, { type: "binary" });




    var sheet_name_list = workbook.SheetNames;
    sheet_name_list.forEach(function (y) { /* iterate through sheets */
        //Convert the cell value to Json
        var roa = XLSX.utils.sheet_to_json(workbook.Sheets[y]);
        if (roa.length > 0) {
            result = roa;
        }
    });

    if (result.length > 0) {
        var f = Number(result[0]['id'])

        document.getElementById("numberStart").value = Number(f) + ""
        document.getElementById("sub").innerHTML = result[0]['text']
    }
    var Table = document.getElementById("bodyTable");
    Table.innerHTML = ""


    //Get the first column first cell value
    for (i = 0; i < result.length; i++) {



        n = Number(result[i]['id'])

        var x =
            '<tr>' +
            '<td id="i' + n + '" >' + result[i]['id'] + '</td>' +
            '<td id="t' + n + '">' + result[i]['text'] + '</td>' +
            '<td id="s' + n + '">' + '</td>' +
            '<td id="e' + n + '">' + '</td>' +
            '<td> <button id="' + n + '" onclick="listenTest(this.id)">nghe thử </button>'
        '</tr>'

        $('#bodyTable').append(x);

    }

}

oReq.send();