function selectFile() {
    document.getElementById('file').click();
}

// 读取本地excel文件
function readWorkbookFromLocalFile(file, callback) {
    var reader = new FileReader();
    reader.onload = function (e) {
        var data = e.target.result;
        var workbook = XLSX.read(data, {type: 'binary'});
        if (callback) callback(workbook);
    };
    reader.readAsBinaryString(file);
}

// 从网络上读取某个excel文件，url必须同域，否则报错
function readWorkbookFromRemoteFile(url, callback) {
    var xhr = new XMLHttpRequest();
    xhr.open('get', url, true);
    xhr.responseType = 'arraybuffer';
    xhr.onload = function (e) {
        if (xhr.status == 200) {
            var data = new Uint8Array(xhr.response)
            var workbook = XLSX.read(data, {type: 'array'});
            if (callback) callback(workbook);
        }
    };
    xhr.send();
}

function readWorkbook(workbook) {
    var sheetNames = workbook.SheetNames;
    document.getElementById('result').innerHTML = sheetNames.length > 0
        ? generateSheetTop(sheetNames) + generateSheetEnd(workbook.Sheets)
        : "表格无数据";
}

function generateSheetTop(sheetNames) {
    let sheetTop = '<ul id="resultTab" class="nav nav-tabs">';
    for (const sheetName of sheetNames) {
        sheetTop += '<li><a href="#' + sheetName + '" data-toggle="tab">' + sheetName + '</a></li>';
    }
    return sheetTop + '</ul>';
}

function generateSheetEnd(sheets) {
    let sheetEnd = '<div id="resultTabContent" class="tab-content">';
    for (const sheetName in sheets) {
        sheetEnd += '<div class="tab-pane fade" id="' + sheetName + '">';
        var worksheet = sheets[sheetName];
        var csv = XLSX.utils.sheet_to_csv(worksheet);
        sheetEnd = sheetEnd + csv2table(csv);
        sheetEnd += '</div>';
    }
    return sheetEnd + '</div>';
}

// 将csv转换成表格
function csv2table(csv) {
    let html = '<table>';
    const rows = csv.split('\n');
    rows.forEach(function (row, idx) {
        var columns = row.split(',');
        html += '<tr>';
        for (const columnsKey in columns) {
            /**
             * 如果是第一列（英文列），则在文本框中追加词汇，否则没有点击按钮
             */
            const column = columns[columnsKey];
            const href = "javascript:addWord('" + window.btoa(window.encodeURIComponent(column)) + "')";
            columnsKey % 2 === 0
                ? html += "<td><a href=" + href + ">" + column + "</a></td>"
                : html += "<td>" + column + "</td>";
        }
        html += '</tr>';
    });
    html += '</table>';
    return html;
}

function loadRemoteFile(url) {
    readWorkbookFromRemoteFile(url, function (workbook) {
        readWorkbook(workbook);
    });
}

function addWord(word) {
    let oldText = document.getElementById("generateText").value;
    console.log(oldText);
    oldText = oldText.trim().charAt(oldText.length - 1) !== "" ? oldText + "," : oldText;
    document.getElementById("generateText").value = oldText + window.decodeURIComponent(window.atob(word));
}

function clearWord() {
    document.getElementById("generateText").value = "";
}

function copyWord() {
    var copyText = document.getElementById('generateText')
    copyText.select();
    navigator.clipboard.writeText(copyText.value);
    alert("已复制: " + copyText.value);
}


/**
 * 启动时默认加载
 */
$(function () {
    document.getElementById('file').addEventListener('change', function (e) {
        var files = e.target.files;
        if (files.length === 0) return;
        var f = files[0];
        if (!/\.xlsx?$/g.test(f.name)) {
            alert('仅支持读取xlsx格式！');
            return;
        }
        readWorkbookFromLocalFile(f, function (workbook) {
            readWorkbook(workbook);
        });
    });
    loadRemoteFile('./sample/word.xlsx');
});
