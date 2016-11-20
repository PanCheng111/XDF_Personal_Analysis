const ipc = require('electron').ipcRenderer
const storage = require('electron-json-storage');

const selectDirBtn = document.getElementById('select-score-import')

selectDirBtn.addEventListener('click', function (event) {
  ipc.send('open-score-import-dialog')
})

ipc.on('selected-score-import-directory', function (event, path) {
  document.getElementById('selected-score-import-file').innerHTML = `You selected: ${path}`
})

ipc.on('display-score-import-directory', function(event) {
    
    storage.get('score-data', function(err, data) {
        if (err) console.log("fetch score-data error!");
        else {
            var display = document.getElementById('displayed-score-import-file');
            var content = "<table class='table'>";
            var head = data[0];
            content += "<tr>";
            for (let i in head) {
                content += "<td>" + i + "</td>";
            }
            content += "</tr>";
            data.forEach(function(element) {
                content += "<tr>";
                for (let i in element) content += "<td>" + element[i] + "</td>";
                content += "</tr>";
            });
            content += "</table>";
            display.innerHTML = content;
        }
    });
})

storage.get('score-data', function(err, data) {
    if (err) console.log("fetch score-data error!");
    else {
        var display = document.getElementById('displayed-score-import-file');
        var content = "<p>您上次导入的数据如下：</p>";
        content += "<table class='table'>";
        var head = data[0];
        content += "<tr>";
        for (let i in head) {
            content += "<td>" + i + "</td>";
        }
        content += "</tr>";
        data.forEach(function(element) {
            content += "<tr>";
            for (let i in element) content += "<td>" + element[i] + "</td>";
            content += "</tr>";
        });
        content += "</table>";
        display.innerHTML = content;
    }
});
