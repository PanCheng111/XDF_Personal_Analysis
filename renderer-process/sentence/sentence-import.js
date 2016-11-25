const ipc = require('electron').ipcRenderer
const storage = require('electron-json-storage');

const selectDirBtn = document.getElementById('select-sentence-import')

selectDirBtn.addEventListener('click', function (event) {
  ipc.send('open-sentence-import-dialog')
})

ipc.on('selected-sentence-import-directory', function (event, path) {
  document.getElementById('selected-sentence-import-file').innerHTML = `You selected: ${path}`
})

ipc.on('display-sentence-import-directory', function(event) {
    
    storage.get('sentence-data', function(err, data) {
        if (err) console.log("fetch sentence-data error!");
        else {
            var display = document.getElementById('displayed-sentence-import-file');
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
});

storage.get('sentence-data', function(err, data) {
    if (err) console.log("fetch sentence-data error!");
    else {
        if (!(data instanceof Array)) return;
        var display = document.getElementById('displayed-sentence-import-file');
        var content = "<p>您上次录入的数据如下：</p>"
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
