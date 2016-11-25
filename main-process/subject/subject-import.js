const ipc = require('electron').ipcMain
const dialog = require('electron').dialog
const XLSX = require('xlsx')
const storage = require('electron-json-storage')

ipc.on('open-subject-import-dialog', function (event) {
  dialog.showOpenDialog({
    properties: ['openFile']
  }, function (files) {
    if (files) event.sender.send('selected-subject-import-directory', files);
    else return;
    var workbook = XLSX.readFile(files[0]);
    var sheet_name_list = workbook.SheetNames;
    sheet_name_list.forEach(function(y) { /* iterate through sheets */
      var worksheet = workbook.Sheets[y];
      storage.set('subject-data', XLSX.utils.sheet_to_json(worksheet, {raw: true}), function(err) {
          if (err) console.log(err);
          else event.sender.send('display-subject-import-directory');
      });
    });
  })
})
