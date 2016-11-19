const ipc = require('electron').ipcRenderer
const storage = require('electron-json-storage')

const selectDirBtn = document.getElementById('select-analysis-generate')

selectDirBtn.addEventListener('click', function (event) {
  ipc.send('open-analysis-generate-dialog')
})

ipc.on('selected-analysis-generate-directory', function (event, path) {
  document.getElementById('selected-analysis-generate-file').innerHTML = `文件保存路径: ${path}`
})
