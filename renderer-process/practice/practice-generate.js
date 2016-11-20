const ipc = require('electron').ipcRenderer
const storage = require('electron-json-storage')

const selectDirBtn = document.getElementById('select-practice-generate')
const selectFileBtn = document.getElementById('select-practice-generate-template')

selectDirBtn.addEventListener('click', function (event) {
  ipc.send('open-practice-generate-dialog')
})

ipc.on('selected-practice-generate-directory', function (event, path) {
  document.getElementById('selected-practice-generate-file').innerHTML = `文件保存路径: ${path}`
})

selectFileBtn.addEventListener('click', function (event) {
  ipc.send('open-practice-generate-template-dialog')
})

ipc.on('selected-practice-generate-template-directory', function (event, path) {
  document.getElementById('selected-practice-generate-template-file').innerHTML = `导入文件: ${path}`
})
