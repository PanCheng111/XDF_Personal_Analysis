const ipc = require('electron').ipcMain
const dialog = require('electron').dialog
const storage = require('electron-json-storage')
const XLSX = require('xlsx')
var fs = require('fs')
var base64 = require('base64-js')
var Docxtemplater = require('docxtemplater')

function calc_score(name, score_item, exam_data) {
    var arrError = [];
    if (score_item['选择题错题号']) {
        score_item['选择题错题号'] += "";
        arrError = score_item['选择题错题号'].split(/,|，/).map(function(x) { return parseInt(x); });
    }
    var score = {};
    score.tot_score = 0;
    score.select_score = 0;
    score.select_max = 0;
    score.giant_score = 0;
    score.giant_max = 0;
    score.detail = [];
    score.detail_max = [];
    for (i = 0; i < exam_data.length; i++) {
        if (exam_data[i]['题目类型'] != '选择题') {
            var no = exam_data[i]['题号'];
            score.giant_max += exam_data[i]['分值'];
            score.detail_max[no] = exam_data[i]['分值'];
            continue;
        }
        var no = exam_data[i]['题号'];
        score.select_max += exam_data[i]['分值'];
        if (!arrError.includes(no)) {
            score.tot_score += exam_data[i]['分值'];
            score.select_score += exam_data[i]['分值'];
            score.detail[no] = exam_data[i]['分值'];
            score.detail_max[no] = exam_data[i]['分值'];
        }
        else {
            score.detail[no] = 0;
            score.detail_max[no] = exam_data[i]['分值'];
        }
    }
    for (key in score_item) {
        if (key != '姓名' && key != '选择题错题号') {
            score.tot_score += score_item[key];
            score.giant_score += score_item[key];
            var no = parseInt(key.split(/\(|（/)[0]);
            if (!score.detail[no]) score.detail[no] = 0;
            score.detail[no] += score_item[key];
        }
    }
    return score;
}

function calc_sentence(profiles, sentence_data) {
    for (var i = 0; i < profiles.length; i++) {
        var score = profiles[i].score_details;
        var select_ratio = score.select_score / score.select_max * 100;
        var giant_ratio = score.giant_score / score.giant_max * 100;
        var whole_ratio = score.tot_score / (score.select_max + score.giant_max) * 100;
        var select_sentence, giant_sentence, whole_sentence;
        for (var j = 0; j < sentence_data.length; j++) {
            var range = sentence_data[j]['正确率'].split(/~/);
            if (sentence_data[j]['类型'] == '选择题') {
                if (select_ratio > range[0] && select_ratio <= range[1]) 
                    select_sentence = sentence_data[j]['对应话术'];
            }
            if (sentence_data[j]['类型'] == '大题') {
                if (giant_ratio > range[0] && giant_ratio <= range[1]) 
                    giant_sentence = sentence_data[j]['对应话术'];
            }
            if (sentence_data[j]['类型'] == '总体得分') {
                if (whole_ratio > range[0] && whole_ratio <= range[1]) 
                    whole_sentence = sentence_data[j]['对应话术'];
            }
        }
        profiles[i].select_sentence = select_sentence;
        profiles[i].giant_sentence = giant_sentence;
        profiles[i].whole_sentence = whole_sentence;
    }
}

function generate_files(dir, profiles) {
    var content = fs.readFileSync(__dirname + "/../../doc_template/xx学员学情分析.docx", "binary");
    for (i = 0; i < profiles.length; i++) {
        var doc = new Docxtemplater(content);
        //set the templateVariables
        doc.setData({
            "name": profiles[i].name,
            "score": profiles[i].score,
            "rank": profiles[i].rank,
            "score_max": profiles[i].score_max,
            "score_drv": profiles[i].score_drv,
            "errors": profiles[i].errors,
            "select_sentence": profiles[i].select_sentence,
            "giant_sentence": profiles[i].giant_sentence,
            "whole_sentence": profiles[i].whole_sentence,
        });
        //apply them (replace all occurences of {first_name} by Hipp, ...)
        doc.render();
        var buf = doc.getZip().generate({type: "nodebuffer"});
        fs.writeFileSync(dir + "/" + profiles[i].name + "学员学情分析.docx", buf);
    }

}

function calc_practice(directory) {
    storage.get('profile-data', function(err, profile_data){
        if (err) {
            console.log(err);
        }
        else {
            storage.get('subject-data', function(err, subject_data) {
                if (err) {
                    dialog.showErrorDialog('错误', '请您导入题库对应信息！');
                }
                else {
                    storage.get('template-data', function(err, template_data) {
                        if (err) {
                            dialog.showOpenDialog('错误', '请您导入题库模板信息！');
                        }
                        else {
                            for (var i = 0; i < profile_data.length; i++) {
                                var err_list = [];
                                var errors = profile_data[i].errors;
                                for (var j = 0; j < errors.length; j++) {
                                    var err_no = errors[j].err_no;
                                    var err_analysis = errors[j].err_analysis;
                                    
                                    for (var k = 0; k < subject_data.length; k++) {
                                        if (subject_data[k]['知识点'] == err_analysis) {
                                            var list = subject_data[k]['题库中对应题号'].split(/,|，/).map(function(x) {return parseInt(x);});
                                            //console.log(list);
                                            err_list = err_list.concat(list);
                                            break;
                                        }
                                    }
                                }
                                var generate = {};
                                for (var j = 0; j < err_list.length; j++) {
                                    generate['no' + err_list[j].toString()] = {}
                                }
                                //console.log(err_list);
                                //console.log(generate);
                                var doc = new Docxtemplater(template_data);
                                doc.setData(generate);
                                doc.render();
                                var buf = doc.getZip().generate({type: "nodebuffer"});
                                fs.writeFileSync(directory + "/" + profile_data[i].name + "学员补充题.docx", buf);
                            }
                            dialog.showMessageBox({
                                type: 'info',
                                buttons: ['知道了'],
                                title: '保存成功',
                                message: '所有的考生补充习题已经输出到指定目录下。',
                                cancelId: 0,
                            });
                        }
                    })
                }
            });
        }
    });
}

ipc.on('open-practice-generate-dialog', function (event) {
  dialog.showOpenDialog({
    properties: ['openDirectory']
  }, function (files) {
    if (files) event.sender.send('selected-practice-generate-directory', files);
    else return;
    var dir = files[0];
    calc_practice(dir);
    // var workbook = XLSX.readFile(files[0]);
    // var sheet_name_list = workbook.SheetNames;
    // sheet_name_list.forEach(function(y) { /* iterate through sheets */
    //   var worksheet = workbook.Sheets[y];
    //   storage.set('exam-data', XLSX.utils.sheet_to_json(worksheet, {raw: true}), function(err) {
    //       if (err) console.log(err);
    //       else event.sender.send('display-analysis-generate-directory');
    //   });
    // });
  })
})

ipc.on('open-practice-generate-template-dialog', function (event) {
  dialog.showOpenDialog({
    properties: ['openFile']
  }, function (files) {
    if (files) event.sender.send('selected-practice-generate-template-directory', files);
    else return;
    var content = fs.readFileSync(files[0], "binary");
    //var data = new Buffer(content).toString('base64');
    //console.log(data);
    storage.set('template-data', content, function(err) {
        if (err) console.log(err);
    });
  })
})