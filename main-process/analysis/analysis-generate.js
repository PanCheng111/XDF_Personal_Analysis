const ipc = require('electron').ipcMain
const dialog = require('electron').dialog
const storage = require('electron-json-storage')
const XLSX = require('xlsx')
var fs = require('fs')
var Docxtemplater = require('docxtemplater')

function calc_score(name, score_item, exam_data) {
    var arrError = [];
    score_item['选择题错题号'] += "";
    console.log('name=', name, 'score=', score_item['选择题错题号']);
    if (score_item['选择题错题号']) {
        if (score_item['选择题错题号'].indexOf(',') == -1 && score_item['选择题错题号'].indexOf('，') == -1)
            arrError = [parseInt(score_item['选择题错题号'])];
        else arrError = score_item['选择题错题号'].split(/,|，/).map(function(x) { return parseInt(x); });
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
            var no = key;
            key += "";
            if (key.indexOf('(') != -1 || key.indexOf('（') != -1 ) no = parseInt(key.split(/\(|（/)[0]);
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
        var select_sentence, giant_sentence, whole_sentence, solve_sentence;
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
            if (sentence_data[j]['类型'] == '解决方案') {
                if (whole_ratio > range[0] && whole_ratio <= range[1]) 
                    solve_sentence = sentence_data[j]['对应话术'];
            }
        }
        profiles[i].select_sentence = select_sentence;
        profiles[i].giant_sentence = giant_sentence;
        profiles[i].whole_sentence = whole_sentence;
        profiles[i].solve_sentence = solve_sentence;
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
            "score_drv": profiles[i].score_drv.toFixed(1),
            "errors": profiles[i].errors,
            "select_sentence": profiles[i].select_sentence,
            "giant_sentence": profiles[i].giant_sentence,
            "whole_sentence": profiles[i].whole_sentence,
            "solve_sentence": profiles[i].solve_sentence,
        });
        //apply them (replace all occurences of {first_name} by Hipp, ...)
        doc.render();
        var buf = doc.getZip().generate({type: "nodebuffer"});
        fs.writeFileSync(dir + "/" + profiles[i].name + "学员学情分析.docx", buf);
    }

}

function calc_profile(directory) {
    storage.get('exam-data', function(err, exam_data){
        if (err) {
            dialog.showErrorBox('错误', '请您导入试题信息！');
        }
        else {
            storage.get('score-data', function(err, score_data) {
                if (err) {
                    dialog.showErrorBox('错误', '请您导入考生答题情况！');
                }
                else {
                    storage.get('sentence-data', function(err, sentence_data) {
                        if (err) {
                            dialog.showErrorBox('错误', '请您导入总结话术！');
                        }
                        else {
                            var profile = [];
                            for (var i = 0; i < score_data.length; i++) {
                                var name = score_data[i]['姓名'];
                                if (name == null) continue;
                                var score = calc_score(name, score_data[i], exam_data);
                                profile[i] = {};
                                profile[i].name = name;
                                profile[i].score = score.tot_score;
                                profile[i].score_details = score;
                                profile[i].errors = [];
                            }
                            var correct_ratio = [];
                            var err_message = [];
                            for (var i = 0; i < exam_data.length; i++) {
                                var no = exam_data[i]['题号'];
                                var sum = 0;
                                for (var j = 0; j < score_data.length; j++) {
                                    if (profile[j].score_details.detail[no] > 0.8 * profile[j].score_details.detail_max[no])
                                        sum ++;
                                }
                                correct_ratio[no] = sum / score_data.length;
                                err_message[no] = exam_data[i]['考察知识点'];
                            }
                            for (var i = 0; i < score_data.length; i++) {
                                var score = profile[i].score_details;
                                for (var j = 0; j < exam_data.length; j++) {
                                    var no = exam_data[j]['题号'];
                                    if (score.detail[no] <= 0.8 * score.detail_max[no]) {
                                        profile[i].errors.push({
                                            'err_no': no, 
                                            'err_correct_ratio': correct_ratio[no],
                                            'err_analysis': err_message[no],
                                        });
                                    }
                                }
                            }
                            profile.sort(function(a, b) {
                                if (a.score > b.score) return -1;
                                if (a.score < b.score) return 1;
                                return 0;
                            })
                            var score_avg = profile.reduce(function(pre, x) { return pre + x.score; }, 0) / profile.length;
                            for (var i = 0; i < profile.length; i++) {
                                if (i > 0 && profile[i].score == profile[i - 1].score)
                                    profile[i].rank = profile[i - 1].rank;
                                else profile[i].rank = i + 1;
                                profile[i].score_drv = profile[i].score - score_avg;
                                profile[i].score_max = profile[0].score;
                            }
                            calc_sentence(profile, sentence_data);
                            generate_files(directory, profile);
                            storage.set('profile-data', profile, function(err) {
                                if (err) console.log(err);
                            });
                            dialog.showMessageBox({
                                type: 'info',
                                buttons: ['知道了'],
                                title: '保存成功',
                                message: '所有的考生分析文档已经输出到指定目录下。',
                                cancelId: 0,
                            });
                        }
                    });
                }
            });
        }
    });
}

ipc.on('open-analysis-generate-dialog', function (event) {
  dialog.showOpenDialog({
    properties: ['openDirectory']
  }, function (files) {
    if (files) event.sender.send('selected-analysis-generate-directory', files);
    else return;
    var dir = files[0];
    calc_profile(dir);
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
