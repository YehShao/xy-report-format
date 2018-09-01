require('dotenv').load();
const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const xls = require('xlsx');
const wordextractor = require('word-extractor');
const { getReformingData, getDocBuffer, getDefaltDataWithName } = require('./Utils');
const NameListTable = require('./NameListTable.js');

const XLS_DIR = `${__dirname}/input/xls/NameList.xls`;
const INPUT_DIR = `${__dirname}/input/`;
const OUTPUT_DIR = `${__dirname}/output/`;
const TEMPLATE_FILE_DIR = `${__dirname}/template/template.docx`;
const OUTPUT_FILENAME = process.env.OUTPUT_FILENAME_PRIFIX + process.env.OUTPUT_FILENAME_SUFFIX + '.docx';
const TEMPLATE_CONTENTS = fs.readFileSync(TEMPLATE_FILE_DIR, 'binary');
// Read the excel file to get a name array
var exceNameList = new NameListTable(xls.readFile(XLS_DIR).Sheets['週六晚4A']);
var nameArray = exceNameList.parseTable();
// Read each input file
var dataArray = [];
var promiseArray = [];
fs.readdirSync(INPUT_DIR).forEach((fileName) => {
    // Filter out non-Word files 
    if (!fileName.match('.docx') && !fileName.match('.doc'))
        return;
    // Prepare Word file parsing tasks
    console.log(`Input file name : ${fileName}`);
    let filePath = INPUT_DIR + fileName;
    if (fileName.match('.docx')) {
        promiseArray.push(mammoth.extractRawText({path: filePath}).then((result) => {
            dataArray.push(getReformingData(result.value, fileName));       
        }));
    } else if (fileName.match('.doc')) {
        let extractor = new wordextractor();
        promiseArray.push(extractor.extract(filePath).then((result) => {
            dataArray.push(getReformingData(result.getBody(), fileName));
        }));
    }
});
const reflect = p => p.then(v => ({v, status: "fulfilled" }),
                            e => ({e, status: "rejected" }));
// Wait for all promises completed, then write the output file.
var nameListForWhoDidntSubmit = [];
var dataCannotBeParsed = [];
Promise.all(promiseArray.map(reflect)).then(function(values) {
    // Map and re-order dataArray with nameArray
    let sortedDataArray = [];
    nameArray.forEach(function(name) {
        var data = dataArray.find(function(x) { return x.name === name })
        if (data === undefined) {
            // Cannot find his/her data in the input files
            sortedDataArray.push(getDefaltDataWithName(name));
            nameListForWhoDidntSubmit.push(name);
            // [todo]: Log and generate name list for whose doc not in the output file.
        } else {
            sortedDataArray.push(data);
            dataArray.splice(dataArray.indexOf(data), 1);
        }
    })
    let buffer = getDocBuffer(TEMPLATE_CONTENTS, {loop: sortedDataArray});
    fs.writeFileSync(path.resolve(OUTPUT_DIR, OUTPUT_FILENAME), buffer);
    dataCannotBeParsed = dataArray.map(function(x) { return x.fileName });
    // Generate report
    var file = fs.createWriteStream(path.resolve(OUTPUT_DIR, 'Report.txt'));
    file.write('＊＊＊缺交名單(或文章無法辨識作者)：\n');
    nameListForWhoDidntSubmit.forEach(function(x) { file.write(x + '\n') });
    file.write('＊＊＊以下檔案無法辨識作者(或未被列在出席名單)):\n');
    dataCannotBeParsed.forEach(function(x) { file.write(x + '\n') });
    file.end();
} )

