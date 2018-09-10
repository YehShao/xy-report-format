require('dotenv').load();
let fs = require('fs');
let path = require('path');
let mammoth = require('mammoth');
let xls = require('xlsx');
let wordextractor = require('word-extractor');
let { getReformingData, getDocBuffer, getDefaultDataWithName } = require('./Utils');
let NameListTable = require('./NameListTable.js');

const XLS_DIR = `${__dirname}/input/xls/NameList.xls`;
const INPUT_DIR = `${__dirname}/input/`;
const OUTPUT_DIR = `${__dirname}/output/`;
const TEMPLATE_FILE_DIR = `${__dirname}/template/template_multi.docx`;
const OUTPUT_FILENAME = process.env.OUTPUT_FILENAME;
const TEMPLATE_CONTENTS = fs.readFileSync(TEMPLATE_FILE_DIR, 'binary');
// Read the excel file to get a name array
let excelNameList = new NameListTable(xls.readFile(XLS_DIR).Sheets['週六晚4A']);
let nameArray = excelNameList.parseTable();
// Read each input file
let promiseArray = [];
fs.readdirSync(INPUT_DIR).forEach((fileName) => {
    // Filter out non-Word files 
    if (!fileName.match('.docx') && !fileName.match('.doc'))
        return;
    // Prepare Word file parsing tasks
    console.log(`Input file name : ${fileName}`);
    let filePath = INPUT_DIR + fileName;
    if (fileName.match('.docx')) {
        promiseArray.push(new Promise((resolve, reject) => {
            return resolve(mammoth.extractRawText({ path: filePath }).then(result => {
                return getReformingData(result.value, fileName);
            }));
        }));
    } else if (fileName.match('.doc')) {
        let extractor = new wordextractor();
        promiseArray.push(new Promise((resolve, reject) => {
            return resolve(extractor.extract(filePath).then((result) => {
                return getReformingData(result.getBody(), fileName);
            }));
        }));
    }
});
const reflect = p => p.then(v => ({v, status: "fulfilled" }),
                            e => ({e, status: "rejected" }));

// Wait for all promises completed, then write the output file.
Promise.all(promiseArray.map(reflect)).then((completedPromiseArray) => {
    dataArray = completedPromiseArray.filter(obj => (obj.v)).map(obj => obj.v);
    // Map and re-order dataArray with nameArray
    let nameListForWhoDidntSubmit = [];
    let dataCannotBeParsed = [];
    let sortedDataArray = [];
    nameArray.forEach((name) => {
        let data = dataArray.find((obj) => {
            return (obj.name === name);
        });
        if (!data) {
            // Cannot find his/her data in the input files
            sortedDataArray.push(getDefaultDataWithName(name));
            nameListForWhoDidntSubmit.push(name);
            // [todo]: Log and generate name list for whose doc not in the output file.
        } else {
            sortedDataArray.push(data);
            //dataArray.shift();
            dataArray.splice(dataArray.indexOf(data), 1);
        }
    });
    let buffer = getDocBuffer(TEMPLATE_CONTENTS, { loop: sortedDataArray });
    fs.writeFileSync(path.resolve(OUTPUT_DIR, OUTPUT_FILENAME), buffer);
    dataCannotBeParsed = dataArray.map((obj) => {
        return obj.fileName;
    });
    // Generate report
    let date = new Date();
    let timeString = new Date().toLocaleString().substr(0, 10).replace(/-/g, '');
    let file = fs.createWriteStream(path.resolve(OUTPUT_DIR, `Report_${timeString}.txt`));
    file.write('＊＊＊缺交名單(或文章無法辨識作者)：\n');
    nameListForWhoDidntSubmit.forEach((obj) => {
        file.write(`${obj}\n`);
    });
    file.write('\n＊＊＊以下檔案無法辨識作者(或未被列在出席名單)：\n');
    dataCannotBeParsed.forEach((obj) => {
        file.write(`${obj}\n`);
    });
    file.end();
}).catch((err) => {
    console.log(err);
});
