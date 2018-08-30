require('dotenv').load();
const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const xls = require('xlsx');
const wordextractor = require('word-extractor');
const { getReformingData, getDocBuffer } = require('./Utils');
const NameListTable = require('./NameListTable.js');

const XLS_DIR = `${__dirname}/input/xls/NameList.xls`;
const INPUT_DIR = `${__dirname}/input/`;
const OUTPUT_DIR = `${__dirname}/output/`;
const TEMPLATE_FILE_DIR = `${__dirname}/template/template.docx`;
const OUTPUT_FILENAME = process.env.OUTPUT_FILENAME_PRIFIX + process.env.OUTPUT_FILENAME_SUFFIX + 'docx';
const TEMPLATE_CONTENTS = fs.readFileSync(TEMPLATE_FILE_DIR, 'binary');
// Read the excel file to get a name array
var exceNameList = new NameListTable(xls.readFile(XLS_DIR).Sheets['週六晚4A']);
var nameArray = exceNameList.parseTable();
// Read each input file
var rawTextArray = [];
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
            rawTextArray.push(getReformingData(result.value));       
        }));
    } else if (fileName.match('.doc')) {
        let extractor = new wordextractor();
        promiseArray.push(extractor.extract(filePath).then((result) => {
            rawTextArray.push(getReformingData(result.getBody()));
        }));
    }
});
const reflect = p => p.then(v => ({v, status: "fulfilled" }),
                            e => ({e, status: "rejected" }));
// Wait for all promises completed, then write the output file.
Promise.all(promiseArray.map(reflect)).then(function(values) {
    let data = {loop: rawTextArray};
    let buffer = getDocBuffer(TEMPLATE_CONTENTS, data);
    fs.writeFileSync(path.resolve(OUTPUT_DIR, OUTPUT_FILENAME), buffer);
} )
