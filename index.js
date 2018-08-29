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
const OUTPUT_FILENAME_PRIFIX = process.env.OUTPUT_FILENAME_PRIFIX;
const OUTPUT_FILENAME_SUFFIX = process.env.OUTPUT_FILENAME_SUFFIX;
const TEMPLATE_CONTENTS = fs.readFileSync(TEMPLATE_FILE_DIR, 'binary');
//Read the excel file to get a name array
var exceNameList = new NameListTable(xls.readFile(XLS_DIR).Sheets['週六晚4A']);
var nameArray = exceNameList.parseTable();
// Read each input file
var dataArray = [];
fs.readdirSync(INPUT_DIR).forEach((fileName) => {
    if (!fileName.match('.docx') && !fileName.match('.doc'))
        return;
    console.log(`Input file name : ${fileName}`);
    let test = INPUT_DIR + fileName;
    let rawText;
    if (fileName.match('.docx')) {
        mammoth.extractRawText({path: test}).then((result) => {
            rawText = result.value; 
            dataArray.push(getReformingData(rawText));         
        });
    } else if (fileName.match('.doc')) {
        let extractor = new wordextractor();
        extractor.extract(INPUT_DIR + fileName).then((result) => {
            rawText = result.getBody();
            dataArray.push(getReformingData(rawText));
        });
    }
    
});

let buffer = getDocBuffer(TEMPLATE_CONTENTS, dataArray);
let outputFileName = `${OUTPUT_FILENAME_PRIFIX}${OUTPUT_FILENAME_SUFFIX}.docx`;
fs.writeFileSync(path.resolve(OUTPUT_DIR, outputFileName), buffer);